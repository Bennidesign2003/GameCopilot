# __mcp_version__ = "3.6.5"
from __future__ import annotations

import datetime
import json as _json
import logging
import logging.handlers
import os
import platform
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Literal

import httpx

from mcp.server.fastmcp import FastMCP
import pynvml

mcp = FastMCP("nvidia-gpu")

# ---------------------------------------------------------------------------
# Logging — file-based so users can debug "the AI says ok but I see nothing"
# ---------------------------------------------------------------------------

_LOG_FILE = Path(__file__).parent / "server.log"
logger = logging.getLogger("nvidia-mcp")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    try:
        _fh = logging.handlers.RotatingFileHandler(
            _LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
        )
        _fh.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        )
        logger.addHandler(_fh)
    except Exception:
        # If we can't open the log file (e.g. permission), fall back silently.
        pass

# ---------------------------------------------------------------------------
# UserCfg.opt helpers
# ---------------------------------------------------------------------------

# MSFS 2024 is the primary target. 2024 paths come first and are preferred.
_USERCFG_CANDIDATES_2024 = [
    # MSFS 2024 Steam (confirmed location)
    Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator 2024" / "UserCfg.opt",
    Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator 2024 (Steam)" / "UserCfg.opt",
    Path(os.environ.get("APPDATA", "")) / "MicrosoftFlightSimulator2024" / "UserCfg.opt",
    # MSFS 2024 MS Store
    Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
    / "Microsoft.Limitless_8wekyb3d8bbwe" / "LocalCache" / "UserCfg.opt",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
    / "Microsoft.FlightSimulator2024_8wekyb3d8bbwe" / "LocalCache" / "UserCfg.opt",
]

_USERCFG_CANDIDATES_2020 = [
    Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator" / "UserCfg.opt",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
    / "Microsoft.FlightSimulator_8wekyb3d8bbwe" / "LocalCache" / "UserCfg.opt",
]
_USERCFG_CANDIDATES = _USERCFG_CANDIDATES_2024 + _USERCFG_CANDIDATES_2020

# Default: MSFS 2024 path (primary target)
DEFAULT_USERCFG_PATH = (
    Path(os.environ.get("APPDATA", ""))
    / "Microsoft Flight Simulator 2024"
    / "UserCfg.opt"
)


def _is_msfs_usercfg(path: Path) -> bool:
    """Check if a file looks like a real MSFS UserCfg.opt."""
    try:
        text = path.read_text(encoding="utf-8", errors="ignore")[:2000]
        # Must contain at least one MSFS-specific section
        return any(s in text for s in ("{Graphics", "{GraphicsVR", "{Video", "{Sound"))
    except Exception:
        return False


def _find_all_usercfg() -> list[dict]:
    """Find ALL UserCfg.opt files on the system with metadata.

    Returns list of dicts with path, mtime, size, and key settings.
    Searches: known candidates, APPDATA/LOCALAPPDATA glob, PowerShell
    filesystem search as last resort.
    """
    import datetime as _dt
    seen: set[str] = set()
    results: list[dict] = []

    def _add(p: Path):
        rp = str(p.resolve())
        if rp in seen:
            return
        seen.add(rp)
        if not _is_msfs_usercfg(p):
            return
        try:
            stat = p.stat()
            mtime = _dt.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            size = stat.st_size

            # Read key settings for display
            text = p.read_text(encoding="utf-8", errors="ignore")
            entries = _parse_usercfg(text)
            settings: dict[str, str] = {}
            for section, line in entries:
                if section in ("Video", "Graphics", "GraphicsVR", "RayTracing"):
                    m = re.match(r"^\s*(\S+)\s+(\S+)", line)
                    if m:
                        key = f"{section}.{m.group(1)}"
                        if key in SETTING_DEFS:
                            defn = SETTING_DEFS[key]
                            raw = m.group(2)
                            display = defn.get("values", {}).get(raw, raw)
                            settings[defn.get("label", key)] = display

            results.append({
                "path": str(p),
                "modified": mtime,
                "size_kb": round(size / 1024, 1),
                "settings": settings,
            })
        except Exception as exc:
            results.append({"path": str(p), "error": str(exc)})

    # 1. Check known candidate paths
    for c in _USERCFG_CANDIDATES:
        if c.exists():
            _add(c)

    # 2. Glob APPDATA and LOCALAPPDATA
    for env_var in ("APPDATA", "LOCALAPPDATA"):
        root = Path(os.environ.get(env_var, ""))
        if not root.exists():
            continue
        try:
            for hit in root.glob("**/UserCfg.opt"):
                _add(hit)
        except Exception:
            pass

    # 3. Glob common Steam library folders
    for steam_root in [
        r"C:\Program Files (x86)\Steam\steamapps",
        r"D:\Steam\steamapps",
        r"D:\SteamLibrary\steamapps",
        r"E:\SteamLibrary\steamapps",
    ]:
        sr = Path(steam_root)
        if sr.exists():
            try:
                for hit in sr.glob("**/UserCfg.opt"):
                    _add(hit)
            except Exception:
                pass

    # 4. PowerShell search as last resort (broader, slower)
    if not results:
        try:
            ps_script = (
                "Get-ChildItem -Path $env:APPDATA,$env:LOCALAPPDATA,"
                "'C:\\Users' -Recurse -Filter 'UserCfg.opt' "
                "-ErrorAction SilentlyContinue | "
                "Select-Object -ExpandProperty FullName"
            )
            r = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                capture_output=True, text=True, timeout=30,
            )
            for line in r.stdout.strip().splitlines():
                p = Path(line.strip())
                if p.exists():
                    _add(p)
        except Exception:
            pass

    # Sort by modification time (newest first)
    results.sort(
        key=lambda x: x.get("modified", ""),
        reverse=True,
    )
    return results


# Module-level cache for _find_usercfg (no-custom-path calls only).
# Avoids repeated filesystem scans when the same function is called many times
# in a single agent session.  Invalidated by _invalidate_usercfg_cache().
_USERCFG_CACHE: "Path | None" = None
_USERCFG_CACHE_VALID: bool = False


def _invalidate_usercfg_cache() -> None:
    """Invalidate the _find_usercfg result cache.
    Call this after any write that modifies UserCfg.opt so the next read
    re-discovers the file from disk.
    """
    global _USERCFG_CACHE, _USERCFG_CACHE_VALID
    _USERCFG_CACHE_VALID = False


def _find_usercfg(custom_path: str = "") -> Path | None:
    """Find the MSFS 2024 UserCfg.opt.

    Priority order:
      1. Custom path (if provided and exists)
      2. Known MSFS 2024 candidate paths (first match wins)
      3. Filesystem search for UserCfg.opt with '2024' in the path

    Results for the default (no custom_path) case are cached in
    _USERCFG_CACHE to avoid repeated filesystem scans.
    """
    global _USERCFG_CACHE, _USERCFG_CACHE_VALID

    if custom_path:
        p = Path(custom_path)
        if p.exists():
            return p
        return None

    # Return cached result when available
    if _USERCFG_CACHE_VALID:
        return _USERCFG_CACHE

    # 1. Check MSFS 2024 candidates — first existing one wins
    result: "Path | None" = None
    for c in _USERCFG_CANDIDATES_2024:
        if c.exists() and _is_msfs_usercfg(c):
            result = c
            break

    # 2. Filesystem search as fallback
    if result is None:
        all_configs = _find_all_usercfg()
        if all_configs:
            for cfg in all_configs:
                if "2024" in cfg["path"]:
                    result = Path(cfg["path"])
                    break
            if result is None:
                result = Path(all_configs[0]["path"])

    _USERCFG_CACHE = result
    _USERCFG_CACHE_VALID = True
    return result


@mcp.tool()
def find_msfs_config(usercfg_path: str = "") -> dict:
    """Find and show the MSFS UserCfg.opt file that will be used for settings changes.

    Call this when the user says settings aren't applying, to diagnose which file is being edited.
    Returns the active config path, all found configs, and key current settings.
    """
    active = _find_usercfg(usercfg_path)
    all_configs = _find_all_usercfg()

    result: dict = {
        "aktive_config": str(active) if active else "NICHT GEFUNDEN",
        "config_exists": active.exists() if active else False,
        "alle_gefundenen": [c["path"] for c in all_configs],
        "anzahl_gefunden": len(all_configs),
    }

    if active and active.exists():
        settings = _read_current_settings(active)
        key_settings: dict = {}
        for key in [
            "Video.DLSS", "GraphicsVR.CloudsQuality", "GraphicsVR.TerrainLoD",
            "Graphics.CloudsQuality", "Graphics.TerrainLoD", "Video.RenderScale",
        ]:
            if key in settings:
                defn = SETTING_DEFS.get(key, {})
                raw = settings[key]
                display = defn.get("values", {}).get(raw, raw)
                key_settings[key] = f"{raw} ({display})"
        result["aktuelle_einstellungen"] = key_settings

    if not active or not active.exists():
        result["fehler"] = (
            "UserCfg.opt wurde nicht gefunden! MSFS muss mindestens einmal gestartet worden sein. "
            "Starte MSFS einmal, ändere eine Einstellung in den Grafikoptionen und beende MSFS. "
            "Dann kann der Agent Einstellungen direkt ändern."
        )

    return result


# Community folder paths (MSFS 2024 only)
COMMUNITY_CANDIDATES = [
    Path(os.environ.get("APPDATA", ""))
    / "Microsoft Flight Simulator 2024"
    / "Packages"
    / "Community",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
    / "Microsoft.Limitless_8wekyb3d8bbwe" / "LocalCache"
    / "Packages" / "Community",
]

SUPPORTED_ARCHIVES = {".zip", ".7z", ".rar"}

# ---------------------------------------------------------------------------
# Driver-update helpers
# ---------------------------------------------------------------------------

NVIDIA_DRIVER_API = (
    "https://gfwsl.geforce.com/services_toolkit/services/com"
    "/nvidia/services/AjaxDriverService.php"
)

# GPU model name fragment -> (psid, pfid) for NVIDIA's driver lookup API.
# psid 101 = GeForce Desktop.  Extend as needed.
GPU_DRIVER_LOOKUP: dict[str, tuple[int, int]] = {
    # RTX 50 Series
    "RTX 5090":          (101, 985),
    "RTX 5080":          (101, 986),
    "RTX 5070 Ti":       (101, 987),
    "RTX 5070":          (101, 988),
    "RTX 5060 Ti":       (101, 989),
    "RTX 5060":          (101, 990),
    # RTX 40 Series
    "RTX 4090":          (101, 933),
    "RTX 4080 SUPER":    (101, 966),
    "RTX 4080":          (101, 934),
    "RTX 4070 Ti SUPER": (101, 967),
    "RTX 4070 Ti":       (101, 935),
    "RTX 4070 SUPER":    (101, 968),
    "RTX 4070":          (101, 936),
    "RTX 4060 Ti":       (101, 937),
    "RTX 4060":          (101, 938),
    # RTX 30 Series
    "RTX 3090 Ti":       (101, 929),
    "RTX 3090":          (101, 876),
    "RTX 3080 Ti":       (101, 904),
    "RTX 3080":          (101, 877),
    "RTX 3070 Ti":       (101, 905),
    "RTX 3070":          (101, 878),
    "RTX 3060 Ti":       (101, 899),
    "RTX 3060":          (101, 892),
    "RTX 3050":          (101, 932),
    # RTX 20 Series
    "RTX 2080 Ti":       (101, 815),
    "RTX 2080 SUPER":    (101, 852),
    "RTX 2080":          (101, 816),
    "RTX 2070 SUPER":    (101, 853),
    "RTX 2070":          (101, 817),
    "RTX 2060 SUPER":    (101, 854),
    "RTX 2060":          (101, 818),
    # GTX 16 Series
    "GTX 1660 Ti":       (101, 820),
    "GTX 1660 SUPER":    (101, 855),
    "GTX 1660":          (101, 821),
    "GTX 1650 SUPER":    (101, 856),
    "GTX 1650":          (101, 822),
    # GTX 10 Series
    "GTX 1080 Ti":       (101, 790),
    "GTX 1080":          (101, 783),
    "GTX 1070 Ti":       (101, 805),
    "GTX 1070":          (101, 784),
    "GTX 1060":          (101, 785),
    "GTX 1050 Ti":       (101, 791),
    "GTX 1050":          (101, 792),
}

# Windows OS IDs for the NVIDIA API
_OS_IDS: dict[str, int] = {
    "win10_64": 57,
    "win11_64": 57,   # same driver channel
}


def _get_gpu_info() -> tuple[str, str]:
    """Return (driver_version, gpu_name) via NVML."""
    pynvml.nvmlInit()
    try:
        driver = pynvml.nvmlSystemGetDriverVersion()
        handle = pynvml.nvmlDeviceGetHandleByIndex(0)
        name = pynvml.nvmlDeviceGetName(handle)
        if isinstance(driver, bytes):
            driver = driver.decode()
        if isinstance(name, bytes):
            name = name.decode()
        return driver, name
    finally:
        pynvml.nvmlShutdown()


def _match_gpu_model(gpu_name: str) -> tuple[int, int] | None:
    """Find the best matching (psid, pfid) for *gpu_name*.

    Matches longest key first so "RTX 4080 SUPER" wins over "RTX 4080".
    """
    upper = gpu_name.upper()
    # Sort by key length descending so more-specific names match first
    for model in sorted(GPU_DRIVER_LOOKUP, key=len, reverse=True):
        if model.upper() in upper:
            return GPU_DRIVER_LOOKUP[model]
    return None


def _lookup_pfid(gpu_name: str) -> tuple[int, int] | None:
    """Dynamically look up the (psid, pfid) for a GPU from NVIDIA's API.

    Queries the product-family list for each known product series (psid/psrid)
    and matches the GPU name.  Falls back to the static table.
    """
    # First try the static table
    static = _match_gpu_model(gpu_name)

    # Product-type IDs to try (Desktop GeForce series)
    # psid=101 (GeForce), ptid varies by generation, osID=57 (Win10/11 64-bit)
    series_ids = [
        ("101", "14"),  # RTX 50 Series
        ("101", "13"),  # RTX 40 Series
        ("101", "12"),  # RTX 30 Series
        ("101", "11"),  # RTX 20 Series / GTX 16
        ("101", "10"),  # GTX 10 Series
    ]

    upper = gpu_name.upper()
    with httpx.Client(follow_redirects=True, timeout=15) as client:
        for psid, ptid in series_ids:
            try:
                resp = client.get(
                    NVIDIA_DRIVER_API,
                    params={
                        "func": "DriverManualLookup",
                        "psid": psid,
                        "ptid": ptid,
                        "osID": "57",
                        "languageCode": "1033",
                        "numberOfResults": "0",  # just need product list
                    },
                )
                # Parse the product families from the response
                for m in re.finditer(
                    r'<LookupValue\b[^>]*ParentID="(\d+)"[^>]*>.*?'
                    r"<Name>\s*([^<]+?)\s*</Name>",
                    resp.text,
                    re.DOTALL,
                ):
                    pfid_str, product_name = m.group(1), m.group(2)
                    if product_name.upper().strip() in upper or upper in product_name.upper().strip():
                        return (int(psid), int(pfid_str))
            except Exception:
                continue

    return static


def _query_latest_driver(psid: int, pfid: int) -> dict | None:
    """Query the NVIDIA driver API and return latest driver info or None."""
    params = {
        "func": "DriverManualLookup",
        "psid": str(psid),
        "pfid": str(pfid),
        "osID": "57",          # Windows 10/11 64-bit
        "languageCode": "1033", # English
        "isWHQL": "1",
        "dch": "1",            # DCH (modern Windows driver model)
        "sort1": "0",
        "numberOfResults": "1",
    }
    with httpx.Client(follow_redirects=True, timeout=30) as client:
        resp = client.get(NVIDIA_DRIVER_API, params=params)
        resp.raise_for_status()

    # The API returns XML with driver details inside <LookupValue> elements.
    try:
        root = ET.fromstring(resp.text)
    except ET.ParseError:
        return None

    lv = root.find(".//{*}LookupValue")
    if lv is None:
        lv = root.find(".//LookupValue")
    if lv is None:
        return None

    def _text(tag: str) -> str:
        el = lv.find(tag)
        if el is None:
            el = lv.find(f"{{*}}{tag}")
        return el.text.strip() if el is not None and el.text else ""

    version = _text("Version")
    download_url = _text("DownloadURL")
    name = _text("Name")
    release_date = _text("ReleaseDateTime")

    if not version:
        return None

    return {
        "version": version,
        "name": name,
        "download_url": download_url,
        "release_date": release_date,
    }


def _version_tuple(v: str) -> tuple[int, ...]:
    """Convert '560.94' to (560, 94) for comparison."""
    return tuple(int(x) for x in re.split(r"[.\-]", v) if x.isdigit())


# -- Presets ----------------------------------------------------------------
# Each preset maps  "SectionName.Key" -> value  (value is always a string).
# Sections in UserCfg.opt look like:  {SectionName   key value   ... }

# ---------------------------------------------------------------------------
# GPU-aware VR presets for MSFS 2024 + Pimax
# ---------------------------------------------------------------------------
# Pimax headsets render at very high resolutions (per eye ~2448x2448 or higher).
# Even top-tier GPUs cannot run Ultra settings at these resolutions.
# These presets are tuned for smooth VR in Pimax — prioritizing stable 45+ FPS
# with Smart Smoothing over raw quality.
#
# GPU classification:
#   - Flagship (RTX 5090, 4090): 24 GB VRAM, can push higher settings
#   - High-end (RTX 5080, 4080, 3090): 16 GB VRAM, needs careful tuning
#   - Mid-range (RTX 4070, 3080, 3070): 8-12 GB VRAM, must be conservative
#   - Entry (RTX 4060, 3060, below): 8 GB or less, minimum settings
# ---------------------------------------------------------------------------

# Known GPU performance tiers for VR (approximate raster perf relative to 4090=100)
_GPU_VR_TIER: dict[str, str] = {
    # RTX 50 Series
    "RTX 5090": "flagship",
    "RTX 5080": "high_end",
    "RTX 5070 Ti": "mid_high",
    "RTX 5070": "mid_range",
    "RTX 5060": "entry",
    # RTX 40 Series
    "RTX 4090": "flagship",
    "RTX 4080 SUPER": "high_end",
    "RTX 4080": "high_end",
    "RTX 4070 Ti SUPER": "mid_high",
    "RTX 4070 Ti": "mid_high",
    "RTX 4070 SUPER": "mid_range",
    "RTX 4070": "mid_range",
    "RTX 4060 Ti": "entry",
    "RTX 4060": "entry",
    # RTX 30 Series
    "RTX 3090 Ti": "high_end",
    "RTX 3090": "high_end",
    "RTX 3080 Ti": "mid_high",
    "RTX 3080": "mid_high",
    "RTX 3070 Ti": "mid_range",
    "RTX 3070": "mid_range",
    "RTX 3060 Ti": "entry",
    "RTX 3060": "entry",
}


def _detect_gpu_tier() -> tuple[str, str, int]:
    """Detect GPU and return (gpu_name, tier, vram_mb).

    Tier is one of: flagship, high_end, mid_high, mid_range, entry, unknown.
    """
    try:
        pynvml.nvmlInit()
        try:
            handle = pynvml.nvmlDeviceGetHandleByIndex(0)
            name = pynvml.nvmlDeviceGetName(handle)
            mem = pynvml.nvmlDeviceGetMemoryInfo(handle)
        finally:
            pynvml.nvmlShutdown()

        if isinstance(name, bytes):
            name = name.decode()
        vram_mb = round(mem.total / 1024 / 1024)

        # Match GPU name against known tiers
        name_upper = name.upper()
        tier = "unknown"
        for gpu_key, gpu_tier in _GPU_VR_TIER.items():
            if gpu_key.upper() in name_upper:
                tier = gpu_tier
                break

        # Fallback by VRAM if not matched
        if tier == "unknown":
            vram_gb = vram_mb / 1024
            if vram_gb >= 20:
                tier = "flagship"
            elif vram_gb >= 14:
                tier = "high_end"
            elif vram_gb >= 10:
                tier = "mid_high"
            elif vram_gb >= 8:
                tier = "mid_range"
            else:
                tier = "entry"

        return name, tier, vram_mb
    except Exception:
        return "Unknown", "unknown", 0


# GPU-tier-specific MSFS 2024 presets for Pimax VR
# Values tuned from real-world testing at Pimax-level resolutions
_VR_PRESETS_BY_TIER: dict[str, dict[str, str]] = {
    "flagship": {
        # RTX 5090 / 4090 — can push more but not Ultra in Pimax VR
        "Video.DLSS": "3",              # Balanced (not Quality — Pimax res is huge)
        "Video.DLSSG": "1",             # Frame Gen On
        "GraphicsVR.TerrainLoD": "2.000000",
        "GraphicsVR.ObjectsLoD": "2.000000",
        "GraphicsVR.CloudsQuality": "2",    # High
        "GraphicsVR.AnisotropicFilter": "4", # 16x
        "GraphicsVR.SSContact": "0",         # SSAO Off (too expensive in VR)
        "GraphicsVR.Reflections": "1",       # Low
        "GraphicsVR.TextureResolution": "3", # Ultra
        "GraphicsVR.MotionBlur": "0",
        "GraphicsVR.ShadowQuality": "2",     # High
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
    "high_end": {
        # RTX 5080 / 4080 / 3090 — good but must be conservative in Pimax VR
        "Video.DLSS": "4",              # Performance (MUST for Pimax resolution)
        "Video.DLSSG": "1",             # Frame Gen On
        "GraphicsVR.TerrainLoD": "1.500000",
        "GraphicsVR.ObjectsLoD": "1.500000",
        "GraphicsVR.CloudsQuality": "1",    # Medium
        "GraphicsVR.AnisotropicFilter": "3", # 8x
        "GraphicsVR.SSContact": "0",         # Off
        "GraphicsVR.Reflections": "0",       # Off
        "GraphicsVR.TextureResolution": "2", # High
        "GraphicsVR.MotionBlur": "0",
        "GraphicsVR.ShadowQuality": "1",     # Medium
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
    "mid_high": {
        # RTX 4070 Ti / 3080 — limited in Pimax VR
        "Video.DLSS": "4",              # Performance
        "Video.DLSSG": "1",
        "GraphicsVR.TerrainLoD": "1.000000",
        "GraphicsVR.ObjectsLoD": "1.000000",
        "GraphicsVR.CloudsQuality": "1",    # Medium
        "GraphicsVR.AnisotropicFilter": "2", # 4x
        "GraphicsVR.SSContact": "0",
        "GraphicsVR.Reflections": "0",
        "GraphicsVR.TextureResolution": "2", # High
        "GraphicsVR.MotionBlur": "0",
        "GraphicsVR.ShadowQuality": "1",     # Medium
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
    "mid_range": {
        # RTX 4070 / 3070 — tight in Pimax VR
        "Video.DLSS": "5",              # Ultra Performance
        "Video.DLSSG": "1",
        "GraphicsVR.TerrainLoD": "1.000000",
        "GraphicsVR.ObjectsLoD": "1.000000",
        "GraphicsVR.CloudsQuality": "0",    # Low
        "GraphicsVR.AnisotropicFilter": "1", # 2x
        "GraphicsVR.SSContact": "0",
        "GraphicsVR.Reflections": "0",
        "GraphicsVR.TextureResolution": "1", # Medium
        "GraphicsVR.MotionBlur": "0",
        "GraphicsVR.ShadowQuality": "0",     # Low
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
    "entry": {
        # RTX 4060 / 3060 — minimum for VR
        "Video.DLSS": "5",              # Ultra Performance
        "Video.DLSSG": "1",
        "GraphicsVR.TerrainLoD": "0.500000",
        "GraphicsVR.ObjectsLoD": "0.500000",
        "GraphicsVR.CloudsQuality": "0",    # Low
        "GraphicsVR.AnisotropicFilter": "1", # 2x
        "GraphicsVR.SSContact": "0",
        "GraphicsVR.Reflections": "0",
        "GraphicsVR.TextureResolution": "0", # Low
        "GraphicsVR.MotionBlur": "0",
        "GraphicsVR.ShadowQuality": "0",     # Low
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
}

# Desktop presets are more forgiving (single-panel, lower resolution)
_DESKTOP_PRESETS_BY_TIER: dict[str, dict[str, str]] = {
    "flagship": {
        "Video.DLSS": "1",              # DLAA
        "Video.DLSSG": "1",
        "Graphics.TerrainLoD": "3.000000",
        "Graphics.ObjectsLoD": "3.000000",
        "Graphics.CloudsQuality": "3",     # Ultra
        "Graphics.AnisotropicFilter": "4",  # 16x
        "Graphics.SSContact": "1",
        "Graphics.Reflections": "3",        # High
        "Graphics.TextureResolution": "3",  # Ultra
        "Graphics.ShadowQuality": "3",      # Ultra
        "RayTracing.Enabled": "1",
        "RayTracing.Reflections": "1",
    },
    "high_end": {
        "Video.DLSS": "2",              # Quality
        "Video.DLSSG": "1",
        "Graphics.TerrainLoD": "2.500000",
        "Graphics.ObjectsLoD": "2.500000",
        "Graphics.CloudsQuality": "2",     # High
        "Graphics.AnisotropicFilter": "4",
        "Graphics.SSContact": "1",
        "Graphics.Reflections": "2",        # Medium
        "Graphics.TextureResolution": "3",  # Ultra
        "Graphics.ShadowQuality": "2",      # High
        "RayTracing.Enabled": "1",
        "RayTracing.Reflections": "1",
    },
    "mid_high": {
        "Video.DLSS": "3",              # Balanced
        "Video.DLSSG": "1",
        "Graphics.TerrainLoD": "2.000000",
        "Graphics.ObjectsLoD": "2.000000",
        "Graphics.CloudsQuality": "2",
        "Graphics.AnisotropicFilter": "3",
        "Graphics.SSContact": "1",
        "Graphics.Reflections": "1",
        "Graphics.TextureResolution": "2",
        "Graphics.ShadowQuality": "2",
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
    "mid_range": {
        "Video.DLSS": "4",              # Performance
        "Video.DLSSG": "1",
        "Graphics.TerrainLoD": "1.500000",
        "Graphics.ObjectsLoD": "1.500000",
        "Graphics.CloudsQuality": "1",
        "Graphics.AnisotropicFilter": "2",
        "Graphics.SSContact": "0",
        "Graphics.Reflections": "0",
        "Graphics.TextureResolution": "2",
        "Graphics.ShadowQuality": "1",
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
    "entry": {
        "Video.DLSS": "4",
        "Video.DLSSG": "1",
        "Graphics.TerrainLoD": "1.000000",
        "Graphics.ObjectsLoD": "1.000000",
        "Graphics.CloudsQuality": "0",
        "Graphics.AnisotropicFilter": "1",
        "Graphics.SSContact": "0",
        "Graphics.Reflections": "0",
        "Graphics.TextureResolution": "1",
        "Graphics.ShadowQuality": "0",
        "RayTracing.Enabled": "0",
        "RayTracing.Reflections": "0",
    },
}

# Legacy alias — existing code references PRESETS["VR"] / PRESETS["Desktop"]
PRESETS: dict[str, dict[str, str]] = {
    "VR": _VR_PRESETS_BY_TIER["high_end"],
    "Desktop": _DESKTOP_PRESETS_BY_TIER["high_end"],
}


def _parse_usercfg(text: str) -> list[tuple[str | None, str]]:
    """Parse UserCfg.opt into a list of (section_name | None, raw_line) tuples.

    Section headers look like ``{SectionName`` and end with ``}``.
    We keep every line verbatim so we can reconstruct the file losslessly.
    """
    entries: list[tuple[str | None, str]] = []
    current_section: str | None = None
    for line in text.splitlines(keepends=True):
        stripped = line.strip()
        if stripped.startswith("{") and not stripped.startswith("}"):
            current_section = stripped.lstrip("{").strip()
            entries.append((None, line))  # section header is not inside a section
        elif stripped == "}":
            entries.append((None, line))
            current_section = None
        else:
            entries.append((current_section, line))
    return entries


def _apply_overrides(
    entries: list[tuple[str | None, str]],
    overrides: dict[str, str],
) -> tuple[list[tuple[str | None, str]], set[str]]:
    """Return a new entries list with the requested key-value overrides applied."""
    # Build a lookup:  section -> key -> new_value
    lookup: dict[str, dict[str, str]] = {}
    for compound_key, value in overrides.items():
        section, key = compound_key.split(".", 1)
        lookup.setdefault(section, {})[key] = value

    applied: set[str] = set()
    new_entries = []
    for section, line in entries:
        if section and section in lookup:
            # Try to match "  Key   Value" pattern
            m = re.match(r"^(\s*)(\S+)(\s+)(\S+)(.*)", line)
            if m:
                key = m.group(2)
                if key in lookup[section]:
                    indent, _, spacing, _, rest = m.groups()
                    new_line = f"{indent}{key}{spacing}{lookup[section][key]}{rest}"
                    if not new_line.endswith(("\n", "\r\n")) and line.endswith(("\n", "\r\n")):
                        new_line += "\n"
                    new_entries.append((section, new_line))
                    applied.add(f"{section}.{key}")
                    continue
        new_entries.append((section, line))

    not_applied = set(overrides.keys()) - applied
    return new_entries, not_applied


def _entries_to_text(entries: list[tuple[str | None, str]]) -> str:
    return "".join(line for _, line in entries)


@mcp.tool()
def get_gpu_status(gpu_index: int = 0) -> dict:
    """Read temperature, utilization and VRAM of an NVIDIA GPU.

    Args:
        gpu_index: Index of the GPU (default 0 for the first GPU).
    """
    if gpu_index < 0:
        return {"error": "gpu_index must be 0 or greater."}
    try:
        pynvml.nvmlInit()
    except Exception as exc:
        return {"error": f"NVML init failed (no NVIDIA GPU or driver issue?): {exc}"}
    try:
        handle = pynvml.nvmlDeviceGetHandleByIndex(gpu_index)
        name = pynvml.nvmlDeviceGetName(handle)
        if isinstance(name, bytes):
            name = name.decode()
        temp = pynvml.nvmlDeviceGetTemperature(handle, pynvml.NVML_TEMPERATURE_GPU)
        util = pynvml.nvmlDeviceGetUtilizationRates(handle)
        mem = pynvml.nvmlDeviceGetMemoryInfo(handle)

        return {
            "gpu_name": name,
            "temperature_c": temp,
            "gpu_utilization_percent": util.gpu,
            "memory_utilization_percent": util.memory,
            "vram_total_mb": round(mem.total / 1024 / 1024),
            "vram_used_mb": round(mem.used / 1024 / 1024),
            "vram_free_mb": round(mem.free / 1024 / 1024),
        }
    except Exception as exc:
        return {"error": f"GPU query failed (index {gpu_index}): {exc}"}
    finally:
        pynvml.nvmlShutdown()


# ---------------------------------------------------------------------------
# MSFS graphics version history
# ---------------------------------------------------------------------------

# History is stored as a JSON file next to UserCfg.opt:
#   ...Microsoft Flight Simulator 2024/UserCfg_history.json
# Structure: [ { "timestamp": ..., "label": ..., "content": "full file" }, ... ]

def _history_path(cfg_path: Path) -> Path:
    return cfg_path.parent / "UserCfg_history.json"


def _load_history(cfg_path: Path) -> list[dict]:
    hp = _history_path(cfg_path)
    if hp.exists():
        try:
            return _json.loads(hp.read_text(encoding="utf-8"))
        except Exception:
            # Corrupted history file — start fresh rather than crashing _snapshot.
            logger.warning("History file %s is corrupted; resetting.", hp)
            return []
    return []


def _save_history(cfg_path: Path, history: list[dict]) -> None:
    hp = _history_path(cfg_path)
    hp.write_text(_json.dumps(history, indent=2, ensure_ascii=False), encoding="utf-8")


def _backup_dir(cfg_path: Path) -> Path:
    """Return (and create) the backup folder next to UserCfg.opt."""
    d = cfg_path.parent / "UserCfg_backups"
    d.mkdir(exist_ok=True)
    return d


def _snapshot(cfg_path: Path, label: str) -> int:
    """Save the current UserCfg.opt as a new history entry AND a .opt file.

    Returns the version number.  Each snapshot produces:
      1. A JSON entry in UserCfg_history.json  (for restore_msfs_graphics)
      2. A timestamped copy in UserCfg_backups/ (for manual recovery)
    """
    content = cfg_path.read_text(encoding="utf-8")
    now = datetime.datetime.now()

    # --- JSON history ---
    history = _load_history(cfg_path)
    version = len(history) + 1
    history.append({
        "version": version,
        "timestamp": now.isoformat(timespec="seconds"),
        "label": label,
        "content": content,
    })
    _save_history(cfg_path, history)

    # --- File backup ---
    ts = now.strftime("%Y%m%d_%H%M%S")
    backup_file = _backup_dir(cfg_path) / f"UserCfg_v{version}_{ts}.opt"
    backup_file.write_text(content, encoding="utf-8")

    # Also keep a simple "latest" backup that's always the most recent state
    latest = _backup_dir(cfg_path) / "UserCfg_latest.opt"
    latest.write_text(content, encoding="utf-8")

    return version


# ---------------------------------------------------------------------------
# MSFS graphics analysis & recommendations
# ---------------------------------------------------------------------------

# Which settings to read from which section, with human-readable names & value maps
SETTING_DEFS: dict[str, dict] = {
    # ---- Video section ----
    "Video.DLSS": {
        "label": "DLSS Mode",
        "values": {"0": "Off", "1": "DLAA", "2": "Quality", "3": "Balanced",
                   "4": "Performance", "5": "Ultra Performance"},
    },
    "Video.DLSSG": {
        "label": "DLSS Frame Generation",
        "values": {"0": "Off", "1": "On"},
    },
    "Video.RenderScale": {
        "label": "Render Scale",
    },
    # ---- Graphics (Desktop) ----
    "Graphics.TerrainLoD": {"label": "Terrain LoD (Desktop)"},
    "Graphics.ObjectsLoD": {"label": "Objects LoD (Desktop)"},
    "Graphics.CloudsQuality": {
        "label": "Clouds Quality (Desktop)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    "Graphics.AnisotropicFilter": {
        "label": "Anisotropic Filter (Desktop)",
        "values": {"1": "2x", "2": "4x", "3": "8x", "4": "16x"},
    },
    "Graphics.SSContact": {
        "label": "SSAO (Desktop)",
        "values": {"0": "Off", "1": "On"},
    },
    "Graphics.Reflections": {
        "label": "SSR Reflections (Desktop)",
        "values": {"0": "Off", "1": "Low", "2": "Medium", "3": "High", "4": "Ultra"},
    },
    "Graphics.TextureResolution": {
        "label": "Texture Resolution (Desktop)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    "Graphics.MotionBlur": {
        "label": "Motion Blur (Desktop)",
        "values": {"0": "Off", "1": "On"},
    },
    "Graphics.DepthOfField": {
        "label": "Depth of Field (Desktop)",
        "values": {"0": "Off", "1": "On"},
    },
    "Graphics.Buildings": {
        "label": "Buildings Quality (Desktop)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    "Graphics.Trees": {
        "label": "Trees Quality (Desktop)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    "Graphics.GrassAndBushes": {
        "label": "Grass & Bushes (Desktop)",
        "values": {"0": "Off", "1": "Low", "2": "Medium", "3": "High", "4": "Ultra"},
    },
    "Graphics.ShadowQuality": {
        "label": "Shadow Quality (Desktop)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    # ---- GraphicsVR ----
    "GraphicsVR.TerrainLoD": {"label": "Terrain LoD (VR)"},
    "GraphicsVR.ObjectsLoD": {"label": "Objects LoD (VR)"},
    "GraphicsVR.CloudsQuality": {
        "label": "Clouds Quality (VR)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    "GraphicsVR.AnisotropicFilter": {
        "label": "Anisotropic Filter (VR)",
        "values": {"1": "2x", "2": "4x", "3": "8x", "4": "16x"},
    },
    "GraphicsVR.SSContact": {
        "label": "SSAO (VR)",
        "values": {"0": "Off", "1": "On"},
    },
    "GraphicsVR.Reflections": {
        "label": "SSR Reflections (VR)",
        "values": {"0": "Off", "1": "Low", "2": "Medium", "3": "High", "4": "Ultra"},
    },
    "GraphicsVR.TextureResolution": {
        "label": "Texture Resolution (VR)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    "GraphicsVR.MotionBlur": {
        "label": "Motion Blur (VR)",
        "values": {"0": "Off", "1": "On"},
    },
    "GraphicsVR.ShadowQuality": {
        "label": "Shadow Quality (VR)",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    # ---- RayTracing ----
    "RayTracing.Enabled": {
        "label": "Ray Tracing",
        "values": {"0": "Off", "1": "On"},
    },
    "RayTracing.Reflections": {
        "label": "RT Reflections",
        "values": {"0": "Off", "1": "On"},
    },
}

# GPU VRAM tiers → what the card can reasonably handle
# (vram_gb_min, tier_name, recommendations)
VRAM_TIERS = [
    (16, "high_end",  "Ultra-Texturen, hohe LoDs, RayTracing und DLSS Quality sind möglich."),
    (10, "mid_high",  "Hohe Texturen, mittlere-hohe LoDs. RayTracing nur mit DLSS Balanced/Performance."),
    (8,  "mid",       "Mittlere Texturen, LoD ~1.5. RayTracing besser aus, DLSS Performance empfohlen."),
    (6,  "low_mid",   "Mittlere Texturen, LoD ~1.0. RayTracing aus, DLSS Performance/Ultra Performance."),
    (0,  "low",       "Niedrige Texturen, minimale LoDs, DLSS Ultra Performance, alles Extras aus."),
]


# Sections we care about when reading all settings
_RELEVANT_SECTIONS = {"Graphics", "GraphicsVR", "Video", "RayTracing", "Misc"}


def _read_current_settings(cfg_path: Path) -> dict[str, str]:
    """Read ALL settings from graphics-relevant sections in UserCfg.opt.

    Returns Section.Key -> raw_value for every key in Graphics, GraphicsVR,
    Video, RayTracing, and Misc sections.
    """
    text = cfg_path.read_text(encoding="utf-8")
    entries = _parse_usercfg(text)
    result = {}
    for section, line in entries:
        if not section or section not in _RELEVANT_SECTIONS:
            continue
        m = re.match(r"^\s*(\S+)\s+(\S+)", line)
        if m:
            key, value = m.group(1), m.group(2)
            result[f"{section}.{key}"] = value
    return result


def _build_recommendations(
    current: dict[str, str], vram_mb: int, gpu_name: str
) -> list[dict]:
    """Generate optimization tips based on current settings and GPU."""
    vram_gb = vram_mb / 1024
    tier_name = "low"
    tier_desc = ""
    for min_gb, name, desc in VRAM_TIERS:
        if vram_gb >= min_gb:
            tier_name, tier_desc = name, desc
            break

    tips: list[dict] = []

    def _tip(key: str, suggestion: str, reason: str):
        tips.append({"setting": SETTING_DEFS[key]["label"], "key": key,
                      "current": current.get(key, "?"), "suggestion": suggestion,
                      "reason": reason})

    # --- DLSS ---
    dlss = current.get("Video.DLSS", "")
    if dlss == "" or dlss == "0":
        if vram_gb < 24:
            _tip("Video.DLSS", "2 (Quality) oder 3 (Balanced)",
                 "DLSS bringt massiv FPS ohne sichtbaren Qualitätsverlust.")
    elif dlss in ("4", "5") and tier_name == "high_end":
        _tip("Video.DLSS", "2 (Quality) oder 1 (DLAA)",
             "Deine GPU ist stark genug für DLSS Quality statt Performance.")

    dlssg = current.get("Video.DLSSG", "")
    if dlssg in ("", "0"):
        if "40" in gpu_name or "50" in gpu_name:
            _tip("Video.DLSSG", "1 (An)",
                 "DLSS Frame Generation verdoppelt die gefühlten FPS auf RTX 40/50.")

    # --- RayTracing ---
    rt = current.get("RayTracing.Enabled", "")
    if rt == "1" and tier_name in ("low", "low_mid", "mid"):
        _tip("RayTracing.Enabled", "0 (Aus)",
             f"RayTracing kostet viel Performance. Bei {vram_gb:.0f} GB VRAM besser deaktivieren.")
    elif rt == "0" and tier_name == "high_end":
        _tip("RayTracing.Enabled", "1 (An)",
             "Deine GPU kann RayTracing in Kombination mit DLSS gut handeln.")

    # --- Terrain/Objects LoD ---
    for prefix in ("Graphics", "GraphicsVR"):
        label = "Desktop" if prefix == "Graphics" else "VR"
        tlod_key = f"{prefix}.TerrainLoD"
        olod_key = f"{prefix}.ObjectsLoD"
        tlod_str = current.get(tlod_key, "")
        olod_str = current.get(olod_key, "")
        if not tlod_str:
            continue
        tlod = float(tlod_str)
        olod = float(olod_str) if olod_str else 2.0

        if tier_name == "high_end":
            if prefix == "Graphics" and tlod < 2.0:
                _tip(tlod_key, "2.5 – 4.0",
                     f"Deine GPU kann höhere Terrain LoD ({label}) problemlos handeln.")
            elif prefix == "GraphicsVR" and tlod < 1.0:
                _tip(tlod_key, "1.0 – 2.0",
                     "In VR kann die LoD etwas höher — deine GPU schafft das.")
        elif tier_name == "mid_high":
            if prefix == "Graphics" and tlod < 1.5:
                _tip(tlod_key, "1.5 – 2.5",
                     f"Deine GPU kann moderate Terrain LoD ({label}) gut handeln.")
            elif prefix == "GraphicsVR" and tlod < 1.0:
                _tip(tlod_key, "1.0 – 1.5",
                     "In VR ist bis LoD 1.5 mit deiner GPU gut machbar.")
            if tlod > 3.0:
                _tip(tlod_key, "2.0 – 2.5",
                     f"Terrain LoD ({label}) über 3.0 kann selbst bei 10–16 GB VRAM zu Rucklern führen.")
        elif tier_name in ("mid", "low_mid"):
            if tlod > 2.0:
                _tip(tlod_key, "1.0 – 1.5",
                     f"Terrain LoD ({label}) über 2.0 frisst viel CPU/GPU.")
            if olod > 2.0:
                _tip(olod_key, "1.0 – 1.5",
                     f"Objects LoD ({label}) über 2.0 ist bei deinem VRAM zu hoch.")

    # --- VR specific ---
    vr_ssao = current.get("GraphicsVR.SSContact", "0")
    vr_refl = current.get("GraphicsVR.Reflections", "0")
    vr_mb = current.get("GraphicsVR.MotionBlur", "0")
    if "VR" in gpu_name.upper() or any(k.startswith("GraphicsVR") for k in current):
        if vr_ssao == "1":
            _tip("GraphicsVR.SSContact", "0 (Aus)",
                 "SSAO in VR kostet FPS und ist im Headset kaum sichtbar.")
        if vr_refl not in ("0", ""):
            _tip("GraphicsVR.Reflections", "0 (Aus)",
                 "SSR Reflections in VR sind performance-intensiv bei wenig Nutzen.")
        if vr_mb == "1":
            _tip("GraphicsVR.MotionBlur", "0 (Aus)",
                 "Motion Blur in VR verursacht Übelkeit — immer ausschalten.")

    # --- Desktop specific ---
    dof = current.get("Graphics.DepthOfField", "")
    if dof == "1" and tier_name not in ("high_end",):
        _tip("Graphics.DepthOfField", "0 (Aus)",
             "Depth of Field kostet FPS und wird von vielen Spielern als störend empfunden.")

    shadow = current.get("Graphics.ShadowQuality", "")
    if shadow == "3" and tier_name in ("mid", "low_mid", "low"):
        _tip("Graphics.ShadowQuality", "2 (High)",
             "Ultra-Schatten sind kaum von High zu unterscheiden, kosten aber deutlich mehr.")
    elif shadow in ("0", "1") and tier_name == "high_end":
        _tip("Graphics.ShadowQuality", "2 (High) oder 3 (Ultra)",
             "Deine GPU schafft höhere Schattenqualität problemlos.")

    clouds = current.get("Graphics.CloudsQuality", "")
    if clouds == "3" and tier_name in ("low_mid", "low"):
        _tip("Graphics.CloudsQuality", "2 (High)",
             "Ultra-Wolken sind sehr GPU-intensiv. High reicht optisch fast immer.")
    elif clouds in ("0", "1") and tier_name in ("high_end", "mid_high"):
        _tip("Graphics.CloudsQuality", "2 (High) oder 3 (Ultra)",
             "Deine GPU kann höhere Wolkenqualität — großer visueller Gewinn.")

    # Texture Resolution
    tex = current.get("Graphics.TextureResolution", "")
    if tex in ("0", "1") and tier_name in ("high_end", "mid_high"):
        _tip("Graphics.TextureResolution", "2 (High) oder 3 (Ultra)",
             "Bei 16 GB VRAM solltest du hohe Texturen nutzen — kein FPS-Verlust.")
    tex_vr = current.get("GraphicsVR.TextureResolution", "")
    if tex_vr in ("0", "1") and tier_name in ("high_end", "mid_high"):
        _tip("GraphicsVR.TextureResolution", "2 (High)",
             "VR-Texturen auf High kosten kaum Performance bei deinem VRAM.")

    # Anisotropic Filter
    af = current.get("Graphics.AnisotropicFilter", "")
    if af in ("1", "2") and tier_name in ("high_end", "mid_high"):
        _tip("Graphics.AnisotropicFilter", "4 (16x)",
             "Anisotroper Filter auf 16x kostet quasi keine Performance, sieht aber besser aus.")
    af_vr = current.get("GraphicsVR.AnisotropicFilter", "")
    if af_vr in ("1", "2") and tier_name in ("high_end", "mid_high"):
        _tip("GraphicsVR.AnisotropicFilter", "3 (8x) oder 4 (16x)",
             "Höherer AF in VR verbessert Texturen in der Ferne deutlich.")

    return tips


def _human_label(compound_key: str) -> str:
    """Return a human-readable label for a Section.Key compound key."""
    if compound_key in SETTING_DEFS:
        return SETTING_DEFS[compound_key]["label"]
    # Auto-generate: "Graphics.TerrainLoD" → "TerrainLoD"
    _, _, key = compound_key.partition(".")
    # CamelCase to spaced: "TerrainLoD" → "Terrain LoD"
    spaced = re.sub(r"([a-z])([A-Z])", r"\1 \2", key)
    return spaced


def _human_value(compound_key: str, raw: str) -> str:
    """Convert a raw value to its human-readable form if a mapping exists."""
    defn = SETTING_DEFS.get(compound_key)
    if defn and "values" in defn:
        return defn["values"].get(raw, raw)
    return raw


def _build_settings_table(
    current: dict[str, str],
    tips: list[dict],
    section_prefix: str,
) -> list[dict]:
    """Build a table of ALL settings for a given section prefix.

    Each row: Setting | Current Value | Status | Recommendation | Reason
    """
    tip_by_key = {t["key"]: t for t in tips}

    rows: list[dict] = []
    for key, raw in current.items():
        # Match prefix exactly: "Graphics." should not match "GraphicsVR."
        section = key.split(".")[0] + "."
        if section != section_prefix:
            continue

        label = _human_label(key)
        display = _human_value(key, raw)
        tip = tip_by_key.get(key)

        rows.append({
            "setting": label,
            "current": display,
            "status": "CHANGE" if tip else "OK",
            "recommendation": tip["suggestion"] if tip else "-",
            "reason": tip["reason"] if tip else "",
        })
    return rows


@mcp.tool()
def diagnose_msfs_config() -> dict:
    """Find ALL MSFS UserCfg.opt files on this PC and show what's inside each.

    Use this FIRST when MSFS settings seem wrong or changes aren't visible.
    It answers the critical question: which config file is MSFS 2024 actually using?

    Shows for each found file:
    - Full path
    - Last modification time (newest = actively used by MSFS)
    - Key settings (DLSS, Clouds, Terrain LoD, etc.)

    Use the path from the newest file as usercfg_path in other tools if
    auto-detection picks the wrong one.

    Verwende dieses Tool wenn:
    - "DLSS ist immer noch auf TAA obwohl ich es geändert habe"
    - "Einstellungen werden nicht übernommen"
    - "Wo speichert MSFS 2024 die Config?"
    """
    all_configs = _find_all_usercfg()

    if not all_configs:
        # Try harder with PowerShell on all drives
        try:
            ps_script = (
                "$drives = (Get-PSDrive -PSProvider FileSystem).Root; "
                "foreach($d in $drives) { "
                "  Get-ChildItem -Path $d -Recurse -Filter 'UserCfg.opt' "
                "  -ErrorAction SilentlyContinue | "
                "  Select-Object -ExpandProperty FullName "
                "}"
            )
            r = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                capture_output=True, text=True, timeout=120,
            )
            if r.stdout.strip():
                return {
                    "status": "deep_search",
                    "gefunden_aber_nicht_validiert": r.stdout.strip().splitlines(),
                    "hinweis": (
                        "Diese Dateien wurden gefunden aber nicht als MSFS-Config erkannt. "
                        "Bitte gib den korrekten Pfad manuell über usercfg_path an."
                    ),
                }
        except Exception:
            pass

        return {
            "error": "Keine UserCfg.opt gefunden!",
            "gesucht": [str(p) for p in _USERCFG_CANDIDATES],
            "tipp": (
                "MSFS 2024 scheint keine Config-Datei zu haben. "
                "Starte MSFS 2024 einmal manuell und beende es — "
                "dabei wird UserCfg.opt erstellt. Dann nochmal versuchen."
            ),
        }

    # Mark which one we would use by default
    active_path = all_configs[0]["path"] if all_configs else None

    return {
        "status": "ok",
        "anzahl_gefunden": len(all_configs),
        "aktive_config": active_path,
        "erklärung": (
            "Die NEUESTE Datei (oben) ist die, die MSFS aktuell nutzt. "
            "Falls das die falsche ist, gib den korrekten Pfad über "
            "usercfg_path an (z.B. set_msfs_setting(..., usercfg_path='...'))"
        ),
        "configs": all_configs,
    }


@mcp.tool()
def analyze_msfs_graphics(
    usercfg_path: str = "",
) -> dict:
    """Analyze current MSFS graphics settings with detailed tables and recommendations.

    Reads UserCfg.opt, queries GPU info (model, VRAM), and returns the results
    as separate tables for Desktop, VR and Video/RayTracing settings.
    Each row shows: Setting | Current Value | Recommendation | Reason.

    Present the results to the user as formatted markdown tables, grouped by
    category (Video/RayTracing, Desktop-Grafik, VR-Grafik).  Mark settings
    that need changes clearly.

    Args:
        usercfg_path: Full path to UserCfg.opt (leave empty for default).
    """
    cfg_path = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH

    if not cfg_path.exists():
        return {"error": f"UserCfg.opt not found at {cfg_path}"}

    # Read current settings
    current = _read_current_settings(cfg_path)

    # Get GPU info via the shared tier detection
    gpu_name, gpu_tier, vram_mb = _detect_gpu_tier()
    vram_gb = vram_mb / 1024

    # Extra GPU stats (temp, utilization)
    gpu_temp = 0
    gpu_util = 0
    try:
        pynvml.nvmlInit()
        try:
            handle = pynvml.nvmlDeviceGetHandleByIndex(0)
            gpu_temp = pynvml.nvmlDeviceGetTemperature(handle, pynvml.NVML_TEMPERATURE_GPU)
            gpu_util = pynvml.nvmlDeviceGetUtilizationRates(handle).gpu
        finally:
            pynvml.nvmlShutdown()
    except Exception:
        pass

    # Map GPU tier to old tier_name for _build_recommendations compatibility
    _TIER_TO_LEGACY = {
        "flagship": "high_end", "high_end": "high_end",
        "mid_high": "mid_high", "mid_range": "mid",
        "entry": "low_mid", "unknown": "mid",
    }
    tier_name = _TIER_TO_LEGACY.get(gpu_tier, "mid")
    tier_desc = f"GPU-Tier: {gpu_tier} ({gpu_name})"

    # Build recommendations
    tips = _build_recommendations(current, vram_mb, gpu_name)

    # Build separate tables per category
    video_rt_table = _build_settings_table(current, tips, "Video.") + \
                     _build_settings_table(current, tips, "RayTracing.")
    desktop_table = _build_settings_table(current, tips, "Graphics.")
    vr_table = _build_settings_table(current, tips, "GraphicsVR.")

    changes_needed = sum(1 for t in tips)

    return {
        "gpu_info": {
            "name": gpu_name,
            "vram_gb": round(vram_gb, 1),
            "temperature_c": gpu_temp,
            "utilization_percent": gpu_util,
            "gpu_tier": gpu_tier,
            "performance_tier": tier_name,
            "tier_description": tier_desc,
        },
        "video_and_raytracing": video_rt_table,
        "desktop_graphics": desktop_table,
        "vr_graphics": vr_table,
        "summary": {
            "total_settings_analyzed": len(current),
            "changes_recommended": changes_needed,
            "verdict": (
                "Alles optimal konfiguriert!"
                if changes_needed == 0
                else f"{changes_needed} Einstellung(en) sollten angepasst werden."
            ),
        },
        "display_instructions": (
            "IMPORTANT: You MUST display the results as THREE separate markdown tables "
            "(one per category). Use this EXACT format for each table:\n\n"
            "## Video & RayTracing\n"
            "| Einstellung | Aktuell | Status | Empfehlung | Grund |\n"
            "|---|---|---|---|---|\n"
            "| DLSS Mode | Quality | OK | - | |\n"
            "| ... | ... | ... | ... | ... |\n\n"
            "## Desktop-Grafik\n"
            "(same table format)\n\n"
            "## VR-Grafik\n"
            "(same table format)\n\n"
            "Rules:\n"
            "- Show ALL rows from each table array, not just the ones that need changes\n"
            "- For 'status' column: use a green checkmark for OK, red X for CHANGE\n"
            "- Show EVERY setting, even if it's OK\n"
            "- At the end show the summary verdict\n"
            "- Answer in German"
        ),
    }


@mcp.tool()
def optimize_msfs_graphics(
    mode: Literal["VR", "Desktop"],
    usercfg_path: str = "",
    custom_overrides: dict[str, str] | None = None,
    auto_restart: bool = True,
    dry_run: bool = False,
) -> dict:
    """Optimize MSFS 2024 graphics — automatically tuned for your GPU.

    Detects your GPU (e.g. RTX 5080 → "high_end" tier) and applies a preset
    specifically tuned for that GPU + Pimax VR resolution. This prevents
    settings that are too aggressive (e.g. Ultra on an RTX 5080 in Pimax VR).

    GPU Tiers: flagship (5090/4090), high_end (5080/4080/3090),
    mid_high (4070Ti/3080), mid_range (4070/3070), entry (4060/3060).

    Saves current settings as a snapshot BEFORE changes. Use
    restore_msfs_graphics to roll back.

    WICHTIG: MSFS muss neu gestartet werden damit Änderungen sichtbar werden!
    Bei auto_restart=True (Standard) wird MSFS automatisch beendet und neu
    gestartet.

    Args:
        mode: 'VR' (Pimax-optimiert, konservativ) oder 'Desktop' (höhere Qualität).
        usercfg_path: Pfad zu UserCfg.opt (leer = auto-detect MSFS 2024).
        custom_overrides: Zusätzliche Overrides {"Section.Key": "value", ...}
                          werden NACH dem GPU-Preset angewendet.
        auto_restart: MSFS automatisch neu starten (Standard: True).
        dry_run: Wenn True, wird nur zurückgegeben WAS sich ändern würde —
                 keine Datei wird angefasst, MSFS wird nicht beendet.
    """
    import time

    cfg_path = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH

    if not cfg_path.exists():
        searched = [str(p) for p in _USERCFG_CANDIDATES]
        return {
            "error": f"UserCfg.opt nicht gefunden: {cfg_path}",
            "gesucht": searched,
        }

    # ── GPU-aware preset selection (works for dry_run too — no side effects) ──
    gpu_name, gpu_tier, vram_mb = _detect_gpu_tier()
    tier_presets = _VR_PRESETS_BY_TIER if mode == "VR" else _DESKTOP_PRESETS_BY_TIER
    effective_tier = gpu_tier if gpu_tier in tier_presets else "high_end"
    overrides = dict(tier_presets[effective_tier])
    if custom_overrides:
        overrides.update(custom_overrides)

    if dry_run:
        current_settings = _read_current_settings(cfg_path)
        diff = {
            k: {"old": current_settings.get(k, "<not set>"), "new": v}
            for k, v in overrides.items()
            if current_settings.get(k) != v
        }
        logger.info("optimize_msfs_graphics DRY RUN mode=%s tier=%s diff=%s",
                    mode, effective_tier, list(diff.keys()))
        return {
            "status": "dry_run",
            "mode": mode,
            "gpu": gpu_name,
            "gpu_tier": effective_tier,
            "would_apply": overrides,
            "diff": diff,
            "config_file": str(cfg_path),
            "note": "Nichts wurde geschrieben. Setze dry_run=False um wirklich anzuwenden.",
        }

    # ── CRITICAL ORDER: Kill FIRST → Write SECOND → Start THIRD ──
    # MSFS overwrites UserCfg.opt on exit. We must kill it before writing.
    msfs_was_running = _is_msfs_running()
    if msfs_was_running and auto_restart:
        _kill_msfs(timeout_s=30)
        _wait_for_file_unlocked(cfg_path, timeout_s=10)

    # Save current state before changing anything (AFTER kill so we get the final state)
    saved_version = _snapshot(cfg_path, f"Before applying '{mode}' preset")
    _prune_msfs_backups(cfg_path)

    original_text = cfg_path.read_text(encoding="utf-8")

    entries = _parse_usercfg(original_text)
    new_entries, not_applied = _apply_overrides(entries, overrides)
    new_text = _entries_to_text(new_entries)

    # Write AFTER MSFS is dead
    cfg_path.write_text(new_text, encoding="utf-8")

    # Verify the write stuck
    verify = _read_current_settings(cfg_path)
    verified_count = sum(
        1 for k, v in overrides.items()
        if k not in not_applied and verify.get(k) == v
    )
    total_applied = len(overrides) - len(not_applied)

    result = {
        "status": "ok",
        "mode": mode,
        "gpu": gpu_name,
        "gpu_tier": effective_tier,
        "vram_mb": vram_mb,
        "tier_hinweis": (
            f"GPU '{gpu_name}' erkannt → Tier '{effective_tier}'. "
            f"Preset ist auf diese GPU-Leistung abgestimmt."
        ),
        "file": str(cfg_path),
        "saved_as_version": saved_version,
        "restore_hint": f"Use restore_msfs_graphics with version={saved_version} to undo.",
        "applied": {k: v for k, v in overrides.items() if k not in not_applied},
        "skipped_keys_not_found": sorted(not_applied) if not_applied else [],
        "verifiziert": f"{verified_count}/{total_applied} Einstellungen in Datei bestätigt",
    }

    # Restart MSFS so changes are visible
    if msfs_was_running and auto_restart:
        time.sleep(2)
        try:
            vr_mode = mode == "VR"
            if vr_mode:
                vr_result = launch_msfs_vr(force_restart=False)
                result["neustart"] = (
                    f"MSFS beendet und in VR neu gestartet. "
                    f"{mode}-Preset mit {total_applied} Einstellungen aktiv."
                )
                result["vr_launch"] = vr_result
            else:
                subprocess.Popen(["cmd", "/c", "start", f"steam://run/{_MSFS2024_STEAM_APP_ID}"])
                result["neustart"] = (
                    f"MSFS beendet und neu gestartet. "
                    f"{mode}-Preset mit {total_applied} Einstellungen aktiv."
                )
        except Exception as exc:
            result["neustart"] = f"MSFS beendet, Neustart fehlgeschlagen: {exc}"
    elif msfs_was_running and not auto_restart:
        result["hinweis"] = (
            "MSFS läuft noch! Die Grafikänderungen werden erst nach einem "
            "Neustart von MSFS sichtbar. Sage 'starte MSFS in VR' um neu zu starten."
        )
    else:
        result["hinweis"] = (
            "Einstellungen gespeichert. Beim nächsten Start von MSFS werden "
            f"die {mode}-Einstellungen automatisch geladen."
        )

    return result


# ---------------------------------------------------------------------------
# German alias map + value maps for individual MSFS settings
# ---------------------------------------------------------------------------

# Maps German (and common shorthand) names → "Section.Key" in UserCfg.opt
_MSFS_SETTING_ALIASES: dict[str, str] = {
    # DLSS
    "dlss": "Video.DLSS",
    "dlss modus": "Video.DLSS",
    "dlss mode": "Video.DLSS",
    "upscaling": "Video.DLSS",
    "upscaler": "Video.DLSS",
    "taa": "Video.DLSS",           # user says "TAA" → they want to switch to DLSS or off
    "antialiasing": "Video.DLSS",
    "anti-aliasing": "Video.DLSS",
    "aa": "Video.DLSS",
    # DLSS Frame Generation
    "dlssg": "Video.DLSSG",
    "dlss fg": "Video.DLSSG",
    "frame generation": "Video.DLSSG",
    "framegen": "Video.DLSSG",
    "fg": "Video.DLSSG",
    # Render Scale — also mapped from "Schärfe / klares Bild" requests
    "render scale": "Video.RenderScale",
    "renderscale": "Video.RenderScale",
    "render auflösung": "Video.RenderScale",
    "auflösung": "Video.RenderScale",
    "klares bild": "Video.RenderScale",
    "klareres bild": "Video.RenderScale",
    "schärfer": "Video.RenderScale",
    "schaerfer": "Video.RenderScale",
    "schärfe": "Video.RenderScale",
    "schaerfe": "Video.RenderScale",
    "sharpness": "Video.RenderScale",
    "sharpen": "Video.RenderScale",
    "klarheit": "Video.RenderScale",
    # Terrain LoD
    "terrain lod": "GraphicsVR.TerrainLoD",
    "terrain": "GraphicsVR.TerrainLoD",
    "gelände": "GraphicsVR.TerrainLoD",
    "gelände detail": "GraphicsVR.TerrainLoD",
    "terrain lod desktop": "Graphics.TerrainLoD",
    "terrain desktop": "Graphics.TerrainLoD",
    # Objects LoD
    "object lod": "GraphicsVR.ObjectsLoD",
    "objekte": "GraphicsVR.ObjectsLoD",
    "objekt detail": "GraphicsVR.ObjectsLoD",
    "objects lod": "GraphicsVR.ObjectsLoD",
    "object lod desktop": "Graphics.ObjectsLoD",
    "objects desktop": "Graphics.ObjectsLoD",
    # Clouds
    "wolken": "GraphicsVR.CloudsQuality",
    "clouds": "GraphicsVR.CloudsQuality",
    "wolken qualität": "GraphicsVR.CloudsQuality",
    "wolken desktop": "Graphics.CloudsQuality",
    "clouds desktop": "Graphics.CloudsQuality",
    # Texture
    "texturen": "GraphicsVR.TextureResolution",
    "textur": "GraphicsVR.TextureResolution",
    "textures": "GraphicsVR.TextureResolution",
    "texture resolution": "GraphicsVR.TextureResolution",
    "texturen desktop": "Graphics.TextureResolution",
    # Anisotropic
    "anisotropic": "GraphicsVR.AnisotropicFilter",
    "anisotropisch": "GraphicsVR.AnisotropicFilter",
    "af": "GraphicsVR.AnisotropicFilter",
    "anisotropic desktop": "Graphics.AnisotropicFilter",
    # Shadows
    "schatten": "GraphicsVR.ShadowQuality",
    "shadows": "GraphicsVR.ShadowQuality",
    "shadow quality": "GraphicsVR.ShadowQuality",
    "schatten desktop": "Graphics.ShadowQuality",
    # SSAO
    "ssao": "GraphicsVR.SSContact",
    "ambient occlusion": "GraphicsVR.SSContact",
    "ssao desktop": "Graphics.SSContact",
    # Reflections
    "reflexionen": "GraphicsVR.Reflections",
    "reflections": "GraphicsVR.Reflections",
    "spiegelungen": "GraphicsVR.Reflections",
    "reflexionen desktop": "Graphics.Reflections",
    # Motion Blur
    "motion blur": "GraphicsVR.MotionBlur",
    "bewegungsunschärfe": "GraphicsVR.MotionBlur",
    "motion blur desktop": "Graphics.MotionBlur",
    # Depth of Field
    "dof": "Graphics.DepthOfField",
    "tiefenschärfe": "Graphics.DepthOfField",
    "depth of field": "Graphics.DepthOfField",
    # Buildings
    "gebäude": "Graphics.Buildings",
    "buildings": "Graphics.Buildings",
    # Trees
    "bäume": "Graphics.Trees",
    "trees": "Graphics.Trees",
    # Grass
    "gras": "Graphics.GrassAndBushes",
    "grass": "Graphics.GrassAndBushes",
    "büsche": "Graphics.GrassAndBushes",
    # RayTracing
    "raytracing": "RayTracing.Enabled",
    "ray tracing": "RayTracing.Enabled",
    "rt": "RayTracing.Enabled",
    "rt reflections": "RayTracing.Reflections",
    "rt reflexionen": "RayTracing.Reflections",
}

# Named value aliases → the numeric string that goes into UserCfg.opt
_MSFS_VALUE_ALIASES: dict[str, dict[str, str]] = {
    "Video.DLSS": {
        "off": "0", "aus": "0", "kein": "0", "keine": "0", "taa": "0",
        "dlaa": "1",
        "quality": "2", "qualität": "2",
        "balanced": "3", "ausgewogen": "3",
        "performance": "4", "leistung": "4",
        "ultra performance": "5", "ultra": "5",
    },
    "Video.DLSSG": {
        "off": "0", "aus": "0", "nein": "0",
        "on": "1", "an": "1", "ja": "1", "ein": "1",
    },
    "_quality_4": {  # applies to Clouds, Texture, Shadows, Buildings, Trees, Reflections
        "off": "0", "aus": "0",
        "low": "0", "niedrig": "0",
        "medium": "1", "mittel": "1",
        "high": "2", "hoch": "2",
        "ultra": "3",
    },
    "_on_off": {
        "off": "0", "aus": "0", "nein": "0",
        "on": "1", "an": "1", "ja": "1", "ein": "1",
    },
    "Graphics.GrassAndBushes": {
        "off": "0", "aus": "0",
        "low": "1", "niedrig": "1",
        "medium": "2", "mittel": "2",
        "high": "3", "hoch": "3",
        "ultra": "4",
    },
    "Graphics.AnisotropicFilter": {
        "2x": "1", "4x": "2", "8x": "3", "16x": "4",
        "2": "1", "4": "2", "8": "3", "16": "4",
    },
    "GraphicsVR.AnisotropicFilter": {
        "2x": "1", "4x": "2", "8x": "3", "16x": "4",
        "2": "1", "4": "2", "8": "3", "16": "4",
    },
}

# Map keys to their value-alias group
_MSFS_VALUE_GROUP: dict[str, str] = {
    "Video.DLSS": "Video.DLSS",
    "Video.DLSSG": "Video.DLSSG",
    "RayTracing.Enabled": "_on_off",
    "RayTracing.Reflections": "_on_off",
    "Graphics.SSContact": "_on_off",
    "GraphicsVR.SSContact": "_on_off",
    "Graphics.MotionBlur": "_on_off",
    "GraphicsVR.MotionBlur": "_on_off",
    "Graphics.DepthOfField": "_on_off",
    "Graphics.AnisotropicFilter": "Graphics.AnisotropicFilter",
    "GraphicsVR.AnisotropicFilter": "GraphicsVR.AnisotropicFilter",
    "Graphics.GrassAndBushes": "Graphics.GrassAndBushes",
}

# Maps VR settings to their desktop counterparts and vice versa
_VR_DESKTOP_PAIRS: dict[str, str] = {
    "GraphicsVR.TerrainLoD":        "Graphics.TerrainLoD",
    "GraphicsVR.ObjectsLoD":        "Graphics.ObjectsLoD",
    "GraphicsVR.CloudsQuality":     "Graphics.CloudsQuality",
    "GraphicsVR.AnisotropicFilter": "Graphics.AnisotropicFilter",
    "GraphicsVR.SSContact":         "Graphics.SSContact",
    "GraphicsVR.Reflections":       "Graphics.Reflections",
    "GraphicsVR.TextureResolution": "Graphics.TextureResolution",
    "GraphicsVR.MotionBlur":        "Graphics.MotionBlur",
    "GraphicsVR.ShadowQuality":     "Graphics.ShadowQuality",
    # reverse mappings
    "Graphics.TerrainLoD":          "GraphicsVR.TerrainLoD",
    "Graphics.ObjectsLoD":          "GraphicsVR.ObjectsLoD",
    "Graphics.CloudsQuality":       "GraphicsVR.CloudsQuality",
    "Graphics.AnisotropicFilter":   "GraphicsVR.AnisotropicFilter",
    "Graphics.SSContact":           "GraphicsVR.SSContact",
    "Graphics.Reflections":         "GraphicsVR.Reflections",
    "Graphics.TextureResolution":   "GraphicsVR.TextureResolution",
    "Graphics.MotionBlur":          "GraphicsVR.MotionBlur",
    "Graphics.ShadowQuality":       "GraphicsVR.ShadowQuality",
}

# Default group for quality-type settings (Clouds, Texture, Shadows, etc.)
_QUALITY_KEYS = {
    "Graphics.CloudsQuality", "GraphicsVR.CloudsQuality",
    "Graphics.TextureResolution", "GraphicsVR.TextureResolution",
    "Graphics.ShadowQuality", "GraphicsVR.ShadowQuality",
    "Graphics.Reflections", "GraphicsVR.Reflections",
    "Graphics.Buildings", "Graphics.Trees",
}


@mcp.tool()
def set_msfs_setting(
    setting: str,
    value: str,
    usercfg_path: str = "",
    auto_restart: bool = True,
    restart_in_vr: bool = False,
    apply_to_both: bool = True,
) -> dict:
    """Change a single MSFS 2024 graphics setting and restart so it takes effect.

    WICHTIG: MSFS muss neu gestartet werden damit Änderungen sichtbar werden!
    Standardmäßig wird MSFS automatisch beendet und neu gestartet.

    Use this when the user says things like:
    - "DLSS auf Quality stellen"
    - "Wolken auf Ultra"
    - "Schatten auf Medium"
    - "TAA auf DLSS umstellen"
    - "RayTracing ausschalten"
    - "Terrain LoD auf 2.0"
    - "Texturen auf High"
    - "Motion Blur aus"
    - "Anisotropisch auf 16x"

    Args:
        setting: Setting name (German or English). Examples:
                 "dlss", "wolken", "schatten", "texturen", "terrain lod",
                 "raytracing", "motion blur", "ssao", "reflexionen", etc.
        value: New value. Can be a name ("quality", "ultra", "aus", "an")
               or a number ("2", "1.5", "0"). For LoD values use decimals (e.g. "2.0").
        usercfg_path: Path to UserCfg.opt (auto-detected if empty).
        auto_restart: Restart MSFS automatically so the change is visible (default True).
        restart_in_vr: If restarting, use VR startup sequence (default False).
        apply_to_both: Also apply the change to the desktop/VR counterpart setting (default True).
                       Ensures the setting is active regardless of whether MSFS runs in VR or desktop mode.
    """
    import time

    cfg_path = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH
    if not cfg_path.exists():
        # Report which paths were checked (all candidates, not just missing ones)
        searched = [str(p) for p in _USERCFG_CANDIDATES]
        return {
            "error": f"UserCfg.opt nicht gefunden: {cfg_path}",
            "gesucht": searched[:6],
            "tipp": "Gib den Pfad zu deiner UserCfg.opt über usercfg_path an.",
        }

    # Resolve setting alias → "Section.Key"
    setting_lower = setting.strip().lower()
    config_key = _MSFS_SETTING_ALIASES.get(setting_lower)

    if config_key is None:
        for k in SETTING_DEFS:
            if k.lower() == setting_lower or k.split(".", 1)[-1].lower() == setting_lower:
                config_key = k
                break

    if config_key is None:
        for alias, key in _MSFS_SETTING_ALIASES.items():
            if setting_lower in alias or alias in setting_lower:
                config_key = key
                break

    if config_key is None:
        available = sorted(set(_MSFS_SETTING_ALIASES.values()))
        return {
            "error": f"Unbekannte Einstellung: '{setting}'",
            "verfügbare_einstellungen": available,
            "beispiele": [
                "dlss → Video.DLSS",
                "wolken → CloudsQuality",
                "schatten → ShadowQuality",
                "terrain → TerrainLoD",
                "raytracing → RayTracing",
            ],
        }

    # Resolve value alias → numeric string
    value_lower = value.strip().lower()
    group = _MSFS_VALUE_GROUP.get(config_key)
    if group is None and config_key in _QUALITY_KEYS:
        group = "_quality_4"

    if group and group in _MSFS_VALUE_ALIASES:
        resolved_value = _MSFS_VALUE_ALIASES[group].get(value_lower, value.strip())
    else:
        resolved_value = value.strip()

    # Read current value BEFORE killing (for display purposes)
    current_settings = _read_current_settings(cfg_path)
    old_value = current_settings.get(config_key, "?")

    defn = SETTING_DEFS.get(config_key, {})
    label = defn.get("label", config_key)
    value_map = defn.get("values", {})
    old_display = value_map.get(str(old_value), str(old_value))
    new_display = value_map.get(str(resolved_value), str(resolved_value))

    # ── CRITICAL ORDER: Kill FIRST → Write SECOND → Start THIRD ──
    # MSFS overwrites UserCfg.opt on exit. If we write first and then
    # kill, MSFS saves its own settings during shutdown and our changes
    # are lost. So we MUST kill MSFS before touching the file.

    msfs_was_running = _is_msfs_running()
    if msfs_was_running:
        _kill_msfs(timeout_s=35)
        # Wait for MSFS to fully flush UserCfg.opt to disk before we write
        _wait_for_file_unlocked(cfg_path, timeout_s=25)
        time.sleep(1.5)  # Extra OS flush buffer

    # Re-read the file AFTER kill (MSFS may have written to it during shutdown)
    current_settings = _read_current_settings(cfg_path)
    old_value_after_kill = current_settings.get(config_key, "?")
    old_display = value_map.get(str(old_value_after_kill), str(old_value_after_kill))

    # Save snapshot before changing
    saved_version = _snapshot(cfg_path, f"Before setting {config_key}={resolved_value}")

    # NOW write the change (MSFS is dead, can't overwrite us)
    original_text = cfg_path.read_text(encoding="utf-8")
    entries = _parse_usercfg(original_text)
    new_entries, not_applied = _apply_overrides(entries, {config_key: resolved_value})
    new_text = _entries_to_text(new_entries)
    cfg_path.write_text(new_text, encoding="utf-8")
    _invalidate_usercfg_cache()

    if config_key in not_applied:
        return {
            "error": (
                f"Einstellung '{config_key}' wurde in UserCfg.opt nicht gefunden. "
                f"Möglicherweise existiert der Schlüssel unter einem anderen Namen."
            ),
            "config_file": str(cfg_path),
            "saved_version": saved_version,
        }

    # Also apply to desktop/VR counterpart so both modes see the change
    twin_key_applied: str | None = None
    if apply_to_both and config_key in _VR_DESKTOP_PAIRS:
        twin_key = _VR_DESKTOP_PAIRS[config_key]
        twin_entries, twin_not_applied = _apply_overrides(new_entries, {twin_key: resolved_value})
        if twin_key not in twin_not_applied:
            new_entries = twin_entries
            cfg_path.write_text(_entries_to_text(new_entries), encoding="utf-8")
            _invalidate_usercfg_cache()
            twin_key_applied = twin_key

    # Verify the write actually stuck
    verify_settings = _read_current_settings(cfg_path)
    verified = verify_settings.get(config_key) == resolved_value

    result = {
        "status": "ok" if verified else "verifizierung_fehlgeschlagen",
        "einstellung": label,
        "key": config_key,
        "vorher": old_display,
        "nachher": new_display,
        "raw_value": resolved_value,
        "config_file": str(cfg_path),
        "saved_version": saved_version,
        "verifiziert": verified,
    }

    if twin_key_applied:
        result["auch_angewendet_auf"] = twin_key_applied

    if not verified:
        result["warnung"] = (
            f"FEHLER: Wert wurde NICHT korrekt gespeichert! "
            f"Datei zeigt: {verify_settings.get(config_key, '?')} "
            f"(erwartet: {resolved_value}). "
            f"Bitte prüfe ob UserCfg.opt schreibgeschützt ist oder ein anderer Prozess sie blockiert."
        )

    # Restart MSFS so changes become visible
    if auto_restart and (msfs_was_running or restart_in_vr):
        time.sleep(2)
        action = "beendet und in VR neu gestartet" if msfs_was_running else "in VR gestartet"
        action_plain = "beendet und neu gestartet" if msfs_was_running else "gestartet"
        if restart_in_vr:
            vr_result = launch_msfs_vr(force_restart=False)
            result["neustart"] = (
                f"MSFS {action}. {label}: {old_display} → {new_display}"
            )
            result["vr_launch"] = vr_result
        else:
            try:
                subprocess.Popen(["cmd", "/c", "start", f"steam://run/{_MSFS2024_STEAM_APP_ID}"])
                result["neustart"] = (
                    f"MSFS {action_plain}. {label}: {old_display} → {new_display}"
                )
            except Exception as exc:
                result["neustart"] = f"MSFS-Aktion fehlgeschlagen: {exc}"
    elif not auto_restart:
        result["hinweis"] = (
            f"Einstellung gespeichert: {label} = {new_display}. "
            f"MSFS muss neu gestartet werden damit die Änderung sichtbar wird."
        )
    else:
        result["hinweis"] = (
            f"Einstellung gespeichert: {label} = {new_display}. "
            f"Wird beim nächsten MSFS-Start übernommen."
        )

    return result


@mcp.tool()
def restore_msfs_graphics(
    version: int = 0,
    usercfg_path: str = "",
) -> dict:
    """Restore MSFS graphics settings to a previously saved version.

    Every call to optimize_msfs_graphics automatically saves the config
    before making changes.  Use this tool to list all saved versions or
    roll back to a specific one.

    Args:
        version: Version number to restore.  Pass 0 (or omit) to list all
                 available versions without restoring.
        usercfg_path: Full path to UserCfg.opt (leave empty for default).
    """
    cfg_path = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH
    history = _load_history(cfg_path)

    if not history:
        return {"error": "No saved versions found. History is empty."}

    # List mode
    if version == 0:
        summary = [
            {
                "version": h["version"],
                "timestamp": h["timestamp"],
                "label": h["label"],
            }
            for h in history
        ]
        return {
            "available_versions": summary,
            "total": len(summary),
            "hint": "Call again with version=N to restore that version.",
        }

    # Find the requested version
    entry = next((h for h in history if h["version"] == version), None)
    if entry is None:
        return {
            "error": f"Version {version} not found.",
            "available": [h["version"] for h in history],
        }

    # Save current state before restoring (so restore is also undoable)
    if cfg_path.exists():
        _snapshot(cfg_path, f"Before restoring to version {version}")

    cfg_path.write_text(entry["content"], encoding="utf-8")

    return {
        "status": "ok",
        "restored_version": version,
        "label": entry["label"],
        "timestamp": entry["timestamp"],
        "file": str(cfg_path),
    }


@mcp.tool()
def backup_msfs_graphics(
    action: Literal["create", "list", "restore_latest"],
    label: str = "Manual backup",
    usercfg_path: str = "",
) -> dict:
    """Manually create, list or restore MSFS graphics backups.

    Backups are stored as individual .opt files in UserCfg_backups/ next to
    UserCfg.opt.  You can also restore the most recent backup at any time.

    Args:
        action: 'create' a new backup, 'list' existing backups, or
                'restore_latest' to restore the last backup.
        label: Description for the backup (only used with 'create').
        usercfg_path: Full path to UserCfg.opt (leave empty for default).
    """
    cfg_path = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH

    if action == "create":
        if not cfg_path.exists():
            return {"error": f"UserCfg.opt not found at {cfg_path}"}
        version = _snapshot(cfg_path, label)
        bdir = _backup_dir(cfg_path)
        return {
            "status": "ok",
            "saved_version": version,
            "backup_folder": str(bdir),
            "label": label,
        }

    if action == "list":
        bdir = _backup_dir(cfg_path)
        if not bdir.exists():
            return {"backups": [], "count": 0}
        files = sorted(bdir.glob("UserCfg_v*.opt"), reverse=True)
        backups = [
            {"file": f.name, "size_kb": round(f.stat().st_size / 1024, 1),
             "modified": datetime.datetime.fromtimestamp(f.stat().st_mtime).isoformat(timespec="seconds")}
            for f in files
        ]
        return {"backups": backups, "count": len(backups), "folder": str(bdir)}

    if action == "restore_latest":
        latest = _backup_dir(cfg_path) / "UserCfg_latest.opt"
        if not latest.exists():
            return {"error": "No backup found. Create one first."}
        # Save current state before restoring
        if cfg_path.exists():
            _snapshot(cfg_path, "Before restoring latest backup")
        shutil.copy2(latest, cfg_path)
        return {
            "status": "ok",
            "restored_from": str(latest),
            "file": str(cfg_path),
        }

    return {"error": f"Unknown action: {action}"}


@mcp.tool()
def check_and_install_driver(
    auto_install: bool = False,
    download_dir: str = "",
) -> dict:
    """Check the installed NVIDIA driver against the latest available version.

    Reads the current driver version and GPU model via NVML, then queries
    the NVIDIA driver API for the newest WHQL-certified driver.  Returns a
    comparison and the download link.  If *auto_install* is True the
    installer is downloaded and launched in silent mode (``/s /noreboot``).

    Args:
        auto_install: If True, download the installer and run it silently.
                      Requires the tool to run with administrator privileges.
        download_dir: Directory to save the downloaded installer.
                      Defaults to the user's Downloads folder.
    """
    # 1. Current driver info via NVML
    try:
        current_version, gpu_name = _get_gpu_info()
    except Exception as exc:
        return {"error": f"Could not read GPU info via NVML: {exc}"}

    # 2. Map GPU name to NVIDIA API parameters (dynamic lookup, then static)
    ids = _lookup_pfid(gpu_name)
    if ids is None:
        ids = _match_gpu_model(gpu_name)
    if ids is None:
        return {
            "error": (
                f"GPU '{gpu_name}' not found via NVIDIA API or static table. "
                "Check manually at: https://www.nvidia.com/Download/index.aspx"
            ),
            "current_driver": current_version,
            "gpu_name": gpu_name,
        }
    psid, pfid = ids

    # 3. Query NVIDIA for the latest driver
    try:
        latest = _query_latest_driver(psid, pfid)
    except Exception as exc:
        return {
            "error": f"NVIDIA API request failed: {exc}",
            "current_driver": current_version,
            "gpu_name": gpu_name,
        }

    if latest is None:
        return {
            "error": "Could not parse a driver version from the NVIDIA API response.",
            "current_driver": current_version,
            "gpu_name": gpu_name,
        }

    # 4. Compare versions
    up_to_date = _version_tuple(current_version) >= _version_tuple(latest["version"])

    result: dict = {
        "gpu_name": gpu_name,
        "current_driver": current_version,
        "latest_driver": latest["version"],
        "driver_name": latest["name"],
        "release_date": latest["release_date"],
        "up_to_date": up_to_date,
        "download_url": latest["download_url"],
    }

    if up_to_date:
        result["message"] = "Your driver is already up to date."
        return result

    result["message"] = (
        f"Update available: {current_version} -> {latest['version']}"
    )

    if not auto_install:
        return result

    # 5. Download the installer
    dl_dir = Path(download_dir) if download_dir else Path.home() / "Downloads"
    dl_dir.mkdir(parents=True, exist_ok=True)

    try:
        installer = _download(latest["download_url"], dl_dir)
    except Exception as exc:
        result["install_error"] = f"Download failed: {exc}"
        return result

    result["installer_path"] = str(installer)

    # 6. Launch in silent mode
    try:
        proc = subprocess.Popen(
            [str(installer), "/s", "/noreboot"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        result["install_status"] = "installer_started"
        result["installer_pid"] = proc.pid
        result["message"] = (
            f"Installer launched in silent mode (PID {proc.pid}). "
            "A reboot may be required after installation completes."
        )
    except Exception as exc:
        result["install_error"] = (
            f"Could not launch installer: {exc}. "
            "Try running as administrator."
        )

    return result


# ---------------------------------------------------------------------------
# install_mod helpers
# ---------------------------------------------------------------------------

def _detect_community_folder(custom_path: str) -> Path | None:
    """Return the Community folder path, or None if not found."""
    if custom_path:
        p = Path(custom_path)
        if p.is_dir():
            return p
        return None
    for candidate in COMMUNITY_CANDIDATES:
        if candidate.is_dir():
            return candidate
    return None


def _download(url: str, dest_dir: Path) -> Path:
    """Download a file from *url* into *dest_dir* and return the local path."""
    with httpx.Client(follow_redirects=True, timeout=300) as client:
        with client.stream("GET", url) as resp:
            resp.raise_for_status()

            # Try to derive a filename from Content-Disposition or the URL
            filename = None
            cd = resp.headers.get("content-disposition", "")
            m = re.search(r'filename\*?=["\']?(?:UTF-8\'\')?([^"\';]+)', cd, re.IGNORECASE)
            if m:
                filename = m.group(1).strip()
            if not filename:
                filename = url.split("/")[-1].split("?")[0]
            if not filename:
                filename = "mod_download.zip"

            local_path = dest_dir / filename
            with open(local_path, "wb") as f:
                for chunk in resp.iter_bytes(chunk_size=1024 * 256):
                    f.write(chunk)
    return local_path


def _extract(archive: Path, dest: Path) -> None:
    """Extract a .zip archive into *dest*."""
    suffix = archive.suffix.lower()
    if suffix == ".zip":
        with zipfile.ZipFile(archive) as zf:
            zf.extractall(dest)
    elif suffix in (".7z", ".rar"):
        # py7zr for .7z, rarfile for .rar — optional deps
        if suffix == ".7z":
            import py7zr
            with py7zr.SevenZipFile(archive) as sz:
                sz.extractall(dest)
        else:
            import rarfile
            with rarfile.RarFile(archive) as rf:
                rf.extractall(dest)
    else:
        raise ValueError(f"Unsupported archive format: {suffix}")


def _find_mod_roots(extract_dir: Path) -> list[Path]:
    """Identify mod root folders inside the extracted tree.

    A mod root is a folder that contains a manifest.json or layout.json
    (MSFS add-on markers).  If none are found we fall back to the
    top-level directories.
    """
    roots: list[Path] = []
    for marker in ("manifest.json", "layout.json"):
        for hit in extract_dir.rglob(marker):
            root = hit.parent
            if root not in roots:
                roots.append(root)
    if not roots:
        # Fallback: top-level dirs inside the extract
        roots = [d for d in extract_dir.iterdir() if d.is_dir()]
    return roots


def _resolve_flightsim_to(url: str) -> str:
    """If *url* is a flightsim.to page URL, try to resolve the direct download link.

    Supports URLs like:
      - https://flightsim.to/file/12345/some-mod
      - https://flightsim.to/addon/12345/some-mod
    Tries the flightsim.to API first, then scrapes the page for download links.
    Returns the original URL unchanged if it's not a flightsim.to page.
    """
    # Only process flightsim.to page URLs (not already direct download links)
    m = re.match(r"https?://flightsim\.to/(?:file|addon)/(\d+)", url)
    if not m:
        return url

    addon_id = m.group(1)

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,*/*",
    }

    with httpx.Client(follow_redirects=True, timeout=30, headers=headers) as client:
        # Try 1: flightsim.to API endpoint
        try:
            api_url = f"https://flightsim.to/api/v1/addon/{addon_id}"
            resp = client.get(api_url)
            if resp.status_code == 200:
                data = resp.json()
                dl = data.get("download_url") or data.get("download", {}).get("url")
                if dl:
                    return dl
        except Exception:
            pass

        # Try 2: Fetch the page and look for download links
        try:
            resp = client.get(url)
            resp.raise_for_status()
            html = resp.text

            # Look for direct download links in the HTML
            patterns = [
                r'href="(https?://flightsim\.to/api/v1/download/[^"]+)"',
                r'href="(https?://flightsim\.to/file/\d+/[^"]*?/download[^"]*)"',
                r'href="(https?://cdn[^"]*flightsim[^"]*\.(?:zip|7z|rar)[^"]*)"',
                r'data-url="(https?://[^"]+\.(?:zip|7z|rar)[^"]*)"',
                r'"downloadUrl"\s*:\s*"(https?://[^"]+)"',
            ]
            for pattern in patterns:
                match = re.search(pattern, html, re.IGNORECASE)
                if match:
                    return match.group(1)

            # Try 3: Construct the standard download page URL
            download_page = f"https://flightsim.to/file/{addon_id}/download"
            resp2 = client.get(download_page)
            if resp2.status_code == 200:
                for pattern in patterns:
                    match = re.search(pattern, resp2.text, re.IGNORECASE)
                    if match:
                        return match.group(1)
        except Exception:
            pass

    return url


@mcp.tool()
def install_mod(
    url: str,
    community_path: str = "",
) -> dict:
    """Download and install an MSFS 2024 mod into the Community folder.

    Accepts both flightsim.to PAGE URLs (e.g. https://flightsim.to/addon/12345/name)
    and direct download URLs.  For flightsim.to pages the download link is
    resolved automatically.

    Supported archive formats: .zip (built-in), .7z (needs py7zr), .rar (needs rarfile).

    Args:
        url: flightsim.to addon page URL or direct download URL (.zip, .7z, .rar).
        community_path: Full path to your MSFS 2024 Community folder.
                        Leave empty for auto-detection (Steam & MS Store).
    """
    community = _detect_community_folder(community_path)
    if community is None:
        return {
            "error": (
                "Community folder not found. "
                "Please pass the full path via the community_path parameter."
            )
        }

    # Resolve flightsim.to page URLs to direct download links
    resolved_url = _resolve_flightsim_to(url)

    with tempfile.TemporaryDirectory(prefix="msfs_mod_") as tmp:
        tmp_path = Path(tmp)

        # 1. Download
        try:
            archive = _download(resolved_url, tmp_path)
        except httpx.HTTPStatusError as exc:
            return {
                "error": f"Download failed: HTTP {exc.response.status_code}",
                "resolved_url": resolved_url,
                "hint": (
                    "Falls die automatische Auflösung nicht funktioniert hat: "
                    "Öffne die Seite im Browser, klicke auf Download, "
                    "kopiere die direkte Download-URL und gib sie hier ein."
                ),
            }
        except Exception as exc:
            return {"error": f"Download failed: {exc}", "resolved_url": resolved_url}

        # 2. Extract
        extract_dir = tmp_path / "extracted"
        extract_dir.mkdir()
        try:
            _extract(archive, extract_dir)
        except Exception as exc:
            return {"error": f"Extraction failed: {exc}"}

        # 3. Locate mod root(s)
        mod_roots = _find_mod_roots(extract_dir)
        if not mod_roots:
            return {"error": "No mod folders found in the archive."}

        # 4. Move into Community
        installed: list[str] = []
        for root in mod_roots:
            dest = community / root.name
            if dest.exists():
                shutil.rmtree(dest)
            shutil.copytree(root, dest)
            installed.append(root.name)

    return {
        "status": "ok",
        "community_folder": str(community),
        "installed_mods": installed,
    }


# ---------------------------------------------------------------------------
# Pimax settings analysis & optimization
# ---------------------------------------------------------------------------

# Possible Pimax config file locations
_PIMAX_CONFIG_CANDIDATES = [
    # Pimax Play / Pimax Client data directories
    Path(os.environ.get("LOCALAPPDATA", "")) / "Pimax" / "PimaxPlay",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Pimax" / "Pimax Play",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Pimax" / "PimaxClient",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Pimax",
    Path(os.environ.get("LOCALAPPDATA", "")) / "PimaxPlay",
    Path(os.environ.get("LOCALAPPDATA", "")) / "PimaxVR",
    Path(os.environ.get("APPDATA", "")) / "Pimax" / "PimaxPlay",
    Path(os.environ.get("APPDATA", "")) / "Pimax",
    Path(os.environ.get("APPDATA", "")) / "PimaxPlay",
    Path(os.environ.get("PROGRAMDATA", "")) / "Pimax" / "PimaxPlay",
    Path(os.environ.get("PROGRAMDATA", "")) / "Pimax",
    # Program Files install directories
    Path(r"C:\Program Files\Pimax\PimaxPlay"),
    Path(r"C:\Program Files\Pimax\Pimax Play"),
    Path(r"C:\Program Files\Pimax\Runtime"),
    Path(r"C:\Program Files\Pimax\PimaxClient"),
    Path(r"C:\Program Files\Pimax"),
    Path(r"C:\Program Files (x86)\Pimax"),
    Path(r"D:\Program Files\Pimax"),
    Path(r"D:\Pimax"),
    # User profile locations
    Path(os.environ.get("USERPROFILE", "")) / "Pimax",
    Path(os.environ.get("USERPROFILE", "")) / ".pimax",
    Path(os.environ.get("USERPROFILE", "")) / "AppData" / "Local" / "Pimax",
]

_PIMAX_CONFIG_FILENAMES = [
    "settings.json", "config.json", "pimax.json",
    "PimaxPlay.json", "user_settings.json", "device_settings.json",
    "pimax_settings.json", "hmd_settings.json", "display_settings.json",
    "general.json", "preference.json", "preferences.json",
]

# Human-readable labels and value maps for Pimax settings
PIMAX_SETTING_DEFS: dict[str, dict] = {
    "renderResolution": {
        "label": "Render Resolution",
        "description": "SteamVR Supersampling-Multiplikator",
    },
    "renderQuality": {
        "label": "Render Quality",
        "values": {"0": "Low", "1": "Medium", "2": "High", "3": "Ultra"},
    },
    "refreshRate": {
        "label": "Refresh Rate (Hz)",
        "description": "Display-Wiederholrate",
    },
    "fov": {
        "label": "Field of View",
        "values": {"small": "Small (100°)", "normal": "Normal (140°)",
                   "large": "Large (170°)", "potato": "Potato (80°)"},
    },
    "fovLevel": {
        "label": "Field of View",
        "values": {"0": "Potato", "1": "Small", "2": "Normal", "3": "Large"},
    },
    "smartSmoothing": {
        "label": "Smart Smoothing",
        "values": {"true": "An", "false": "Aus", "True": "An", "False": "Aus",
                   "0": "Aus", "1": "An"},
    },
    "compulsorySmoothing": {
        "label": "Compulsory Smoothing",
        "values": {"true": "An", "false": "Aus", "True": "An", "False": "Aus",
                   "0": "Aus", "1": "An"},
    },
    "ffrLevel": {
        "label": "Fixed Foveated Rendering",
        "values": {"0": "Aus", "1": "Low", "2": "Medium", "3": "High", "4": "Ultra",
                   "off": "Aus", "low": "Low", "medium": "Medium", "high": "High"},
    },
    "parallelProjection": {
        "label": "Parallel Projection",
        "values": {"true": "An", "false": "Aus", "True": "An", "False": "Aus",
                   "0": "Aus", "1": "An"},
    },
    "brightness": {
        "label": "Backlight Brightness",
        "description": "Display-Helligkeit (Bereich variiert je nach Pimax-Version)",
    },
    "contrast": {
        "label": "Contrast",
    },
    "ipd": {
        "label": "IPD (mm)",
    },
    "ipdOffset": {
        "label": "IPD Offset",
    },
    "smoothingMode": {
        "label": "Smoothing Mode",
    },
    "eyeTracking": {
        "label": "Eye Tracking",
        "values": {"true": "An", "false": "Aus", "0": "Aus", "1": "An"},
    },
    "dynamicFoveatedRendering": {
        "label": "Dynamic Foveated Rendering",
        "values": {"true": "An", "false": "Aus", "0": "Aus", "1": "An"},
    },
    # ── Pimax Play echte Config-Keys (Device Settings → Games → Color Options) ──
    "piplay_color_brightness_0": {
        "label": "Brightness Left Eye",
        "description": "Pimax Play Device Settings Brightness linkes Auge. Slider-Wert (Ganzzahl).",
    },
    "piplay_color_brightness_1": {
        "label": "Brightness Right Eye",
        "description": "Pimax Play Device Settings Brightness rechtes Auge. Slider-Wert (Ganzzahl).",
    },
    "piplay_color_contrast_0": {
        "label": "Contrast Left Eye",
        "description": "Pimax Play Device Settings Contrast linkes Auge. Slider-Wert (Ganzzahl).",
    },
    "piplay_color_contrast_1": {
        "label": "Contrast Right Eye",
        "description": "Pimax Play Device Settings Contrast rechtes Auge. Slider-Wert (Ganzzahl).",
    },
    "piplay_color_saturation_0": {
        "label": "Saturation Left Eye",
        "description": "Pimax Play Device Settings Saturation linkes Auge. Slider-Wert.",
    },
    "piplay_color_saturation_1": {
        "label": "Saturation Right Eye",
        "description": "Pimax Play Device Settings Saturation rechtes Auge. Slider-Wert.",
    },
    "piplay_refreshrate": {
        "label": "Refresh Rate (Hz)",
        "description": "Pimax Play Display-Wiederholrate.",
    },
    "piplay_fov": {
        "label": "Field of View",
    },
    "piplay_smartsmoothing": {
        "label": "Smart Smoothing",
        "values": {"true": "An", "false": "Aus", "0": "Aus", "1": "An"},
    },
    "piplay_ffr": {
        "label": "Fixed Foveated Rendering",
    },
    "piplay_pp": {
        "label": "Parallel Projection",
        "values": {"true": "An", "false": "Aus", "0": "Aus", "1": "An"},
    },
    "piplay_ipd": {
        "label": "IPD (mm)",
    },
}

# Pimax presets: key → value overrides
PIMAX_PRESETS: dict[str, dict] = {
    "Performance": {
        "renderResolution": 1.0,
        "refreshRate": 90,
        "fov": "normal",
        "fovLevel": 2,
        "smartSmoothing": True,
        "compulsorySmoothing": False,
        "ffrLevel": 2,
        "parallelProjection": False,
    },
    "Balanced": {
        "renderResolution": 1.25,
        "refreshRate": 90,
        "fov": "normal",
        "fovLevel": 2,
        "smartSmoothing": True,
        "compulsorySmoothing": False,
        "ffrLevel": 1,
        "parallelProjection": False,
    },
    "Quality": {
        "renderResolution": 1.5,
        "refreshRate": 90,
        "fov": "large",
        "fovLevel": 3,
        "smartSmoothing": False,
        "compulsorySmoothing": False,
        "ffrLevel": 0,
        "parallelProjection": False,
    },
    "MSFS_VR_Optimized": {
        "renderResolution": 1.0,
        "refreshRate": 72,
        "fov": "normal",
        "fovLevel": 2,
        "smartSmoothing": True,
        "compulsorySmoothing": True,
        "ffrLevel": 2,
        "parallelProjection": True,
        "brightness": 1,
    },
}

# Pimax VRAM-based recommendations
PIMAX_VRAM_TIPS: list[tuple[int, str, list[tuple[str, str, str]]]] = [
    # (min_vram_gb, tier, [(setting, suggestion, reason), ...])
    (16, "high_end", [
        ("renderResolution", "1.25 – 1.5",
         "Deine RTX 5080 kann höhere Render-Auflösung in Pimax gut handeln."),
        ("ffrLevel", "0 (Aus) oder 1 (Low)",
         "FFR bei 16 GB VRAM kaum nötig — nur bei Engpässen aktivieren."),
        ("refreshRate", "90",
         "90 Hz ist ideal für MSFS. 120 Hz nur wenn stabile FPS möglich."),
    ]),
    (10, "mid_high", [
        ("renderResolution", "1.0 – 1.25",
         "Render Resolution moderat halten, DLSS übernimmt das Upscaling."),
        ("ffrLevel", "1 (Low) oder 2 (Medium)",
         "FFR auf Low/Medium spart VRAM bei kaum sichtbarem Verlust."),
    ]),
    (8, "mid", [
        ("renderResolution", "1.0",
         "Bei 8 GB VRAM Pimax Render Resolution auf 1.0 lassen."),
        ("ffrLevel", "2 (Medium) oder 3 (High)",
         "FFR unbedingt aktivieren um VRAM zu sparen."),
        ("fov", "Normal oder Small",
         "FOV reduzieren spart massiv GPU-Leistung."),
    ]),
    (0, "low", [
        ("renderResolution", "0.75 – 1.0", "Render Resolution minimieren."),
        ("ffrLevel", "3 (High)", "FFR auf Maximum für spielbare FPS."),
        ("fov", "Small oder Potato", "FOV stark reduzieren."),
    ]),
]


_PIMAX_CONFIG_KEYS = (
    "renderResolution", "refreshRate", "fov", "fovLevel",
    "ffrLevel", "smartSmoothing", "parallelProjection",
    "brightness", "contrast", "ipd", "renderQuality",
    "compulsorySmoothing", "eyeTracking",
)


def _looks_like_pimax_config(path: Path) -> bool:
    """Check if a JSON file looks like a Pimax settings file."""
    try:
        data = _json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict) and any(k in data for k in _PIMAX_CONFIG_KEYS):
            return True
    except Exception:
        pass
    return False


def _find_pimax_config(custom_path: str = "") -> Path | None:
    """Find the Pimax settings JSON file.

    Search order:
    1. Custom path (if provided)
    2. Known config directories + filenames
    3. Any .json/.cfg in known directories that looks like Pimax config
    4. Registry-based install path
    5. Recursive search in Pimax install directories
    """
    if custom_path:
        p = Path(custom_path)
        if p.is_file():
            return p
        if p.is_dir():
            for fname in _PIMAX_CONFIG_FILENAMES:
                candidate = p / fname
                if candidate.exists():
                    return candidate
            # Search recursively in the given directory
            for json_file in p.rglob("*.json"):
                if _looks_like_pimax_config(json_file):
                    return json_file
        return None

    # Step 2: Known directories + known filenames
    for config_dir in _PIMAX_CONFIG_CANDIDATES:
        if not config_dir.exists():
            continue
        for fname in _PIMAX_CONFIG_FILENAMES:
            candidate = config_dir / fname
            if candidate.exists():
                return candidate

    # Step 3: Any .json/.cfg in known directories
    for config_dir in _PIMAX_CONFIG_CANDIDATES:
        if not config_dir.exists():
            continue
        for pattern in ("*.json", "*.cfg"):
            for f in config_dir.glob(pattern):
                if _looks_like_pimax_config(f):
                    return f

    # Step 4: Find Pimax install via registry, then search there
    try:
        result = _reg(
            ["query",
             r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
             "/s", "/f", "Pimax", "/d"],
            timeout=10,
        )
        for line in result.stdout.splitlines():
            m = re.search(r"InstallLocation\s+REG_SZ\s+(.+)", line)
            if m:
                install_dir = Path(m.group(1).strip())
                if install_dir.exists():
                    for fname in _PIMAX_CONFIG_FILENAMES:
                        candidate = install_dir / fname
                        if candidate.exists():
                            return candidate
                    for f in install_dir.rglob("*.json"):
                        if _looks_like_pimax_config(f):
                            return f
    except Exception:
        pass

    # Step 5: Also check HKCU registry
    try:
        result = _reg(
            ["query",
             r"HKCU\SOFTWARE\Pimax",
             "/s", "/v", "ConfigPath"],
            timeout=10,
        )
        for line in result.stdout.splitlines():
            m = re.search(r"ConfigPath\s+REG_SZ\s+(.+)", line)
            if m:
                p = Path(m.group(1).strip())
                if p.is_file():
                    return p
                if p.is_dir():
                    for fname in _PIMAX_CONFIG_FILENAMES:
                        candidate = p / fname
                        if candidate.exists():
                            return candidate
    except Exception:
        pass

    # Step 6: Brute-force search in all Pimax-related Program Files dirs
    for root in [r"C:\Program Files", r"C:\Program Files (x86)", r"D:\Program Files"]:
        root_path = Path(root)
        if not root_path.exists():
            continue
        for pimax_dir in root_path.glob("Pimax*"):
            if pimax_dir.is_dir():
                for f in pimax_dir.rglob("*.json"):
                    if _looks_like_pimax_config(f):
                        return f

    return None


def _read_pimax_registry() -> dict:
    """Read Pimax settings from the Windows registry."""
    settings: dict = {}
    reg_paths = [
        r"HKCU\SOFTWARE\Pimax",
        r"HKCU\SOFTWARE\Pimax\PimaxPlay",
        r"HKCU\SOFTWARE\Pimax\Pimax Play",
        r"HKLM\SOFTWARE\Pimax",
        r"HKLM\SOFTWARE\Pimax\PimaxPlay",
    ]
    for reg_path in reg_paths:
        try:
            result = _reg(["query", reg_path, "/s"], timeout=5)
            if result.returncode != 0:
                continue
            for line in result.stdout.splitlines():
                line = line.strip()
                # Match lines like: piplay_color_brightness_0    REG_DWORD    0x62
                m = re.match(r"(\S+)\s+REG_(?:DWORD|SZ|QWORD)\s+(.+)", line)
                if m:
                    key, val = m.group(1), m.group(2).strip()
                    # Convert hex DWORD values
                    if val.startswith("0x"):
                        try:
                            settings[key] = int(val, 16)
                        except ValueError:
                            settings[key] = val
                    else:
                        try:
                            settings[key] = int(val)
                        except ValueError:
                            try:
                                settings[key] = float(val)
                            except ValueError:
                                settings[key] = val
        except Exception:
            continue
    return settings


def _read_pimax_settings(config_path: Path) -> dict:
    """Read Pimax settings from config file + registry.

    Registry values override JSON values since Pimax Play
    typically stores live settings in the registry.
    """
    # Read JSON config
    data: dict = {}
    try:
        text = config_path.read_text(encoding="utf-8")
        parsed = _json.loads(text)
        if isinstance(parsed, dict):
            data = parsed
    except Exception:
        pass

    # Merge with registry values (registry wins for overlapping keys)
    reg = _read_pimax_registry()
    if reg:
        data.update(reg)

    return data


def _pimax_not_found_error() -> dict:
    """Return a helpful error when Pimax config is not found, with auto-search."""
    # Try to find it via PowerShell
    found_paths: list[str] = []
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             r'Get-ChildItem -Path "C:\","D:\" -Recurse -Filter "*.json" '
             r'-ErrorAction SilentlyContinue | '
             r'Where-Object { $_.DirectoryName -match "Pimax" } | '
             r'Select-Object -First 20 -ExpandProperty FullName'],
            capture_output=True, text=True, timeout=30,
        )
        if result.stdout.strip():
            found_paths = [p.strip() for p in result.stdout.strip().splitlines() if p.strip()]
    except Exception:
        pass

    msg = "Pimax-Konfigurationsdatei nicht automatisch gefunden."
    if found_paths:
        msg += " Folgende Pimax-Dateien wurden auf dem System gefunden:"
        return {
            "error": msg,
            "found_pimax_files": found_paths,
            "hint": (
                "Gib den Pfad zur richtigen Datei über config_path an. "
                "Suche nach einer JSON-Datei die Einstellungen wie "
                "brightness, refreshRate, fov etc. enthält."
            ),
        }
    else:
        return {
            "error": msg,
            "searched_locations": [str(p) for p in _PIMAX_CONFIG_CANDIDATES[:8]],
            "hint": (
                "Ist Pimax Play installiert? Gib den Installationsordner "
                "oder den Config-Pfad über config_path an."
            ),
        }


def _pimax_human_value(key: str, raw) -> str:
    """Convert a Pimax setting value to human-readable form."""
    defn = PIMAX_SETTING_DEFS.get(key)
    if defn and "values" in defn:
        return defn["values"].get(str(raw), str(raw))
    return str(raw)


@mcp.tool()
def analyze_pimax_settings(
    config_path: str = "",
) -> dict:
    """Analyze current Pimax VR headset settings with recommendations.

    Reads the Pimax configuration file, displays all settings in a table,
    and provides GPU-aware optimization recommendations.

    Present the results as markdown tables with columns:
    Setting | Aktuell | Status | Empfehlung | Grund

    Args:
        config_path: Path to Pimax config file or folder (auto-detected if empty).
    """
    cfg = _find_pimax_config(config_path)
    if cfg is None:
        return _pimax_not_found_error()

    settings = _read_pimax_settings(cfg)
    if not settings:
        return {"error": f"Konnte {cfg} nicht als Pimax-Config lesen."}

    # GPU info for recommendations
    gpu_name = "Unknown"
    vram_mb = 0
    try:
        pynvml.nvmlInit()
        try:
            handle = pynvml.nvmlDeviceGetHandleByIndex(0)
            gpu_name = pynvml.nvmlDeviceGetName(handle)
            if isinstance(gpu_name, bytes):
                gpu_name = gpu_name.decode()
            vram_mb = round(pynvml.nvmlDeviceGetMemoryInfo(handle).total / 1024 / 1024)
        finally:
            pynvml.nvmlShutdown()
    except Exception:
        pass

    vram_gb = vram_mb / 1024

    # Build recommendations based on VRAM
    tip_by_key: dict[str, dict] = {}
    for min_gb, tier, tips_list in PIMAX_VRAM_TIPS:
        if vram_gb >= min_gb:
            for key, suggestion, reason in tips_list:
                if key in settings:
                    current = settings[key]
                    tip_by_key[key] = {
                        "suggestion": suggestion,
                        "reason": reason,
                    }
            break

    # MSFS-specific Pimax tips (always apply)
    if settings.get("parallelProjection") in (False, "false", 0, "0"):
        tip_by_key["parallelProjection"] = {
            "suggestion": "An",
            "reason": "MSFS 2024 benötigt Parallel Projection für korrekte Darstellung in Pimax.",
        }
    refresh = settings.get("refreshRate", 90)
    if isinstance(refresh, (int, float)) and refresh > 90:
        tip_by_key["refreshRate"] = {
            "suggestion": "72 oder 90",
            "reason": "MSFS schafft selten über 45 FPS in VR. 72/90 Hz mit Smart Smoothing ist flüssiger.",
        }

    # Build table
    table: list[dict] = []
    for key, raw in settings.items():
        defn = PIMAX_SETTING_DEFS.get(key, {})
        label = defn.get("label", key)
        display = _pimax_human_value(key, raw)
        tip = tip_by_key.get(key)

        table.append({
            "setting": label,
            "current": display,
            "status": "CHANGE" if tip else "OK",
            "recommendation": tip["suggestion"] if tip else "-",
            "reason": tip["reason"] if tip else "",
        })

    changes_needed = len(tip_by_key)

    return {
        "config_file": str(cfg),
        "gpu": gpu_name,
        "vram_gb": round(vram_gb, 1),
        "pimax_settings": table,
        "summary": {
            "total_settings": len(table),
            "changes_recommended": changes_needed,
            "verdict": (
                "Pimax-Einstellungen sind optimal!"
                if changes_needed == 0
                else f"{changes_needed} Einstellung(en) sollten angepasst werden."
            ),
        },
        "display_instructions": (
            "IMPORTANT: Display ALL rows as a markdown table with columns: "
            "Einstellung | Aktuell | Status | Empfehlung | Grund. "
            "Use green checkmarks for OK, red X for CHANGE. Show EVERY setting. "
            "Answer in German."
        ),
    }


@mcp.tool()
def diagnose_pimax() -> dict:
    """Full diagnostic of all Pimax config files, registry entries and processes.

    Use this to find WHERE Pimax stores its settings. Shows:
    - All Pimax-related JSON/config files and their contents
    - All Pimax registry entries with values
    - Running Pimax processes
    - Pimax install locations

    Call this when Pimax settings appear incorrect to debug.
    """
    result: dict = {"sources": []}

    # 1. Find ALL Pimax config files
    config_files: list[dict] = []
    search_roots = [
        Path(os.environ.get("LOCALAPPDATA", "")),
        Path(os.environ.get("APPDATA", "")),
        Path(os.environ.get("PROGRAMDATA", "")),
        Path(os.environ.get("USERPROFILE", "")),
    ]
    for root in search_roots:
        if not root.exists():
            continue
        try:
            for p in root.rglob("*"):
                if not p.is_file():
                    continue
                if "pimax" not in str(p).lower():
                    continue
                if p.suffix.lower() in (".json", ".cfg", ".ini", ".xml", ".config"):
                    entry: dict = {"path": str(p), "size": p.stat().st_size}
                    if p.suffix.lower() == ".json" and p.stat().st_size < 500_000:
                        try:
                            data = _json.loads(p.read_text(encoding="utf-8"))
                            if isinstance(data, dict):
                                # Show all keys with brightness/contrast/color
                                relevant = {
                                    k: v for k, v in data.items()
                                    if any(term in k.lower() for term in
                                           ("bright", "contrast", "color", "render",
                                            "refresh", "fov", "smooth", "ffr", "ipd",
                                            "parallel", "resolution", "quality"))
                                }
                                entry["relevant_settings"] = relevant
                                entry["all_keys"] = list(data.keys())
                        except Exception:
                            pass
                    config_files.append(entry)
        except PermissionError:
            continue

    # Also check Program Files
    for pf in [r"C:\Program Files\Pimax", r"C:\Program Files (x86)\Pimax",
               r"D:\Program Files\Pimax", r"D:\Pimax"]:
        pf_path = Path(pf)
        if not pf_path.exists():
            continue
        try:
            for p in pf_path.rglob("*"):
                if not p.is_file():
                    continue
                if p.suffix.lower() in (".json", ".cfg", ".ini", ".xml", ".config", ".db", ".sqlite"):
                    entry = {"path": str(p), "size": p.stat().st_size}
                    if p.suffix.lower() == ".json" and p.stat().st_size < 500_000:
                        try:
                            data = _json.loads(p.read_text(encoding="utf-8"))
                            if isinstance(data, dict):
                                relevant = {
                                    k: v for k, v in data.items()
                                    if any(term in k.lower() for term in
                                           ("bright", "contrast", "color", "render",
                                            "refresh", "fov", "smooth", "ffr", "ipd"))
                                }
                                entry["relevant_settings"] = relevant
                                entry["all_keys"] = list(data.keys())
                        except Exception:
                            pass
                    config_files.append(entry)
        except PermissionError:
            continue

    result["config_files"] = config_files

    # 2. Read ALL Pimax registry entries
    registry: dict[str, dict] = {}
    reg_roots = [
        r"HKCU\SOFTWARE\Pimax",
        r"HKLM\SOFTWARE\Pimax",
        r"HKCU\SOFTWARE\PimaxPlay",
        r"HKLM\SOFTWARE\PimaxPlay",
        r"HKCU\SOFTWARE\Pimax Play",
    ]
    for rp in reg_roots:
        try:
            r = _reg(["query", rp, "/s"], timeout=10)
            if r.returncode != 0:
                continue
            current_subkey = rp
            entries: dict = {}
            for line in r.stdout.splitlines():
                line_stripped = line.strip()
                if not line_stripped:
                    continue
                if line_stripped.startswith("HKEY_"):
                    current_subkey = line_stripped
                    if current_subkey not in registry:
                        registry[current_subkey] = {}
                else:
                    m = re.match(r"(\S+)\s+(REG_\S+)\s+(.*)", line_stripped)
                    if m:
                        name, reg_type, val = m.group(1), m.group(2), m.group(3).strip()
                        display_val = val
                        if reg_type == "REG_DWORD" and val.startswith("0x"):
                            try:
                                display_val = f"{val} ({int(val, 16)})"
                            except ValueError:
                                pass
                        registry.setdefault(current_subkey, {})[name] = {
                            "type": reg_type,
                            "value": display_val,
                        }
        except Exception:
            continue

    result["registry"] = registry

    # 3. Running Pimax processes
    processes: list[str] = []
    try:
        r = subprocess.run(
            ["tasklist", "/fo", "csv", "/nh"],
            capture_output=True, text=True, timeout=10,
        )
        for line in r.stdout.splitlines():
            if "pimax" in line.lower() or "pitool" in line.lower():
                processes.append(line.strip().split(",")[0].strip('"'))
    except Exception:
        pass
    result["running_processes"] = processes

    # 4. Currently detected config
    cfg = _find_pimax_config()
    result["auto_detected_config"] = str(cfg) if cfg else None

    result["display_instructions"] = (
        "Show ALL config files with their relevant settings. "
        "Show ALL registry entries. Highlight any keys containing "
        "'brightness', 'contrast', or 'color'. "
        "This helps identify which source has the correct values."
    )

    return result


@mcp.tool()
def optimize_pimax_settings(
    preset: Literal["Performance", "Balanced", "Quality", "MSFS_VR_Optimized"] = "Balanced",
    custom_overrides: dict | None = None,
    config_path: str = "",
    restart_service: bool = True,
    force_restart: bool = False,
    dry_run: bool = False,
) -> dict:
    """Optimize Pimax VR headset settings with a preset.

    Saves the current config as backup before applying changes.
    Writes to BOTH the JSON config and the Pimax registry hive — Pimax Play
    reads its live settings from the registry, so JSON-only writes are
    silently ignored. Restarts Pimax Play afterwards so the changes become
    visible (skipped during active VR sessions unless force_restart=True).

    Available presets:
    - Performance: 1.0x, 90Hz, Normal FOV, FFR Medium, Smart Smoothing An
    - Balanced: 1.25x, 90Hz, Normal FOV, FFR Low, Smart Smoothing An
    - Quality: 1.5x, 90Hz, Large FOV, FFR Aus, Smart Smoothing Aus
    - MSFS_VR_Optimized: Speziell für Flight Sim — 1.0x, 72Hz, Parallel Projection An

    Args:
        preset: Which preset to apply.
        custom_overrides: Optional extra settings as {"key": value} applied after the preset.
        config_path: Path to Pimax config file (auto-detected if empty).
        restart_service: If True, restarts Pimax Play so changes take effect live.
        force_restart: If True, restarts Pimax even during a running VR session
                       (interrupts it). Use only with explicit user consent.
        dry_run: If True, only return what WOULD change. Does not write
                 anything, does not create a backup, does not restart Pimax.
    """
    cfg = _find_pimax_config(config_path)
    if cfg is None:
        return _pimax_not_found_error()

    current = _read_pimax_settings(cfg)
    if not current:
        return {"error": f"Konnte {cfg} nicht lesen."}

    overrides = dict(PIMAX_PRESETS.get(preset, {}))
    if custom_overrides:
        overrides.update(custom_overrides)

    diff: dict[str, dict] = {}
    applied: dict = {}
    for key, value in overrides.items():
        if key in current:
            old_val = current[key]
            if old_val != value:
                diff[key] = {"old": old_val, "new": value}
            applied[key] = value

    if dry_run:
        logger.info("optimize_pimax_settings DRY RUN preset=%s diff=%s", preset, diff)
        return {
            "status": "dry_run",
            "preset": preset,
            "config_file": str(cfg),
            "would_apply": {k: str(v) for k, v in applied.items()},
            "diff": diff,
            "not_found_in_config": [k for k in overrides if k not in applied],
            "note": "Nichts wurde geschrieben. Setze dry_run=False um wirklich anzuwenden.",
        }

    backup_file = _pimax_create_backup(cfg, current)

    # Apply changes to in-memory dict, then write JSON
    for key, value in applied.items():
        current[key] = value
    cfg.write_text(_json.dumps(current, indent=2, ensure_ascii=False), encoding="utf-8")

    # Write to registry — without this, Pimax Play silently ignores JSON changes
    registry_audit = _apply_pimax_settings_to_registry(applied, verify=True)

    logger.info("optimize_pimax_settings applied preset=%s changes=%s", preset, list(applied.keys()))
    result = {
        "status": "ok",
        "preset": preset,
        "config_file": str(cfg),
        "backup": str(backup_file),
        "applied": {k: str(v) for k, v in applied.items()},
        "diff": diff,
        "registry_written": registry_audit,
        "not_found_in_config": [k for k in overrides if k not in applied],
    }

    _pimax_apply_restart(result, restart_service, force_restart)
    return result


@mcp.tool()
def restore_pimax_settings(
    config_path: str = "",
) -> dict:
    """Restore Pimax settings from the latest backup.

    Args:
        config_path: Path to Pimax config file (auto-detected if empty).
    """
    cfg = _find_pimax_config(config_path)
    if cfg is None:
        return _pimax_not_found_error()

    backup_dir = cfg.parent / "pimax_backups"
    latest = backup_dir / "pimax_config_latest.json"

    if not latest.exists():
        # List available backups
        if backup_dir.exists():
            backups = sorted(backup_dir.glob("pimax_config_*.json"), reverse=True)
            return {
                "error": "Kein 'latest' Backup gefunden.",
                "available_backups": [f.name for f in backups[:10]],
            }
        return {"error": "Kein Pimax-Backup vorhanden."}

    backup_data = latest.read_text(encoding="utf-8")
    cfg.write_text(backup_data, encoding="utf-8")

    return {
        "status": "ok",
        "restored_from": str(latest),
        "config_file": str(cfg),
        "hint": "Pimax Play neu starten, damit die Änderungen wirksam werden.",
    }


# ── Shared helpers for Pimax write/restart, used by set_pimax_setting,
#    optimize_pimax_settings, and improve_image_clarity ──────────────────

_PIMAX_BACKUP_KEEP = 20
_MSFS_BACKUP_KEEP = 20
_MSFS_HISTORY_KEEP = 30


def _prune_pimax_backups(backup_dir: Path, keep: int = _PIMAX_BACKUP_KEEP) -> int:
    """Delete oldest pimax_config_*.json backups, keeping the newest `keep`.

    Never deletes pimax_config_latest.json. Returns number of files removed.
    """
    if not backup_dir.exists():
        return 0
    backups = sorted(
        (
            p for p in backup_dir.glob("pimax_config_*.json")
            if p.name != "pimax_config_latest.json"
        ),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    deleted = 0
    for old in backups[keep:]:
        try:
            old.unlink()
            deleted += 1
        except Exception as exc:
            logger.warning("Could not prune backup %s: %s", old, exc)
    if deleted:
        logger.info("Pruned %d old Pimax backup(s) from %s", deleted, backup_dir)
    return deleted


def _pimax_create_backup(cfg: Path, current_state: dict) -> Path:
    """Save a timestamped backup of `current_state` and update 'latest'.

    Auto-prunes old backups so the folder doesn't grow unbounded.
    Returns the path to the new timestamped backup file.
    """
    backup_dir = cfg.parent / "pimax_backups"
    backup_dir.mkdir(exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = backup_dir / f"pimax_config_{ts}.json"
    payload = _json.dumps(current_state, indent=2, ensure_ascii=False)
    backup_file.write_text(payload, encoding="utf-8")
    latest = backup_dir / "pimax_config_latest.json"
    latest.write_text(payload, encoding="utf-8")
    _prune_pimax_backups(backup_dir, keep=_PIMAX_BACKUP_KEEP)
    logger.info("Pimax backup created: %s", backup_file.name)
    return backup_file


def _prune_msfs_backups(cfg_path: Path, keep: int = _MSFS_BACKUP_KEEP) -> int:
    """Delete oldest UserCfg_v*.opt backups, keeping the newest `keep`.

    Never deletes UserCfg_latest.opt. Returns number of files removed.
    """
    backup_dir = cfg_path.parent / "UserCfg_backups"
    if not backup_dir.exists():
        return 0
    backups = sorted(
        (
            p for p in backup_dir.glob("UserCfg_v*.opt")
            if p.name != "UserCfg_latest.opt"
        ),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    deleted = 0
    for old in backups[keep:]:
        try:
            old.unlink()
            deleted += 1
        except Exception as exc:
            logger.warning("Could not prune backup %s: %s", old, exc)
    if deleted:
        logger.info("Pruned %d old MSFS backup(s)", deleted)
    return deleted


def _wait_for_file_unlocked(path: Path, timeout_s: int = 10) -> bool:
    """Wait until `path` can be opened for writing (no other process holds it).

    Returns True once the file can be opened r+b, False on timeout.
    Replaces blind sleep() calls after killing processes that held the file.
    """
    import time as _time
    deadline = _time.time() + timeout_s
    while _time.time() < deadline:
        try:
            with open(path, "r+b"):
                return True
        except (PermissionError, OSError):
            _time.sleep(0.3)
    logger.warning("File still locked after %ds: %s", timeout_s, path)
    return False


def _wait_for_process_exit(name: str, timeout_s: int = 30) -> bool:
    """Wait until the named process is no longer running. Returns True if exited."""
    import time as _time
    deadline = _time.time() + timeout_s
    while _time.time() < deadline:
        if not _is_running(name):
            return True
        _time.sleep(0.5)
    return False


def _apply_pimax_settings_to_registry(
    updates: dict[str, object],
    verify: bool = True,
) -> dict[str, dict]:
    """Write key→value pairs into the Pimax registry hives.

    Pimax Play reads live settings from the registry, so JSON-only writes
    are silently ignored. Only writes a key if it already exists in one of
    the candidate hives.

    If `verify` is True (default), reads each value back after writing and
    flags any mismatches in the result.

    Returns: dict of {key: {"hive": ..., "written_value": ..., "verified": bool, "actual": ...}}
    """
    result: dict[str, dict] = {}
    reg_paths_to_try = [
        r"HKCU\SOFTWARE\Pimax",
        r"HKCU\SOFTWARE\Pimax\PimaxPlay",
    ]
    for mk, new_val in updates.items():
        for rp in reg_paths_to_try:
            try:
                check = _reg(["query", rp, "/v", mk], timeout=5)
                if check.returncode != 0:
                    continue
                if isinstance(new_val, bool):
                    reg_val = "1" if new_val else "0"
                    reg_type = "REG_DWORD"
                elif isinstance(new_val, int):
                    reg_val = str(new_val)
                    reg_type = "REG_DWORD"
                elif isinstance(new_val, float):
                    reg_val = str(new_val)
                    reg_type = "REG_SZ"
                else:
                    reg_val = str(new_val)
                    reg_type = "REG_SZ"
                add_res = _reg_add(rp, mk, reg_type, reg_val, timeout=5)
                entry: dict = {
                    "hive": rp,
                    "written_value": reg_val,
                    "reg_type": reg_type,
                    "verified": None,
                }
                if verify:
                    readback = _reg(["query", rp, "/v", mk], timeout=5)
                    actual: str | None = None
                    if readback.returncode == 0:
                        m = re.search(
                            rf"{re.escape(mk)}\s+REG_(?:DWORD|SZ|QWORD)\s+(\S+)",
                            readback.stdout,
                        )
                        if m:
                            raw = m.group(1).strip()
                            if raw.startswith("0x"):
                                try:
                                    actual = str(int(raw, 16))
                                except ValueError:
                                    actual = raw
                            else:
                                actual = raw
                    entry["actual"] = actual
                    entry["verified"] = (actual == reg_val)
                    if not entry["verified"]:
                        logger.warning(
                            "Pimax registry write NOT verified: %s in %s "
                            "(wrote %r, read %r). Could be permission or type mismatch.",
                            mk, rp, reg_val, actual,
                        )
                if add_res.returncode != 0:
                    logger.warning(
                        "reg add for %s returned %d: %s",
                        mk, add_res.returncode, add_res.stderr.strip(),
                    )
                else:
                    logger.info("Pimax registry: %s = %r in %s", mk, reg_val, rp)
                result[mk] = entry
                break
            except Exception as exc:
                logger.exception("Failed to write Pimax registry key %s: %s", mk, exc)
                continue
    return result


def _pimax_apply_restart(
    result: dict,
    restart_service: bool,
    force_restart: bool = False,
) -> None:
    """Restart Pimax Play to apply changes; mutates `result` with status messages.

    By default Pimax is NOT restarted while a VR session is active because
    that crashes the session. Pass force_restart=True to override.
    """
    import time

    if not restart_service:
        result["hinweis"] = (
            "Einstellungen gespeichert (ohne Neustart). "
            "Setze restart_service=True um sofort zu übernehmen."
        )
        return

    vr_game_running = _is_msfs_running() or _is_running("vrserver.exe")
    pimax_is_running = any(_is_running(name) for name in _PIMAX_EXE_NAMES)

    if pimax_is_running and vr_game_running and not force_restart:
        result["hinweis"] = (
            "Einstellungen in Config/Registry gespeichert. "
            "Pimax Play wurde NICHT neu gestartet (VR-Session läuft — "
            "Neustart würde sie crashen). Damit die Änderung sichtbar wird: "
            "VR-Spiel beenden und 'Pimax neu starten' sagen, ODER diesen "
            "Aufruf mit force_restart=True wiederholen (unterbricht die VR-Session)."
        )
        return

    if not pimax_is_running:
        result["hinweis"] = (
            "Einstellungen gespeichert. Pimax Play läuft nicht — "
            "Änderungen werden beim nächsten Start übernommen."
        )
        return

    restarted = False
    for exe in _PIMAX_EXE_NAMES:
        if not _is_running(exe):
            continue
        try:
            subprocess.run(
                ["taskkill", "/IM", exe, "/F"],
                capture_output=True, timeout=10,
            )
            time.sleep(2)
            pimax_exe = _find_pimax()
            if pimax_exe:
                subprocess.Popen(
                    [pimax_exe],
                    creationflags=getattr(subprocess, "DETACHED_PROCESS", 0),
                )
                restarted = True
        except Exception:
            pass
        break

    if restarted:
        if force_restart and vr_game_running:
            result["service_restart"] = (
                "Pimax Play wurde neu gestartet — Änderungen aktiv "
                "(VR-Session wurde unterbrochen wegen force_restart=True)."
            )
        else:
            result["service_restart"] = (
                "Pimax Play wurde neu gestartet — Änderungen sind sofort aktiv."
            )
    else:
        result["hinweis"] = "Pimax Play muss ggf. manuell neu gestartet werden."


@mcp.tool()
def set_pimax_setting(
    setting: str,
    value: str,
    config_path: str = "",
    restart_service: bool = True,
    force_restart: bool = False,
    dry_run: bool = False,
) -> dict:
    """Set a single Pimax setting directly (e.g. brightness, contrast, refresh rate).

    Use this when the user asks to change a specific Pimax setting like
    "Helligkeit erhöhen", "Contrast auf 5", "Refresh Rate auf 72" etc.

    Creates a backup before making changes. Optionally restarts the Pimax
    service so the change takes effect immediately.

    WICHTIG: Wenn der User in VR ist und sagt "ich sehe keine Änderung",
    dann liegt das daran dass restart_service während VR-Sessions blockiert
    ist (sonst Crash). Schlage dann vor: VR beenden und Pimax neu starten,
    ODER rufe dieses Tool nochmal mit force_restart=True auf.

    Für "klares Bild" / "schärfer" / "sharpness": dieses Tool ändert nur
    eine einzelne Einstellung. Verwende stattdessen improve_image_clarity,
    das mehrere Schärfe-relevante Einstellungen kombiniert.

    Setze dry_run=True um vorher zu sehen, was sich ändern würde,
    ohne tatsächlich zu schreiben.

    Known settings and typical values:
    - brightness / helligkeit: Integer slider value (as shown in Pimax Play). Sets both eyes.
    - contrast / kontrast: Integer slider value. Sets both eyes.
    - saturation / sättigung: Integer slider value. Sets both eyes.
    - renderResolution: 0.5-2.0 (e.g. 1.0, 1.25, 1.5)
    - refreshRate: 72, 90, 120, 144
    - fov / fovLevel: "small"/"normal"/"large" or 0-3
    - smartSmoothing: true/false
    - compulsorySmoothing: true/false
    - ffrLevel: 0 (off) - 4 (ultra). Niedriger = klareres Bild in der Peripherie.
    - parallelProjection: true/false
    - ipd: 55-75 (mm)
    - eyeTracking: true/false

    Supports relative values: "+10", "-20" to adjust from current.

    Args:
        setting: Setting key (e.g. "brightness", "contrast", "refreshRate").
        value: New value as string. Will be auto-converted to the correct type.
        config_path: Path to Pimax config (auto-detected if empty).
        restart_service: If True, restarts Pimax Play so changes take effect live.
        force_restart: If True, restarts Pimax even if a VR game is running
                       (this WILL interrupt the VR session). Use only when the
                       user explicitly accepts that.
        dry_run: If True, only return what WOULD change. No write, no backup, no restart.
    """
    cfg = _find_pimax_config(config_path)
    if cfg is None:
        return _pimax_not_found_error()

    current = _read_pimax_settings(cfg)
    if not current:
        return {"error": f"Konnte {cfg} nicht lesen."}

    # Normalize setting name — German aliases → search keywords
    # These map to a keyword used to find matching keys in the config
    _ALIASES = {
        "helligkeit": "brightness",
        "heller": "brightness",
        "dunkler": "brightness",
        "kontrast": "contrast",
        "sättigung": "saturation",
        "auflösung": "renderresolution",
        "render": "renderresolution",
        "render_resolution": "renderresolution",
        "render resolution": "renderresolution",
        "wiederholrate": "refreshrate",
        "refresh": "refreshrate",
        "refresh_rate": "refreshrate",
        "refresh rate": "refreshrate",
        "hz": "refreshrate",
        "sichtfeld": "fov",
        "field_of_view": "fov",
        "field of view": "fov",
        "fov_level": "fovlevel",
        "fov level": "fovlevel",
        "smart_smoothing": "smartsmoothing",
        "smart smoothing": "smartsmoothing",
        "smoothing": "smartsmoothing",
        "glättung": "smartsmoothing",
        "compulsory_smoothing": "compulsorysmoothing",
        "compulsory smoothing": "compulsorysmoothing",
        "ffr": "ffrlevel",
        "ffr_level": "ffrlevel",
        "ffr level": "ffrlevel",
        "foveated": "ffr",
        "foveated_rendering": "ffr",
        "parallel_projection": "parallelProjection",
        "parallel projection": "parallelProjection",
        "pp": "parallelProjection",
        "eye_tracking": "eyetracking",
        "eye tracking": "eyetracking",
        "augentracking": "eyetracking",
        "dynamic_foveated": "dynamicfoveated",
        "dynamic foveated": "dynamicfoveated",
        "dfr": "dynamicfoveated",
        "ipd_offset": "ipdoffset",
        "ipd offset": "ipdoffset",
        # Schärfe / Klarheit — höhere Render-Auflösung = klareres Bild
        "klares bild": "renderresolution",
        "klareres bild": "renderresolution",
        "schärfer": "renderresolution",
        "schaerfer": "renderresolution",
        "schärfe": "renderresolution",
        "schaerfe": "renderresolution",
        "sharper": "renderresolution",
        "sharpness": "renderresolution",
        "klarheit": "renderresolution",
        "supersampling": "renderresolution",
    }
    search_term = _ALIASES.get(setting.lower().strip(), setting.strip().lower())

    # Find ALL matching keys in config (e.g. "brightness" matches
    # "piplay_color_brightness_0" AND "piplay_color_brightness_1")
    matching_keys: list[str] = []

    # 1. Exact match
    if setting.strip() in current:
        matching_keys = [setting.strip()]
    else:
        # 2. Case-insensitive exact match
        for k in current:
            if k.lower() == search_term:
                matching_keys = [k]
                break

    # 3. Substring/fuzzy match — find all keys containing the search term
    if not matching_keys:
        for k in current:
            if search_term in k.lower():
                matching_keys.append(k)

    # 4. If still nothing, try the original setting name as substring
    if not matching_keys:
        for k in current:
            if setting.strip().lower() in k.lower():
                matching_keys.append(k)

    # 5. If still nothing, add as new key
    if not matching_keys:
        matching_keys = [setting.strip()]

    actual_key = matching_keys[0]  # primary key for display

    # Auto-convert value to the correct type
    parsed_value: object
    v_lower = value.strip().lower()
    if v_lower in ("true", "an", "ein", "on", "ja", "yes", "aktiv"):
        parsed_value = True
    elif v_lower in ("false", "aus", "off", "nein", "no", "deaktiv"):
        parsed_value = False
    else:
        # Try int, then float, else keep string
        try:
            parsed_value = int(value)
        except ValueError:
            try:
                parsed_value = float(value)
            except ValueError:
                parsed_value = value.strip()

    # Relative adjustments: "+10", "-20"
    is_relative = False
    v_stripped = value.strip()
    if v_stripped.startswith("+") or (v_stripped.startswith("-") and len(v_stripped) > 1):
        try:
            delta = float(v_stripped)
            is_relative = True
        except ValueError:
            pass

    # Compute new values BEFORE writing — supports dry_run
    applied: dict[str, dict] = {}
    new_values: dict[str, object] = {}
    for mk in matching_keys:
        old_val = current.get(mk, 0)
        if is_relative and isinstance(old_val, (int, float)):
            new_val = type(old_val)(old_val + delta)
        else:
            new_val = parsed_value
        # Don't clamp — the actual range depends on Pimax version/setting
        new_values[mk] = new_val
        defn = PIMAX_SETTING_DEFS.get(mk, {})
        applied[mk] = {
            "label": defn.get("label", mk),
            "old": _pimax_human_value(mk, old_val),
            "new": _pimax_human_value(mk, new_val),
            "old_raw": old_val,
            "new_raw": new_val,
        }

    if dry_run:
        logger.info("set_pimax_setting DRY RUN setting=%r value=%r → %s", setting, value, applied)
        return {
            "status": "dry_run",
            "setting": setting,
            "matched_keys": matching_keys,
            "would_change": {
                k: {"label": v["label"], "old": v["old"], "new": v["new"]}
                for k, v in applied.items()
            },
            "config_file": str(cfg),
            "note": "Nichts wurde geschrieben. Setze dry_run=False um wirklich anzuwenden.",
        }

    # Persist to in-memory dict
    for mk, nv in new_values.items():
        current[mk] = nv

    backup_file = _pimax_create_backup(cfg, current)

    # Write JSON config — only update keys that already exist in JSON
    try:
        json_data = _json.loads(cfg.read_text(encoding="utf-8"))
        if isinstance(json_data, dict):
            for mk in matching_keys:
                if mk in json_data:
                    json_data[mk] = current[mk]
            cfg.write_text(_json.dumps(json_data, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:
        logger.warning("Could not update Pimax JSON config %s: %s", cfg, exc)

    # Write to registry — Pimax Play reads live values from registry, not JSON
    registry_audit = _apply_pimax_settings_to_registry(
        {mk: current[mk] for mk in matching_keys}, verify=True,
    )

    logger.info(
        "set_pimax_setting setting=%r value=%r matched=%s",
        setting, value, matching_keys,
    )

    # Build result
    if len(applied) == 1:
        info = applied[actual_key]
        result = {
            "status": "ok",
            "setting": info["label"],
            "key": actual_key,
            "old_value": info["old"],
            "new_value": info["new"],
            "config_file": str(cfg),
            "backup": str(backup_file),
            "registry_audit": registry_audit,
        }
    else:
        result = {
            "status": "ok",
            "changed_keys": {
                k: {"label": v["label"], "old": v["old"], "new": v["new"]}
                for k, v in applied.items()
            },
            "config_file": str(cfg),
            "backup": str(backup_file),
            "registry_audit": registry_audit,
        }

    _pimax_apply_restart(result, restart_service, force_restart)
    return result


@mcp.tool()
def improve_image_clarity(
    target: Literal["pimax", "msfs", "both"] = "both",
    strength: Literal["mild", "medium", "strong"] = "medium",
    config_path: str = "",
    usercfg_path: str = "",
    auto_restart: bool = True,
    force_restart: bool = False,
    dry_run: bool = False,
) -> dict:
    """Make the VR image visibly clearer by adjusting multiple settings at once.

    USE THIS when the user says things like:
    - "klares Bild" / "klareres Bild"
    - "schärfer" / "Schärfe erhöhen"
    - "sharper image" / "improve clarity"
    - "ich sehe alles unscharf"
    - "kann man das schärfer machen?"

    Single-setting tools cannot fulfill those requests because clarity is
    not one slider — it's a combination of render resolution, foveated
    rendering, anti-aliasing mode, and texture/anisotropic quality.

    What this tool does:
        Pimax (writes JSON + registry, restarts Pimax Play unless VR-Session aktiv):
        - mild:    renderResolution=1.25, ffrLevel=1 (Low)
        - medium:  renderResolution=1.5,  ffrLevel=0 (Aus)
        - strong:  renderResolution=1.75, ffrLevel=0, smartSmoothing=False

        MSFS 2024 (kill MSFS → write UserCfg.opt → relaunch in VR):
        - mild:    DLSS=Quality, RenderScale=100
        - medium:  DLSS=Quality, RenderScale=110, AnisotropicFilter=16x, TextureRes=Ultra
        - strong:  DLSS=DLAA,    RenderScale=120, AnisotropicFilter=16x, TextureRes=Ultra

    Higher strength costs more GPU/VRAM. Default 'medium' is a good
    starting point for an RTX 4080/5080-class GPU.

    Args:
        target: "pimax", "msfs", or "both" (default).
        strength: "mild", "medium" (default), or "strong".
        config_path: Pimax config path (auto-detected if empty).
        usercfg_path: MSFS UserCfg.opt path (auto-detected if empty).
        auto_restart: Restart Pimax / MSFS so changes become visible.
        force_restart: For Pimax: restart even during an active VR session
                       (interrupts it). Only with explicit user consent.
        dry_run: If True, only return what WOULD change for both targets.
                 No writes, no kills, no restart.
    """
    return _apply_combo_profile(
        domain="clarity",
        target=target,
        strength=strength,
        pimax_profiles={
            "mild":   {"renderResolution": 1.25, "ffrLevel": 1},
            "medium": {"renderResolution": 1.5,  "ffrLevel": 0},
            "strong": {"renderResolution": 1.75, "ffrLevel": 0, "smartSmoothing": False},
        },
        msfs_profiles={
            "mild": {
                "Video.DLSS": "2",
                "Video.RenderScale": "100",
            },
            "medium": {
                "Video.DLSS": "2",
                "Video.RenderScale": "110",
                "GraphicsVR.AnisotropicFilter": "4",
                "GraphicsVR.TextureResolution": "3",
            },
            "strong": {
                "Video.DLSS": "1",
                "Video.RenderScale": "120",
                "GraphicsVR.AnisotropicFilter": "4",
                "GraphicsVR.TextureResolution": "3",
            },
        },
        config_path=config_path,
        usercfg_path=usercfg_path,
        auto_restart=auto_restart,
        force_restart=force_restart,
        dry_run=dry_run,
    )


def _apply_combo_profile(
    domain: str,
    target: str,
    strength: str,
    pimax_profiles: dict[str, dict],
    msfs_profiles: dict[str, dict[str, str]],
    config_path: str,
    usercfg_path: str,
    auto_restart: bool,
    force_restart: bool,
    dry_run: bool,
) -> dict:
    """Shared engine for multi-setting combo tools (clarity, performance, …).

    Applies a Pimax profile + MSFS profile in one coordinated step:
    backup → write JSON → write registry (Pimax only) → kill+relaunch MSFS.
    """
    import time

    out: dict = {
        "domain": domain,
        "target": target,
        "strength": strength,
        "dry_run": dry_run,
    }

    # ── Pimax ───────────────────────────────────────────────────────────────
    if target in ("pimax", "both"):
        cfg = _find_pimax_config(config_path)
        if cfg is None:
            out["pimax"] = _pimax_not_found_error()
        else:
            current = _read_pimax_settings(cfg)
            if not current:
                out["pimax"] = {"error": f"Konnte {cfg} nicht lesen."}
            else:
                profile = pimax_profiles[strength]
                diff = {
                    k: {"old": current.get(k, "<not set>"), "new": v}
                    for k, v in profile.items()
                    if current.get(k) != v
                }
                if dry_run:
                    out["pimax"] = {
                        "status": "dry_run",
                        "would_apply": {k: str(v) for k, v in profile.items()},
                        "diff": diff,
                        "config_file": str(cfg),
                    }
                else:
                    backup_file = _pimax_create_backup(cfg, current)
                    for k, v in profile.items():
                        current[k] = v
                    cfg.write_text(
                        _json.dumps(current, indent=2, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    registry_audit = _apply_pimax_settings_to_registry(profile, verify=True)
                    logger.info("%s combo Pimax: applied %s", domain, list(profile.keys()))
                    pimax_result: dict = {
                        "status": "ok",
                        "applied": {k: str(v) for k, v in profile.items()},
                        "diff": diff,
                        "registry_audit": registry_audit,
                        "config_file": str(cfg),
                        "backup": str(backup_file),
                    }
                    _pimax_apply_restart(pimax_result, auto_restart, force_restart)
                    out["pimax"] = pimax_result

    # ── MSFS ────────────────────────────────────────────────────────────────
    if target in ("msfs", "both"):
        cfg_path = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH
        if not cfg_path.exists():
            out["msfs"] = {
                "error": f"UserCfg.opt nicht gefunden: {cfg_path}",
                "tipp": "Gib usercfg_path an oder starte MSFS einmal damit die Datei angelegt wird.",
            }
        else:
            updates = msfs_profiles[strength]
            current_settings = _read_current_settings(cfg_path)
            diff = {
                k: {"old": current_settings.get(k, "<not set>"), "new": v}
                for k, v in updates.items()
                if current_settings.get(k) != v
            }

            if dry_run:
                out["msfs"] = {
                    "status": "dry_run",
                    "would_apply": updates,
                    "diff": diff,
                    "config_file": str(cfg_path),
                }
            else:
                # Kill MSFS first — sonst überschreibt MSFS unsere Änderungen beim Beenden
                msfs_was_running = _is_msfs_running()
                if msfs_was_running:
                    _kill_msfs(timeout_s=30)
                    # Wait for the file lock to clear instead of blind sleep
                    _wait_for_file_unlocked(cfg_path, timeout_s=10)

                saved_version = _snapshot(
                    cfg_path, f"Before {domain} combo strength={strength}"
                )
                _prune_msfs_backups(cfg_path)

                text = cfg_path.read_text(encoding="utf-8")
                entries = _parse_usercfg(text)
                new_entries, not_applied = _apply_overrides(entries, updates)
                cfg_path.write_text(_entries_to_text(new_entries), encoding="utf-8")

                verify = _read_current_settings(cfg_path)
                applied_settings = {k: v for k, v in updates.items() if k not in not_applied}
                verified = {k: verify.get(k) == v for k, v in applied_settings.items()}
                logger.info(
                    "%s combo MSFS: applied %s, not_applied=%s",
                    domain, list(applied_settings.keys()), list(not_applied),
                )

                msfs_result: dict = {
                    "status": "ok",
                    "applied": applied_settings,
                    "not_applied": list(not_applied),
                    "diff": diff,
                    "verified": verified,
                    "saved_version": saved_version,
                    "config_file": str(cfg_path),
                }

                if auto_restart and msfs_was_running:
                    time.sleep(2)
                    vr_result = launch_msfs_vr(force_restart=False)
                    msfs_result["neustart"] = (
                        "MSFS in VR neu gestartet — Änderungen sind sofort aktiv."
                    )
                    msfs_result["vr_launch"] = vr_result
                elif msfs_was_running:
                    msfs_result["hinweis"] = (
                        "MSFS wurde beendet aber nicht neu gestartet (auto_restart=False)."
                    )
                else:
                    msfs_result["hinweis"] = "Wird beim nächsten MSFS-Start übernommen."

                out["msfs"] = msfs_result

    return out


@mcp.tool()
def improve_performance(
    target: Literal["pimax", "msfs", "both"] = "both",
    strength: Literal["mild", "medium", "strong"] = "medium",
    config_path: str = "",
    usercfg_path: str = "",
    auto_restart: bool = True,
    force_restart: bool = False,
    dry_run: bool = False,
) -> dict:
    """Improve VR FPS / reduce stutter by adjusting multiple settings at once.

    USE THIS when the user says things like:
    - "mehr FPS" / "bessere Performance"
    - "weniger Ruckler" / "weniger Stuttering"
    - "schneller" / "smoother"
    - "ist zu ruckelig"
    - "improve framerate" / "reduce stutter"

    The opposite of improve_image_clarity — trades visual quality for
    headroom. Use mild first if the user just wants a small boost,
    strong when they're really struggling.

    What this tool does:
        Pimax:
        - mild:    renderResolution=1.0,  ffrLevel=2 (Medium), smartSmoothing=True
        - medium:  renderResolution=0.9,  ffrLevel=3 (High),   smartSmoothing=True
        - strong:  renderResolution=0.75, ffrLevel=4 (Ultra),  smartSmoothing=True,
                   compulsorySmoothing=True

        MSFS 2024:
        - mild:    DLSS=Performance, RenderScale=90, Wolken=Medium
        - medium:  DLSS=Performance, RenderScale=80, Wolken=Low, Schatten=Medium,
                   Reflexionen=Low, Motion Blur=Off
        - strong:  DLSS=Ultra Performance, RenderScale=70, Wolken=Low, Schatten=Low,
                   Reflexionen=Low, SSContact=Off, Motion Blur=Off, RayTracing=Off

    Args:
        target: "pimax", "msfs", or "both" (default).
        strength: "mild", "medium" (default), or "strong".
        config_path: Pimax config path (auto-detected if empty).
        usercfg_path: MSFS UserCfg.opt path (auto-detected if empty).
        auto_restart: Restart Pimax / MSFS so changes become visible.
        force_restart: For Pimax: restart even during an active VR session.
        dry_run: If True, only show what would change without writing.
    """
    return _apply_combo_profile(
        domain="performance",
        target=target,
        strength=strength,
        pimax_profiles={
            "mild":   {"renderResolution": 1.0,  "ffrLevel": 2, "smartSmoothing": True},
            "medium": {"renderResolution": 0.9,  "ffrLevel": 3, "smartSmoothing": True},
            "strong": {
                "renderResolution": 0.75, "ffrLevel": 4,
                "smartSmoothing": True, "compulsorySmoothing": True,
            },
        },
        msfs_profiles={
            "mild": {
                "Video.DLSS": "4",
                "Video.RenderScale": "90",
                "GraphicsVR.CloudsQuality": "1",
            },
            "medium": {
                "Video.DLSS": "4",
                "Video.RenderScale": "80",
                "GraphicsVR.CloudsQuality": "0",
                "GraphicsVR.ShadowQuality": "1",
                "GraphicsVR.Reflections": "0",
                "GraphicsVR.MotionBlur": "0",
            },
            "strong": {
                "Video.DLSS": "5",
                "Video.RenderScale": "70",
                "GraphicsVR.CloudsQuality": "0",
                "GraphicsVR.ShadowQuality": "0",
                "GraphicsVR.Reflections": "0",
                "GraphicsVR.SSContact": "0",
                "GraphicsVR.MotionBlur": "0",
                "RayTracing.Enabled": "0",
            },
        },
        config_path=config_path,
        usercfg_path=usercfg_path,
        auto_restart=auto_restart,
        force_restart=force_restart,
        dry_run=dry_run,
    )


@mcp.tool()
def restart_pimax() -> dict:
    """Restart Pimax Play so that changed settings take effect.

    Use this after changing Pimax settings while a VR game was running.
    Typical workflow:
      1. User changes Pimax settings (brightness etc.) while in VR
      2. Settings are saved but Pimax is not restarted (VR session active)
      3. User exits VR game
      4. User says "Pimax neu starten" → this tool restarts Pimax Play

    Will NOT restart if a VR game (MSFS) is still running.
    """
    if _is_msfs_running():
        return {
            "error": "MSFS läuft noch! Bitte zuerst das VR-Spiel beenden, "
                     "dann 'Pimax neu starten' sagen.",
        }

    pimax_is_running = any(_is_running(name) for name in _PIMAX_EXE_NAMES)
    if not pimax_is_running:
        # Just start it
        pimax_exe = _find_pimax()
        if pimax_exe:
            subprocess.Popen(
                [pimax_exe],
                creationflags=getattr(subprocess, "DETACHED_PROCESS", 0),
            )
            return {
                "status": "ok",
                "message": "Pimax Play gestartet. Gespeicherte Einstellungen werden geladen.",
            }
        return {"error": "Pimax Play konnte nicht gefunden werden."}

    # Kill and restart
    for exe in _PIMAX_EXE_NAMES:
        if _is_running(exe):
            try:
                subprocess.run(
                    ["taskkill", "/IM", exe, "/F"],
                    capture_output=True, timeout=10,
                )
            except Exception:
                pass

    import time
    time.sleep(3)

    pimax_exe = _find_pimax()
    if pimax_exe:
        subprocess.Popen(
            [pimax_exe],
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0),
        )
        return {
            "status": "ok",
            "message": "Pimax Play neu gestartet — alle gespeicherten Einstellungen sind jetzt aktiv.",
        }
    return {"error": "Pimax Play beendet, aber konnte nicht neu gestartet werden."}


@mcp.tool()
def status_check(
    config_path: str = "",
    usercfg_path: str = "",
) -> dict:
    """One-shot health check of the whole Pimax/MSFS/GPU stack.

    Call this at the START of a session to find out what's available
    before suggesting actions. Reports:
    - Pimax config: found path, currently running?, current resolution + brightness
    - MSFS UserCfg.opt: found path, MSFS running?, current DLSS + RenderScale
    - GPU: detected name, VRAM tier
    - Active VR session: yes/no (vrserver.exe)
    - Last log entries (helps debug "ich sehe keine Änderung")

    Args:
        config_path: Pimax config path (auto-detected if empty).
        usercfg_path: MSFS UserCfg.opt path (auto-detected if empty).
    """
    out: dict = {}

    # ── Pimax ──
    pimax_block: dict = {}
    cfg = _find_pimax_config(config_path)
    if cfg is None:
        pimax_block["config"] = "not_found"
    else:
        pimax_block["config_file"] = str(cfg)
        try:
            current = _read_pimax_settings(cfg)
            pimax_block["renderResolution"] = current.get("renderResolution")
            pimax_block["refreshRate"] = current.get("refreshRate")
            pimax_block["ffrLevel"] = current.get("ffrLevel")
            pimax_block["brightness_left"] = current.get("piplay_color_brightness_0")
            pimax_block["brightness_right"] = current.get("piplay_color_brightness_1")
        except Exception as exc:
            pimax_block["read_error"] = str(exc)
    pimax_block["pimax_running"] = any(_is_running(n) for n in _PIMAX_EXE_NAMES)
    out["pimax"] = pimax_block

    # ── MSFS ──
    msfs_block: dict = {}
    msfs_cfg = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH
    msfs_block["config_file"] = str(msfs_cfg)
    msfs_block["config_exists"] = msfs_cfg.exists()
    if msfs_cfg.exists():
        try:
            settings = _read_current_settings(msfs_cfg)
            dlss_raw = settings.get("Video.DLSS", "?")
            dlss_def = SETTING_DEFS.get("Video.DLSS", {}).get("values", {})
            msfs_block["dlss"] = dlss_def.get(str(dlss_raw), str(dlss_raw))
            msfs_block["render_scale"] = settings.get("Video.RenderScale")
            msfs_block["clouds_vr"] = settings.get("GraphicsVR.CloudsQuality")
            msfs_block["shadows_vr"] = settings.get("GraphicsVR.ShadowQuality")
        except Exception as exc:
            msfs_block["read_error"] = str(exc)
    msfs_block["msfs_running"] = _is_msfs_running()
    out["msfs"] = msfs_block

    # ── VR session ──
    out["vr_session_active"] = _is_running("vrserver.exe")

    # ── GPU ──
    try:
        gpu_name, gpu_tier, vram_mb = _detect_gpu_tier()
        out["gpu"] = {"name": gpu_name, "tier": gpu_tier, "vram_mb": vram_mb}
    except Exception as exc:
        out["gpu"] = {"error": str(exc)}

    # ── Recent log lines ──
    try:
        if _LOG_FILE.exists():
            lines = _LOG_FILE.read_text(encoding="utf-8").splitlines()
            out["recent_log"] = lines[-10:]
        else:
            out["recent_log"] = []
    except Exception as exc:
        out["recent_log"] = [f"<log read error: {exc}>"]

    return out


@mcp.tool()
def revert_last_change(
    target: Literal["pimax", "msfs", "auto"] = "auto",
    config_path: str = "",
    usercfg_path: str = "",
) -> dict:
    """Undo the most recent settings change.

    Convenience wrapper. The user just says "mach das rückgängig" / "revert"
    and this tool figures out which file to restore based on which has the
    more recent backup.

    Args:
        target: "pimax", "msfs", or "auto" (use whichever has the newer backup).
        config_path: Pimax config path.
        usercfg_path: MSFS UserCfg.opt path.
    """
    pimax_cfg = _find_pimax_config(config_path)
    msfs_cfg = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH

    pimax_latest = pimax_cfg.parent / "pimax_backups" / "pimax_config_latest.json" if pimax_cfg else None
    msfs_latest = msfs_cfg.parent / "UserCfg_backups" / "UserCfg_latest.opt" if msfs_cfg.exists() else None

    pimax_mtime = pimax_latest.stat().st_mtime if pimax_latest and pimax_latest.exists() else 0
    msfs_mtime = msfs_latest.stat().st_mtime if msfs_latest and msfs_latest.exists() else 0

    if target == "auto":
        if pimax_mtime == 0 and msfs_mtime == 0:
            return {
                "error": "Keine Backups gefunden — entweder wurde noch nichts geändert "
                         "oder die Backup-Ordner wurden gelöscht.",
                "pimax_backup_dir": str(pimax_latest.parent) if pimax_latest else None,
                "msfs_backup_dir": str(msfs_latest.parent) if msfs_latest else None,
            }
        target = "pimax" if pimax_mtime > msfs_mtime else "msfs"
        logger.info("revert_last_change auto-selected target=%s", target)

    if target == "pimax":
        if not pimax_latest or not pimax_latest.exists():
            return {"error": "Kein Pimax-Backup gefunden."}
        if not pimax_cfg:
            return {"error": "Pimax-Config-Pfad nicht gefunden."}
        pimax_cfg.write_text(pimax_latest.read_text(encoding="utf-8"), encoding="utf-8")
        # The latest backup is the JSON snapshot — registry is NOT auto-restored
        # (we don't have a registry snapshot). Warn the user.
        logger.info("Pimax config restored from %s", pimax_latest)
        return {
            "status": "ok",
            "target": "pimax",
            "restored_from": str(pimax_latest),
            "config_file": str(pimax_cfg),
            "warnung": (
                "JSON wiederhergestellt. Werte in der Windows-Registry werden NICHT "
                "automatisch zurückgesetzt — bei Bedarf 'Pimax neu starten' nutzen."
            ),
        }

    if target == "msfs":
        if not msfs_latest or not msfs_latest.exists():
            return {"error": "Kein MSFS-Backup gefunden."}
        # Kill MSFS first if running
        if _is_msfs_running():
            _kill_msfs(timeout_s=30)
            _wait_for_file_unlocked(msfs_cfg, timeout_s=10)
        msfs_cfg.write_text(msfs_latest.read_text(encoding="utf-8"), encoding="utf-8")
        logger.info("MSFS UserCfg.opt restored from %s", msfs_latest)
        return {
            "status": "ok",
            "target": "msfs",
            "restored_from": str(msfs_latest),
            "config_file": str(msfs_cfg),
            "hinweis": "MSFS muss gestartet (oder neu gestartet) werden damit die Änderung sichtbar wird.",
        }

    return {"error": f"Unbekannter target: {target!r}"}


@mcp.tool()
def adjust_pimax_brightness(
    direction: Literal["up", "down", "set", "get"] = "up",
    amount: int = 10,
    target_value: int | None = None,
    config_path: str = "",
) -> dict:
    """Adjust Pimax headset brightness quickly.

    Shortcut tool for the most common request. Use this when the user says
    things like "Helligkeit erhöhen", "heller machen", "Brightness up/down".

    Changes both eyes simultaneously. Use diagnose_pimax first if
    values seem wrong — Pimax stores settings in multiple places.

    Args:
        direction: "up" to increase, "down" to decrease, "set" for absolute
                   value, "get" to read the current brightness without changing.
        amount: How much to change (default 10). Ignored if direction is "set" or "get".
        target_value: Absolute value when direction is "set". Must be 0–255.
        config_path: Path to Pimax config (auto-detected if empty).
    """
    if direction == "get":
        cfg = _find_pimax_config(config_path)
        if cfg is None:
            return _pimax_not_found_error()
        try:
            settings = _read_pimax_settings(cfg)
            brightness_l = settings.get("piplay_color_brightness_0",
                            settings.get("brightness", None))
            brightness_r = settings.get("piplay_color_brightness_1", brightness_l)
            return {
                "status": "ok",
                "brightness_left": brightness_l,
                "brightness_right": brightness_r,
                "config_file": str(cfg),
            }
        except Exception as e:
            return {"error": str(e)}

    if direction == "set":
        if target_value is None:
            return {"error": "target_value is required when direction='set'."}
        if not (0 <= target_value <= 255):
            return {
                "error": (
                    f"target_value must be between 0 and 255, got {target_value}. "
                    "Pimax brightness range is 0 (darkest) to 255 (brightest)."
                )
            }
        val_str = str(target_value)
    elif direction == "up":
        val_str = f"+{amount}"
    else:
        val_str = f"-{amount}"

    result = set_pimax_setting(
        setting="brightness",
        value=val_str,
        config_path=config_path,
    )
    # Enrich result with confirmed new value for easy verification
    if isinstance(result, dict) and result.get("status") == "ok":
        new_raw = result.get("new_value")
        if new_raw is None:
            new_raw = (result.get("changed_keys", {})
                       .get("piplay_color_brightness_0", {})
                       .get("new"))
        if new_raw is not None:
            result["confirmed_brightness"] = new_raw
    return result


# ---------------------------------------------------------------------------
# Pimax — extended headset info, comprehensive settings, MSFS optimizer
# ---------------------------------------------------------------------------

# ── Per-model capability database ─────────────────────────────────────────
# Keys are lowercase fragments of the model name string.
# Entries are ordered longest-first so Crystal Super matches before Crystal.

_PIMAX_MODEL_DB: dict[str, dict] = {
    "crystal super": {
        "display_name": "Pimax Crystal Super",
        "resolution_per_eye": (3840, 2160),
        "max_refresh_hz": [72, 90, 120],
        "eye_tracking": True,
        "foveated_rendering": True,
        "brightness_control": True,
        "pitool_compatible": False,
        "pimax_play_compatible": True,
    },
    "crystal light": {
        "display_name": "Pimax Crystal Light",
        "resolution_per_eye": (2880, 2880),
        "max_refresh_hz": [72, 90, 120],
        "eye_tracking": True,
        "foveated_rendering": True,
        "brightness_control": True,
        "pitool_compatible": False,
        "pimax_play_compatible": True,
    },
    "crystal": {
        "display_name": "Pimax Crystal",
        "resolution_per_eye": (2880, 2880),
        "max_refresh_hz": [72, 90, 120],
        "eye_tracking": True,
        "foveated_rendering": True,
        "brightness_control": True,
        "pitool_compatible": False,
        "pimax_play_compatible": True,
    },
    "8kx": {
        "display_name": "Pimax 8KX",
        "resolution_per_eye": (3840, 2160),
        "max_refresh_hz": [72, 90, 120],
        "eye_tracking": False,
        "foveated_rendering": False,
        "brightness_control": True,
        "pitool_compatible": True,
        "pimax_play_compatible": True,
    },
    "8k+": {
        "display_name": "Pimax 8K+",
        "resolution_per_eye": (3840, 2160),
        "max_refresh_hz": [72, 90, 110],
        "eye_tracking": False,
        "foveated_rendering": False,
        "brightness_control": True,
        "pitool_compatible": True,
        "pimax_play_compatible": True,
    },
    "5k super": {
        "display_name": "Pimax 5K Super",
        "resolution_per_eye": (2560, 1440),
        "max_refresh_hz": [72, 90, 120, 144, 180],
        "eye_tracking": False,
        "foveated_rendering": False,
        "brightness_control": True,
        "pitool_compatible": True,
        "pimax_play_compatible": False,
    },
    "5k+": {
        "display_name": "Pimax 5K+",
        "resolution_per_eye": (2560, 1440),
        "max_refresh_hz": [72, 90, 110],
        "eye_tracking": False,
        "foveated_rendering": False,
        "brightness_control": True,
        "pitool_compatible": True,
        "pimax_play_compatible": False,
    },
    "5k": {
        "display_name": "Pimax 5K",
        "resolution_per_eye": (2560, 1440),
        "max_refresh_hz": [72, 90],
        "eye_tracking": False,
        "foveated_rendering": False,
        "brightness_control": True,
        "pitool_compatible": True,
        "pimax_play_compatible": False,
    },
}


def _match_pimax_model(model_str: str) -> dict:
    """Match a raw model string to a known Pimax model capability entry.

    Tries longest model-name keys first to avoid 'crystal' matching before
    'crystal super' / 'crystal light'.
    """
    if not model_str:
        return {}
    ml = model_str.lower()
    for key in sorted(_PIMAX_MODEL_DB, key=len, reverse=True):
        if key in ml:
            return dict(_PIMAX_MODEL_DB[key])
    return {
        "display_name": model_str,
        "eye_tracking": False,
        "foveated_rendering": False,
        "brightness_control": True,
        "max_refresh_hz": [72, 90],
    }


def _find_pitool_config() -> Path | None:
    """Find PiTool legacy configuration JSON.

    Searches:
    - %APPDATA%\\Pimax\\PiTool\\pitool.json
    - %LOCALAPPDATA%\\Pimax\\PiTool\\pitool.json
    - Common Program Files install directories
    - Registry HKCU\\SOFTWARE\\Pimax\\PiTool → InstallPath
    """
    appdata = os.environ.get("APPDATA", "")
    localappdata = os.environ.get("LOCALAPPDATA", "")
    candidates = [
        Path(appdata) / "Pimax" / "PiTool" / "pitool.json",
        Path(localappdata) / "Pimax" / "PiTool" / "pitool.json",
        Path(appdata) / "Pimax" / "pitool.json",
        Path(localappdata) / "Pimax" / "pitool.json",
        Path(r"C:\Program Files\Pimax\PiTool\pitool.json"),
        Path(r"C:\Program Files (x86)\Pimax\PiTool\pitool.json"),
        Path(r"C:\Program Files\PiTool\pitool.json"),
        Path(r"C:\Program Files (x86)\PiTool\pitool.json"),
        Path(r"D:\Program Files\Pimax\PiTool\pitool.json"),
        Path(r"D:\PiTool\pitool.json"),
    ]
    for p in candidates:
        if p.exists():
            return p
    # Registry fallback
    try:
        r = _reg(["query", r"HKCU\SOFTWARE\Pimax\PiTool", "/v", "InstallPath"], timeout=5)
        if r.returncode == 0:
            for line in r.stdout.splitlines():
                if "InstallPath" in line:
                    parts = line.strip().split(None, 2)
                    if len(parts) >= 3:
                        install_path = Path(parts[-1].strip())
                        candidate = install_path / "pitool.json"
                        if candidate.exists():
                            return candidate
    except Exception:
        pass
    return None


def _find_pimax_play_config() -> Path | None:
    """Find Pimax Play / Pimax Client configuration JSON.

    Searches:
    - %APPDATA%\\Pimax\\runtime\\default.vrsettings
    - %LOCALAPPDATA%\\Pimax\\PimaxPlay\\settings.json
    - C:\\Program Files\\Pimax\\runtime\\
    """
    appdata = os.environ.get("APPDATA", "")
    localappdata = os.environ.get("LOCALAPPDATA", "")
    candidates = [
        Path(appdata) / "Pimax" / "runtime" / "default.vrsettings",
        Path(appdata) / "Pimax" / "runtime" / "settings.json",
        Path(localappdata) / "Pimax" / "PimaxPlay" / "settings.json",
        Path(localappdata) / "Pimax" / "PimaxPlay" / "config.json",
        Path(localappdata) / "Pimax" / "Pimax Play" / "settings.json",
        Path(localappdata) / "Pimax" / "PimaxClient" / "settings.json",
        Path(appdata) / "Pimax" / "PimaxPlay" / "settings.json",
        Path(r"C:\Program Files\Pimax\runtime\default.vrsettings"),
        Path(r"C:\Program Files\Pimax\runtime\settings.json"),
        Path(r"C:\Program Files\Pimax\PimaxPlay\settings.json"),
        Path(r"C:\Program Files\Pimax\PimaxPlay\config.json"),
        Path(r"C:\Program Files (x86)\Pimax\runtime\default.vrsettings"),
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def _find_pimax_openxr_config() -> Path | None:
    """Find the Pimax Play OpenXR layer JSON configuration file.

    When Pimax Play is the active OpenXR runtime, rendering overrides
    (supersampling, foveated rendering) live in this file.
    """
    appdata = os.environ.get("APPDATA", "")
    candidates = [
        Path(appdata) / "Pimax" / "runtime" / "openxr_default.json",
        Path(appdata) / "Pimax" / "runtime" / "openxr.json",
        Path(r"C:\Program Files\Pimax\runtime\openxr_default.json"),
        Path(r"C:\Program Files\Pimax\runtime\openxr.json"),
        Path(r"C:\Program Files (x86)\Pimax\runtime\openxr_default.json"),
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def _get_hw_for_tier() -> tuple[float, str, int, float]:
    """Return (vram_gb, gpu_name, cpu_cores, ram_gb) for VR tier classification.

    Used by optimize_pimax_for_msfs and get_pimax_recommended_settings.
    Falls back to conservative defaults if hardware queries fail.
    """
    vram_gb: float = 8.0
    gpu_name: str = ""
    cpu_cores: int = 8
    ram_gb: float = 16.0
    try:
        pynvml.nvmlInit()
        try:
            h = pynvml.nvmlDeviceGetHandleByIndex(0)
            name = pynvml.nvmlDeviceGetName(h)
            if isinstance(name, bytes):
                name = name.decode()
            mem = pynvml.nvmlDeviceGetMemoryInfo(h)
            gpu_name = name
            vram_gb = round(mem.total / 1024 ** 3, 1)
        finally:
            pynvml.nvmlShutdown()
    except Exception:
        pass
    try:
        cpu_raw = _ps_json(
            "Get-CimInstance Win32_Processor | Select-Object NumberOfCores"
        )
        cpu_cores = int(cpu_raw[0].get("NumberOfCores", 8)) if cpu_raw else 8
    except Exception:
        pass
    try:
        ram_raw = _ps_json(
            "Get-CimInstance Win32_PhysicalMemory | Select-Object Capacity"
        )
        total_bytes = sum(int(r.get("Capacity", 0) or 0) for r in ram_raw)
        ram_gb = round(total_bytes / 1024 ** 3, 1) if total_bytes > 0 else 16.0
    except Exception:
        pass
    return vram_gb, gpu_name, cpu_cores, ram_gb


@mcp.tool()
def get_pimax_headset_info() -> dict:
    """Detect the connected Pimax headset model, firmware, refresh rate, and feature set.

    Returns:
    - model name and per-eye resolution
    - maximum supported refresh rates for this model
    - feature flags: eye_tracking, foveated_rendering, brightness_control
    - which software stack is active: pitool vs pimax_play
    - paths to detected config files (PiTool, Pimax Play, OpenXR layer)
    - current refresh rate (if readable from config/registry)
    - serial number and firmware version (if available via registry)

    Call this first before optimizing Pimax settings so the AI knows which
    features are available and which config files to target.
    """
    try:
        result: dict = {
            "model": "unknown",
            "software_stack": "unknown",
            "pitool_installed": False,
            "pimax_play_installed": False,
            "pimax_running": False,
            "pitool_running": False,
            "pitool_config": None,
            "pimax_play_config": None,
            "openxr_config": None,
            "features": {
                "eye_tracking": False,
                "foveated_rendering": False,
                "brightness_control": True,
            },
            "max_refresh_hz": [72, 90],
            "resolution_per_eye": None,
        }

        # Detect running processes
        pimax_running = any(_is_running(name) for name in _PIMAX_EXE_NAMES)
        pitool_running = _is_running("PiTool.exe") or _is_running("PiServer.exe")
        result["pimax_running"] = pimax_running
        result["pitool_running"] = pitool_running

        # Config file detection
        pitool_cfg = _find_pitool_config()
        piplay_cfg = _find_pimax_play_config()
        openxr_cfg = _find_pimax_openxr_config()
        result["pitool_installed"] = pitool_cfg is not None
        result["pimax_play_installed"] = piplay_cfg is not None
        result["pitool_config"] = str(pitool_cfg) if pitool_cfg else None
        result["pimax_play_config"] = str(piplay_cfg) if piplay_cfg else None
        result["openxr_config"] = str(openxr_cfg) if openxr_cfg else None

        # Determine active software stack
        if piplay_cfg and (pimax_running or not pitool_running):
            result["software_stack"] = "pimax_play"
        elif pitool_cfg or pitool_running:
            result["software_stack"] = "pitool"
        elif pimax_running:
            result["software_stack"] = "pimax_play"

        # Model / serial / firmware detection via registry
        model_str = ""
        serial = ""
        firmware = ""
        current_refresh: int | None = None
        for reg_path in [
            r"HKCU\SOFTWARE\Pimax",
            r"HKCU\SOFTWARE\Pimax\PimaxPlay",
            r"HKLM\SOFTWARE\Pimax",
            r"HKLM\SOFTWARE\Pimax\PimaxPlay",
        ]:
            try:
                r = _reg(["query", reg_path], timeout=5)
                if r.returncode != 0:
                    continue
                for line in r.stdout.splitlines():
                    ll = line.strip().lower()
                    parts = line.strip().split(None, 2)
                    val = parts[-1].strip() if len(parts) >= 3 else ""
                    if not model_str and any(
                        kw in ll for kw in ("devicename", "headsetmodel", "productname")
                    ):
                        model_str = val
                    if not serial and "serial" in ll:
                        serial = val
                    if not firmware and "firmware" in ll:
                        firmware = val
                    if current_refresh is None and "refreshrate" in ll:
                        try:
                            current_refresh = int(val)
                        except ValueError:
                            pass
            except Exception:
                pass

        # WMI USB device fallback
        if not model_str:
            try:
                usb_raw = _ps_json(
                    "Get-PnpDevice | Where-Object { $_.FriendlyName -match 'Pimax' } | "
                    "Select-Object FriendlyName, InstanceId, Status"
                )
                for dev in usb_raw:
                    fname = dev.get("FriendlyName", "") or ""
                    if fname:
                        model_str = fname
                        break
            except Exception:
                pass

        # Config-file fallback for model name
        if not model_str:
            for cfg_path in [piplay_cfg, pitool_cfg]:
                if cfg_path and cfg_path.exists():
                    try:
                        data = _json.loads(cfg_path.read_text(encoding="utf-8"))
                        for key in (
                            "deviceName", "headsetModel", "model", "productName",
                            "DeviceName", "HeadsetModel", "Model",
                        ):
                            v = data.get(key, "")
                            if v and isinstance(v, str):
                                model_str = v
                                break
                    except Exception:
                        pass
                if model_str:
                    break

        # Current refresh rate from config files
        if current_refresh is None:
            for cfg_path in [piplay_cfg, pitool_cfg]:
                if cfg_path and cfg_path.exists():
                    try:
                        data = _json.loads(cfg_path.read_text(encoding="utf-8"))
                        for key in (
                            "refreshRate", "refresh_rate", "piplay_refreshrate",
                            "RefreshRate", "displayHz",
                        ):
                            v = data.get(key)
                            if v is not None:
                                try:
                                    current_refresh = int(v)
                                    break
                                except (ValueError, TypeError):
                                    pass
                    except Exception:
                        pass
                if current_refresh is not None:
                    break

        if serial:
            result["serial_number"] = serial
        if firmware:
            result["firmware_version"] = firmware
        if current_refresh:
            result["current_refresh_hz"] = current_refresh

        # Map to capability record
        model_info = _match_pimax_model(model_str)
        if model_info:
            result["model"] = model_info.get("display_name", model_str or "unknown")
            result["features"] = {
                "eye_tracking": model_info.get("eye_tracking", False),
                "foveated_rendering": model_info.get("foveated_rendering", False),
                "brightness_control": model_info.get("brightness_control", True),
            }
            result["max_refresh_hz"] = model_info.get("max_refresh_hz", [72, 90])
            if model_info.get("resolution_per_eye"):
                w, h = model_info["resolution_per_eye"]
                result["resolution_per_eye"] = f"{w}x{h}"
        elif model_str:
            result["model"] = model_str

        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_pimax_settings(config_path: str = "") -> dict:
    """Read ALL current Pimax settings with values and valid ranges.

    Detects whether PiTool or Pimax Play is installed and reads from the
    correct location. Merges PiTool config, Pimax Play config, and registry
    into a unified view.

    Returns:
    - current_values: all detected setting key/value pairs
    - valid_ranges: type, min/max or options, and description per key
    - sources: list of config files/registry entries that were read
    - software_stack: pitool / pimax_play / unknown
    - openxr_settings: contents of the Pimax OpenXR layer config (if found)

    Use this before set_pimax_settings or optimize_pimax_for_msfs to
    understand what is currently configured and what can be changed.
    """
    try:
        result: dict = {
            "current_values": {},
            "valid_ranges": {},
            "sources": [],
            "software_stack": "unknown",
        }

        _VALID_RANGES: dict[str, dict] = {
            "renderResolution": {
                "type": "float", "min": 0.5, "max": 2.0,
                "description": "Supersampling multiplier (1.0 = native resolution)",
            },
            "refreshRate": {
                "type": "int", "options": [72, 90, 110, 120, 144],
                "description": "Display refresh rate in Hz (model-dependent)",
            },
            "fov": {
                "type": "str", "options": ["small", "normal", "large", "potato"],
                "description": "Field of view preset",
            },
            "fovLevel": {
                "type": "int", "options": [0, 1, 2, 3],
                "description": "FOV level: 0=Potato, 1=Small, 2=Normal, 3=Large",
            },
            "smartSmoothing": {
                "type": "bool",
                "description": "Asynchronous Spacewarp / motion smoothing reprojection",
            },
            "compulsorySmoothing": {
                "type": "bool",
                "description": "Always-on reprojection (ignores FPS threshold)",
            },
            "ffrLevel": {
                "type": "int", "options": [0, 1, 2, 3, 4],
                "description": "Fixed Foveated Rendering level: 0=Off, 1=Low … 4=Ultra",
            },
            "parallelProjection": {
                "type": "bool",
                "description": "Required True for correct MSFS VR rendering (prevents warping)",
            },
            "brightness": {
                "type": "int", "min": 0, "max": 255,
                "description": "Backlight brightness applied to both eyes",
            },
            "contrast": {
                "type": "int", "min": 0, "max": 255,
                "description": "Display contrast level",
            },
            "eyeTracking": {
                "type": "bool",
                "description": "Eye tracking (Pimax Crystal / Crystal Super only)",
            },
            "dynamicFoveatedRendering": {
                "type": "bool",
                "description": "Dynamic FFR driven by eye tracking gaze data",
            },
            "ipd": {
                "type": "float", "min": 55.0, "max": 75.0,
                "description": "Interpupillary distance in mm",
            },
            "hiddenAreaMask": {
                "type": "bool",
                "description": "Hidden area mask to skip rendering outside visible FOV",
            },
            "piplay_color_brightness_0": {
                "type": "int", "min": 0, "max": 100,
                "description": "Left eye brightness slider (Pimax Play Device Settings)",
            },
            "piplay_color_brightness_1": {
                "type": "int", "min": 0, "max": 100,
                "description": "Right eye brightness slider (Pimax Play Device Settings)",
            },
            "piplay_color_contrast_0": {
                "type": "int", "min": 0, "max": 100,
                "description": "Left eye contrast slider",
            },
            "piplay_color_contrast_1": {
                "type": "int", "min": 0, "max": 100,
                "description": "Right eye contrast slider",
            },
            "piplay_color_saturation_0": {
                "type": "int", "min": 0, "max": 100,
                "description": "Left eye saturation slider",
            },
            "piplay_color_saturation_1": {
                "type": "int", "min": 0, "max": 100,
                "description": "Right eye saturation slider",
            },
            "piplay_refreshrate": {
                "type": "int", "options": [72, 90, 110, 120, 144],
                "description": "Refresh rate as stored by Pimax Play",
            },
        }
        result["valid_ranges"] = _VALID_RANGES

        # PiTool config
        pitool_cfg = _find_pitool_config()
        if pitool_cfg:
            try:
                data = _json.loads(pitool_cfg.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    result["current_values"].update(data)
                    result["sources"].append({"file": str(pitool_cfg), "type": "pitool"})
                    result["software_stack"] = "pitool"
            except Exception as exc:
                result["sources"].append(
                    {"file": str(pitool_cfg), "type": "pitool", "error": str(exc)}
                )

        # Main Pimax config (via existing _find_pimax_config)
        cfg = _find_pimax_config(config_path)
        if cfg and cfg != pitool_cfg:
            try:
                settings = _read_pimax_settings(cfg)
                result["current_values"].update(settings)
                result["sources"].append({"file": str(cfg), "type": "main_config"})
                if result["software_stack"] == "unknown":
                    result["software_stack"] = "pimax_play"
            except Exception as exc:
                result["sources"].append(
                    {"file": str(cfg), "type": "main_config", "error": str(exc)}
                )

        # Pimax Play config
        piplay_cfg = _find_pimax_play_config()
        if piplay_cfg and piplay_cfg != cfg:
            try:
                data = _json.loads(piplay_cfg.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    result["current_values"].update(data)
                    result["sources"].append({"file": str(piplay_cfg), "type": "pimax_play"})
                    if result["software_stack"] == "unknown":
                        result["software_stack"] = "pimax_play"
            except Exception as exc:
                result["sources"].append(
                    {"file": str(piplay_cfg), "type": "pimax_play", "error": str(exc)}
                )

        # Registry overlay (highest priority — Pimax Play reads live from registry)
        try:
            reg_vals = _read_pimax_registry()
            if reg_vals:
                result["current_values"].update(reg_vals)
                result["sources"].append({"file": "registry", "type": "registry"})
        except Exception:
            pass

        # OpenXR layer config
        openxr_cfg = _find_pimax_openxr_config()
        if openxr_cfg:
            try:
                data = _json.loads(openxr_cfg.read_text(encoding="utf-8"))
                result["openxr_settings"] = data
                result["sources"].append({"file": str(openxr_cfg), "type": "openxr"})
            except Exception as exc:
                result["sources"].append(
                    {"file": str(openxr_cfg), "type": "openxr", "error": str(exc)}
                )

        if not result["current_values"]:
            return _pimax_not_found_error()

        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_pimax_settings(
    render_quality: float | None = None,
    smart_smoothing: bool | None = None,
    hidden_area_mask: bool | None = None,
    fov_mode: str | None = None,
    refresh_rate: int | None = None,
    brightness: int | None = None,
    contrast: int | None = None,
    foveated_rendering: bool | None = None,
    eye_tracking: bool | None = None,
    parallel_projection: bool | None = None,
    config_path: str = "",
    restart_service: bool = True,
    force_restart: bool = False,
    dry_run: bool = False,
) -> dict:
    """Set multiple Pimax settings at once in a single operation.

    Preferred over calling set_pimax_setting repeatedly — this applies all
    changes in one backup + write + registry update + optional restart cycle.

    Parameters (all optional — omit to leave unchanged):
        render_quality: Supersampling multiplier 0.5–2.0 (maps to renderResolution).
        smart_smoothing: Enable/disable ASW motion reprojection.
        hidden_area_mask: Enable/disable hidden area mask.
        fov_mode: "small" / "normal" / "large" / "potato".
        refresh_rate: 72 / 90 / 110 / 120 / 144 Hz (model-dependent).
        brightness: 0–255 backlight brightness (applied to both eyes).
        contrast: 0–255 contrast level (applied to both eyes).
        foveated_rendering: Enable/disable Fixed Foveated Rendering (Crystal only).
        eye_tracking: Enable/disable eye tracking + DFR (Crystal / Crystal Super only).
        parallel_projection: Set True — required for correct MSFS VR rendering.
        config_path: Override auto-detected Pimax config path.
        restart_service: Restart Pimax Play after applying (default True).
        force_restart: Restart even if a VR session is active.
        dry_run: Preview what would change without writing anything.
    """
    try:
        # Input validation
        if render_quality is not None and not (0.5 <= render_quality <= 2.0):
            return {
                "error": f"render_quality must be between 0.5 and 2.0, got {render_quality}"
            }
        if brightness is not None and not (0 <= brightness <= 255):
            return {
                "error": f"brightness must be between 0 and 255, got {brightness}. "
                         "Pimax brightness range is 0 (darkest) to 255 (brightest)."
            }
        if contrast is not None and not (0 <= contrast <= 255):
            return {"error": f"contrast must be between 0 and 255, got {contrast}"}
        if refresh_rate is not None and refresh_rate not in (72, 90, 110, 120, 144, 160, 180):
            return {
                "error": f"refresh_rate {refresh_rate} Hz is not a standard Pimax value. "
                         "Valid: 72, 90, 110, 120, 144"
            }
        if fov_mode is not None and fov_mode.lower() not in ("small", "normal", "large", "potato"):
            return {
                "error": f"fov_mode must be 'small', 'normal', 'large', or 'potato', "
                         f"got '{fov_mode}'"
            }

        cfg = _find_pimax_config(config_path)
        if cfg is None:
            return _pimax_not_found_error()

        current = _read_pimax_settings(cfg)
        if not current:
            return {"error": f"Could not read Pimax settings from {cfg}"}

        # Build updates dict — only from non-None arguments
        updates: dict[str, object] = {}
        if render_quality is not None:
            updates["renderResolution"] = render_quality
        if smart_smoothing is not None:
            updates["smartSmoothing"] = smart_smoothing
        if hidden_area_mask is not None:
            updates["hiddenAreaMask"] = hidden_area_mask
        if fov_mode is not None:
            fov_lower = fov_mode.lower()
            updates["fov"] = fov_lower
            fov_level_map = {"potato": 0, "small": 1, "normal": 2, "large": 3}
            updates["fovLevel"] = fov_level_map.get(fov_lower, 2)
        if refresh_rate is not None:
            updates["refreshRate"] = refresh_rate
            updates["piplay_refreshrate"] = refresh_rate
        if brightness is not None:
            updates["brightness"] = brightness
            updates["piplay_color_brightness_0"] = brightness
            updates["piplay_color_brightness_1"] = brightness
        if contrast is not None:
            updates["contrast"] = contrast
            updates["piplay_color_contrast_0"] = contrast
            updates["piplay_color_contrast_1"] = contrast
        if foveated_rendering is not None:
            updates["ffrLevel"] = 2 if foveated_rendering else 0
        if eye_tracking is not None:
            updates["eyeTracking"] = eye_tracking
            updates["dynamicFoveatedRendering"] = eye_tracking
        if parallel_projection is not None:
            updates["parallelProjection"] = parallel_projection

        if not updates:
            return {"status": "no_changes", "note": "No settings were specified."}

        # Build preview of changes
        would_change = {
            k: {
                "old": current.get(k, "<not set>"),
                "new": v,
                "label": PIMAX_SETTING_DEFS.get(k, {}).get("label", k),
            }
            for k, v in updates.items()
        }

        if dry_run:
            return {
                "status": "dry_run",
                "would_change": would_change,
                "config_file": str(cfg),
                "note": "Set dry_run=False to apply.",
            }

        # Apply to in-memory state
        for k, v in updates.items():
            current[k] = v

        backup_file = _pimax_create_backup(cfg, current)

        # Write JSON config (only update keys that already exist in the file)
        try:
            json_data = _json.loads(cfg.read_text(encoding="utf-8"))
            if isinstance(json_data, dict):
                for k, v in updates.items():
                    if k in json_data:
                        json_data[k] = v
                cfg.write_text(
                    _json.dumps(json_data, indent=2, ensure_ascii=False), encoding="utf-8"
                )
        except Exception as exc:
            logger.warning("set_pimax_settings: could not write JSON config: %s", exc)

        registry_audit = _apply_pimax_settings_to_registry(updates, verify=True)

        result: dict = {
            "status": "ok",
            "applied": {
                k: {
                    "label": PIMAX_SETTING_DEFS.get(k, {}).get("label", k),
                    "old": would_change[k]["old"],
                    "new": v,
                }
                for k, v in updates.items()
            },
            "config_file": str(cfg),
            "backup": str(backup_file),
            "registry_audit": registry_audit,
        }
        _pimax_apply_restart(result, restart_service, force_restart)
        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_pimax_recommended_settings(hardware_tier: str = "") -> dict:
    """Return recommended Pimax settings for MSFS VR based on hardware — dry run only.

    Does NOT apply any changes. Returns a full recommendation dict showing
    exactly what optimize_pimax_for_msfs would set and why, so the user
    can review and decide before committing.

    The hardware tier is auto-detected from GPU/VRAM if not supplied.
    Override with hardware_tier: "ultra", "high", "mid_high", "mid", or "low".

    Returns:
    - detected_tier / tier_label: which GPU performance tier was used
    - gpu_name / vram_gb / cpu_cores / ram_gb: detected hardware
    - recommended: dict of setting → {value, reason}
    - summary: one-line human-readable description
    """
    try:
        vram_gb, gpu_name, cpu_cores, ram_gb = _get_hw_for_tier()
        tier = (
            hardware_tier.strip().lower()
            if hardware_tier.strip()
            else _classify_vr_tier(vram_gb, gpu_name, cpu_cores, ram_gb)
        )

        _TIER_RECOMMENDATIONS: dict[str, dict[str, tuple]] = {
            "ultra": {
                "renderResolution": (1.3, "RTX 4090/5090 class GPU can sustain 1.3× SS in MSFS VR"),
                "smartSmoothing": (False, "Ultra tier maintains 45+ FPS natively; smoothing not needed"),
                "fov": ("large", "Sufficient GPU headroom for maximum Large FOV"),
                "fovLevel": (3, "Large FOV for maximum immersion"),
                "ffrLevel": (0, "No FFR needed at this GPU tier"),
                "parallelProjection": (True, "Required for correct MSFS VR rendering"),
                "refreshRate": (90, "90 Hz stable on ultra tier GPUs"),
            },
            "high": {
                "renderResolution": (1.1, "RTX 4080/4070Ti can handle 1.1× SS boost comfortably"),
                "smartSmoothing": (False, "High tier GPU sustains 45+ FPS without smoothing"),
                "fov": ("large", "High-end GPU handles Large FOV without major FPS loss"),
                "fovLevel": (3, "Large FOV"),
                "ffrLevel": (1, "Light FFR as insurance for complex scenery"),
                "parallelProjection": (True, "Required for MSFS VR"),
                "refreshRate": (90, "90 Hz with good frame pacing on high tier"),
            },
            "mid_high": {
                "renderResolution": (1.0, "RTX 3080/4070 at native resolution for stable framerate"),
                "smartSmoothing": (True, "Smart Smoothing maintains 45 FPS floor"),
                "fov": ("normal", "Normal FOV balances quality and performance"),
                "fovLevel": (2, "Normal FOV"),
                "ffrLevel": (2, "Medium FFR saves GPU load in peripheral vision"),
                "parallelProjection": (True, "Required for MSFS VR"),
                "refreshRate": (90, "90 Hz with Smart Smoothing halfrate at 45 FPS"),
            },
            "mid": {
                "renderResolution": (0.8, "RTX 3070/6700XT needs sub-native SS in MSFS VR"),
                "smartSmoothing": (True, "Smart Smoothing essential on mid-tier GPU"),
                "fov": ("normal", "Normal FOV keeps pixel count manageable"),
                "fovLevel": (2, "Normal FOV"),
                "ffrLevel": (3, "High FFR to maintain playable FPS on mid-tier"),
                "parallelProjection": (True, "Required for MSFS VR"),
                "refreshRate": (72, "72 Hz more achievable than 90 Hz on mid-tier"),
            },
            "low": {
                "renderResolution": (0.6, "Low-tier GPU requires significant SS reduction for MSFS"),
                "smartSmoothing": (True, "Smart Smoothing essential on low-tier hardware"),
                "fov": ("small", "Small FOV reduces pixel count significantly"),
                "fovLevel": (1, "Small FOV"),
                "ffrLevel": (4, "Maximum FFR to save GPU on low-tier hardware"),
                "parallelProjection": (True, "Required for MSFS VR"),
                "refreshRate": (72, "72 Hz is the only realistic target on low-tier hardware"),
            },
        }

        rec_raw = _TIER_RECOMMENDATIONS.get(tier, _TIER_RECOMMENDATIONS["mid"])
        recommended = {
            k: {"value": v, "reason": reason}
            for k, (v, reason) in rec_raw.items()
        }

        tier_labels = {
            "ultra": "Ultra (RTX 4090 / RTX 5090 class)",
            "high": "High (RTX 4080 / 4070 Ti class)",
            "mid_high": "Mid-High (RTX 3080 / 4070 class)",
            "mid": "Mid (RTX 3070 / 6700 XT class)",
            "low": "Low (RTX 3060 and below)",
        }
        rr = rec_raw["renderResolution"][0]
        ss = rec_raw["smartSmoothing"][0]
        fov = rec_raw["fov"][0]
        hz = rec_raw["refreshRate"][0]

        return {
            "status": "recommendation_only",
            "note": "Dry run — call optimize_pimax_for_msfs to apply these settings.",
            "detected_tier": tier,
            "tier_label": tier_labels.get(tier, tier),
            "gpu_name": gpu_name or "unknown",
            "vram_gb": vram_gb,
            "ram_gb": ram_gb,
            "cpu_cores": cpu_cores,
            "recommended": recommended,
            "summary": (
                f"For {tier_labels.get(tier, tier)} in MSFS VR: "
                f"render_quality={rr}, smart_smoothing={'on' if ss else 'off'}, "
                f"fov={fov}, {hz}Hz"
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def optimize_pimax_for_msfs(
    hardware_tier: str = "",
    config_path: str = "",
    restart_service: bool = True,
    force_restart: bool = False,
    dry_run: bool = False,
) -> dict:
    """Apply hardware-aware Pimax settings optimized specifically for MSFS VR.

    Auto-detects GPU tier and writes the appropriate Pimax settings to get
    the best performance/quality balance in Microsoft Flight Simulator VR.

    GPU tier → what gets applied:
    - ultra  (4090/5090):    render=1.3, smoothing=off, fov=large,  90Hz, FFR=off
    - high   (4080/4070Ti):  render=1.1, smoothing=off, fov=large,  90Hz, FFR=low
    - mid_high (3080/4070):  render=1.0, smoothing=on,  fov=normal, 90Hz, FFR=med
    - mid    (3070/6700XT):  render=0.8, smoothing=on,  fov=normal, 72Hz, FFR=high
    - low:                   render=0.6, smoothing=on,  fov=small,  72Hz, FFR=ultra

    All tiers force parallelProjection=True (required for MSFS VR).

    Args:
        hardware_tier: Override detected tier: "ultra"/"high"/"mid_high"/"mid"/"low".
        config_path: Pimax config path (auto-detected if empty).
        restart_service: Restart Pimax Play after applying (default True).
        force_restart: Restart even if a VR session is currently active.
        dry_run: Preview what would be set without writing anything.
    """
    try:
        vram_gb, gpu_name, cpu_cores, ram_gb = _get_hw_for_tier()
        tier = (
            hardware_tier.strip().lower()
            if hardware_tier.strip()
            else _classify_vr_tier(vram_gb, gpu_name, cpu_cores, ram_gb)
        )

        _TIER_SETTINGS: dict[str, dict] = {
            "ultra": {
                "renderResolution": 1.3,
                "smartSmoothing": False,
                "compulsorySmoothing": False,
                "fov": "large",
                "fovLevel": 3,
                "ffrLevel": 0,
                "parallelProjection": True,
                "refreshRate": 90,
            },
            "high": {
                "renderResolution": 1.1,
                "smartSmoothing": False,
                "compulsorySmoothing": False,
                "fov": "large",
                "fovLevel": 3,
                "ffrLevel": 1,
                "parallelProjection": True,
                "refreshRate": 90,
            },
            "mid_high": {
                "renderResolution": 1.0,
                "smartSmoothing": True,
                "compulsorySmoothing": False,
                "fov": "normal",
                "fovLevel": 2,
                "ffrLevel": 2,
                "parallelProjection": True,
                "refreshRate": 90,
            },
            "mid": {
                "renderResolution": 0.8,
                "smartSmoothing": True,
                "compulsorySmoothing": True,
                "fov": "normal",
                "fovLevel": 2,
                "ffrLevel": 3,
                "parallelProjection": True,
                "refreshRate": 72,
            },
            "low": {
                "renderResolution": 0.6,
                "smartSmoothing": True,
                "compulsorySmoothing": True,
                "fov": "small",
                "fovLevel": 1,
                "ffrLevel": 4,
                "parallelProjection": True,
                "refreshRate": 72,
            },
        }

        _REASONS: dict[str, dict[str, str]] = {
            "ultra": {
                "renderResolution": "4090/5090 class GPU can sustain 1.3× SS in MSFS VR",
                "smartSmoothing": "Ultra tier GPU runs 45+ FPS natively; smoothing not needed",
                "fov": "Enough GPU headroom for maximum Large FOV",
                "ffrLevel": "No FFR needed at this GPU tier",
                "refreshRate": "90 Hz stable on ultra tier",
            },
            "high": {
                "renderResolution": "4080/4070Ti handles 1.1× SS comfortably",
                "smartSmoothing": "High tier sustains 45+ FPS without smoothing",
                "fov": "High-end GPU handles Large FOV",
                "ffrLevel": "Light FFR as insurance for complex scenery",
                "refreshRate": "90 Hz with good frame pacing on high tier",
            },
            "mid_high": {
                "renderResolution": "3080/4070 at native resolution for stable framerate",
                "smartSmoothing": "Smart Smoothing maintains 45 FPS floor",
                "fov": "Normal FOV balances quality and performance",
                "ffrLevel": "Medium FFR saves GPU load in peripheral vision",
                "refreshRate": "90 Hz with Smart Smoothing halfrate at 45 FPS",
            },
            "mid": {
                "renderResolution": "3070/6700XT needs sub-native SS in MSFS VR",
                "smartSmoothing": "Smart Smoothing essential on mid-tier GPU",
                "fov": "Normal FOV keeps pixel count manageable",
                "ffrLevel": "High FFR required to maintain playable FPS",
                "refreshRate": "72 Hz more achievable than 90 Hz on mid-tier",
            },
            "low": {
                "renderResolution": "Low-tier GPU requires significant SS reduction for MSFS",
                "smartSmoothing": "Smart Smoothing is essential on low-tier hardware",
                "fov": "Small FOV reduces pixel count significantly",
                "ffrLevel": "Maximum FFR to save GPU on low-tier hardware",
                "refreshRate": "72 Hz is the only realistic target on low-tier hardware",
            },
        }

        updates = dict(_TIER_SETTINGS.get(tier, _TIER_SETTINGS["mid"]))
        reasons = _REASONS.get(tier, _REASONS["mid"])

        if dry_run:
            return {
                "status": "dry_run",
                "detected_tier": tier,
                "gpu_name": gpu_name or "unknown",
                "vram_gb": vram_gb,
                "would_set": {
                    k: {"value": v, "reason": reasons.get(k, "")}
                    for k, v in updates.items()
                },
                "note": "Set dry_run=False to apply these settings.",
            }

        cfg = _find_pimax_config(config_path)
        if cfg is None:
            return _pimax_not_found_error()

        current = _read_pimax_settings(cfg)
        if not current:
            return {"error": f"Could not read Pimax settings from {cfg}"}

        changed = {
            k: {"old": current.get(k, "<not set>"), "new": v, "reason": reasons.get(k, "")}
            for k, v in updates.items()
        }

        for k, v in updates.items():
            current[k] = v

        backup_file = _pimax_create_backup(cfg, current)

        try:
            json_data = _json.loads(cfg.read_text(encoding="utf-8"))
            if isinstance(json_data, dict):
                for k, v in updates.items():
                    if k in json_data:
                        json_data[k] = v
                cfg.write_text(
                    _json.dumps(json_data, indent=2, ensure_ascii=False), encoding="utf-8"
                )
        except Exception as exc:
            logger.warning("optimize_pimax_for_msfs: could not write JSON: %s", exc)

        registry_audit = _apply_pimax_settings_to_registry(updates, verify=True)

        result: dict = {
            "status": "ok",
            "detected_tier": tier,
            "gpu_name": gpu_name or "unknown",
            "vram_gb": vram_gb,
            "applied": changed,
            "config_file": str(cfg),
            "backup": str(backup_file),
            "registry_audit": registry_audit,
            "summary": (
                f"Pimax optimized for MSFS VR ({tier} tier): "
                f"render={updates['renderResolution']}, "
                f"smoothing={'on' if updates['smartSmoothing'] else 'off'}, "
                f"fov={updates['fov']}, {updates['refreshRate']}Hz"
            ),
        }
        _pimax_apply_restart(result, restart_service, force_restart)
        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_pimax_openxr_settings() -> dict:
    """Read the Pimax Play OpenXR layer configuration file.

    When Pimax Play is the active OpenXR runtime, some rendering settings
    (supersampling override, foveated rendering mode) are controlled via
    the OpenXR layer JSON file rather than the main Pimax config.

    Returns the raw settings dict plus the config file path.
    Use set_pimax_openxr_settings to write changes to this file.
    """
    try:
        cfg = _find_pimax_openxr_config()
        if cfg is None:
            return {
                "error": "Pimax OpenXR config not found.",
                "searched": [
                    str(
                        Path(os.environ.get("APPDATA", ""))
                        / "Pimax" / "runtime" / "openxr_default.json"
                    ),
                    r"C:\Program Files\Pimax\runtime\openxr_default.json",
                ],
                "hint": (
                    "This file only exists if Pimax Play is installed "
                    "and has been run at least once."
                ),
            }
        data = _json.loads(cfg.read_text(encoding="utf-8"))
        return {
            "status": "ok",
            "config_file": str(cfg),
            "settings": data,
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_pimax_openxr_settings(
    settings: dict,
    config_path: str = "",
    dry_run: bool = False,
) -> dict:
    """Write settings to the Pimax Play OpenXR layer JSON configuration.

    When Pimax Play is the active OpenXR runtime, this file controls
    OpenXR-level rendering overrides such as supersampling and foveated
    rendering mode. Changes are merged into the existing config.

    Args:
        settings: Dict of key→value pairs to write into the OpenXR config.
                  Example: {"supersampling": 1.2, "foveatedRendering": True}
        config_path: Override the config file path (auto-detected if empty).
        dry_run: Show what would change without writing anything.
    """
    try:
        if config_path:
            cfg = Path(config_path)
        else:
            cfg = _find_pimax_openxr_config()

        if cfg is None or not cfg.exists():
            return {
                "error": "Pimax OpenXR config not found. Cannot write settings.",
                "hint": (
                    "Ensure Pimax Play is installed and has been launched at least once "
                    "so the runtime config file is created."
                ),
            }

        try:
            current_data = _json.loads(cfg.read_text(encoding="utf-8"))
        except Exception:
            current_data = {}

        if not isinstance(current_data, dict):
            current_data = {}

        would_change = {
            k: {"old": current_data.get(k, "<not set>"), "new": v}
            for k, v in settings.items()
            if current_data.get(k) != v
        }

        if dry_run:
            return {
                "status": "dry_run",
                "config_file": str(cfg),
                "would_change": would_change,
                "note": "Set dry_run=False to apply.",
            }

        current_data.update(settings)
        cfg.write_text(
            _json.dumps(current_data, indent=2, ensure_ascii=False), encoding="utf-8"
        )

        return {
            "status": "ok",
            "config_file": str(cfg),
            "applied": would_change,
            "hint": "Restart Pimax Play for OpenXR layer changes to take effect.",
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# MSFS VR Launch Sequence
# ---------------------------------------------------------------------------

# Known executable names for Pimax (varies by version)
_PIMAX_EXE_NAMES = [
    "PimaxPlay.exe", "Pimax Play.exe", "PimaxClient.exe",
    "Pimax Client.exe", "PiTool.exe", "Pimax.exe",
]

# Known paths for VR software and MSFS
_PIMAX_PLAY_CANDIDATES = [
    r"C:\Program Files\Pimax\PimaxPlay\PimaxPlay.exe",
    r"C:\Program Files (x86)\Pimax\PimaxPlay\PimaxPlay.exe",
    r"C:\Program Files\Pimax\Pimax Play\PimaxPlay.exe",
    r"C:\Program Files\Pimax\PimaxPlay\Pimax Play.exe",
    r"C:\Program Files\Pimax\Runtime\PimaxPlay.exe",
    r"C:\Program Files\Pimax\Pimax Client\PimaxClient.exe",
    r"C:\Program Files (x86)\Pimax\Pimax Client\PimaxClient.exe",
    r"C:\Program Files\PiTool\PiTool.exe",
    r"D:\Program Files\Pimax\PimaxPlay\PimaxPlay.exe",
    r"D:\Pimax\PimaxPlay\PimaxPlay.exe",
]

_STEAMVR_CANDIDATES = [
    r"C:\Program Files (x86)\Steam\steamapps\common\SteamVR\bin\win64\vrstartup.exe",
    r"D:\Steam\steamapps\common\SteamVR\bin\win64\vrstartup.exe",
    r"D:\SteamLibrary\steamapps\common\SteamVR\bin\win64\vrstartup.exe",
    r"E:\SteamLibrary\steamapps\common\SteamVR\bin\win64\vrstartup.exe",
]

# MSFS 2024 Steam App ID
_MSFS2024_STEAM_APP_ID = "2537590"

# Known MSFS 2024 process names
_MSFS_EXE_NAMES = [
    "FlightSimulator2024.exe",
    "FlightSimulator2024.Windows.exe",
]


def _kill_msfs(timeout_s: int = 30) -> dict:
    """Kill all running MSFS processes and wait for them to exit.

    Returns dict with status information.
    """
    import time
    killed = []
    for exe in _MSFS_EXE_NAMES:
        if _is_running(exe):
            try:
                subprocess.run(
                    ["taskkill", "/IM", exe, "/F"],
                    capture_output=True, timeout=10,
                )
                killed.append(exe)
            except Exception:
                pass

    if not killed:
        return {"status": "not_running", "message": "MSFS war nicht gestartet."}

    # Wait for processes to actually exit
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        still_running = [e for e in killed if _is_running(e)]
        if not still_running:
            break
        time.sleep(1)
    else:
        still_running = [e for e in killed if _is_running(e)]
        if still_running:
            return {
                "status": "partial",
                "killed": killed,
                "still_running": still_running,
                "message": f"Nicht alle MSFS-Prozesse konnten beendet werden: {still_running}",
            }

    # Brief safety wait for OS-level cleanup (GPU resources, etc.).
    # Callers that know the cfg path should use _wait_for_file_unlocked()
    # instead of relying on this fixed delay.
    time.sleep(1)
    logger.info("MSFS killed: %s", killed)
    return {"status": "killed", "processes": killed, "message": f"MSFS beendet: {', '.join(killed)}"}


def _is_msfs_running() -> bool:
    """Check if any MSFS process is currently running."""
    return any(_is_running(exe) for exe in _MSFS_EXE_NAMES)


def _find_exe(candidates: list[str], custom: str = "") -> str | None:
    """Return the first existing path from candidates, or custom if given."""
    if custom and Path(custom).exists():
        return custom
    for p in candidates:
        if Path(p).exists():
            return p
    return None


def _find_pimax() -> str | None:
    """Find PimaxPlay.exe using multiple strategies."""
    # 1. Try known paths
    for p in _PIMAX_PLAY_CANDIDATES:
        if Path(p).exists():
            return p

    # 2. Search registry for install location
    try:
        result = _reg(
            ["query",
             r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
             "/s", "/f", "Pimax", "/d"],
            timeout=10,
        )
        for line in result.stdout.splitlines():
            m = re.search(r"InstallLocation\s+REG_SZ\s+(.+)", line)
            if m:
                install_dir = m.group(1).strip()
                for exe_name in _PIMAX_EXE_NAMES:
                    candidate = Path(install_dir) / exe_name
                    if candidate.exists():
                        return str(candidate)
    except Exception:
        pass

    # 3. Search common Program Files folders
    for root in [r"C:\Program Files", r"C:\Program Files (x86)", r"D:\Program Files"]:
        root_path = Path(root)
        if not root_path.exists():
            continue
        for pimax_dir in root_path.glob("Pimax*"):
            if pimax_dir.is_dir():
                for exe_name in _PIMAX_EXE_NAMES:
                    for hit in pimax_dir.rglob(exe_name):
                        return str(hit)

    # 4. Check Start Menu shortcuts
    start_menus = [
        Path(os.environ.get("APPDATA", "")) / r"Microsoft\Windows\Start Menu\Programs",
        Path(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"),
    ]
    for sm in start_menus:
        for lnk in sm.rglob("*Pimax*"):
            if lnk.suffix == ".lnk":
                # Read shortcut target via PowerShell
                try:
                    target = subprocess.run(
                        ["powershell", "-NoProfile", "-Command",
                         f"(New-Object -ComObject WScript.Shell).CreateShortcut('{lnk}').TargetPath"],
                        capture_output=True, text=True, timeout=5,
                    ).stdout.strip()
                    if target and Path(target).exists():
                        return target
                except Exception:
                    pass

    return None


def _is_running(process_name: str) -> bool:
    """Check if a process with the given name is running."""
    try:
        r = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {process_name}"],
            capture_output=True, text=True, timeout=10,
        )
        return process_name.lower() in r.stdout.lower()
    except Exception:
        return False


def _wait_for_process(name: str, timeout_s: int = 30) -> bool:
    """Wait until a process is running, up to timeout_s seconds."""
    import time
    for _ in range(timeout_s):
        if _is_running(name):
            return True
        time.sleep(1)
    return False


@mcp.tool()
def launch_msfs_vr(
    pimax_path: str = "",
    steamvr_path: str = "",
    skip_pimax: bool = False,
    skip_steamvr: bool = False,
    pimax_wait_s: int = 10,
    steamvr_wait_s: int = 15,
    force_restart: bool = True,
    kill_wait_s: int = 30,
) -> dict:
    """Launch MSFS 2024 in VR mode with the correct startup sequence.

    If MSFS is already running it will be closed first and then restarted
    with VR (unless force_restart is False).

    Startreihenfolge:
      1. MSFS beenden (falls es läuft)
      2. Pimax Play starten (Headset-Runtime)
      3. SteamVR starten (VR-Runtime)
      4. MSFS 2024 über Steam starten

    Jeder Schritt wartet, bis die vorherige Anwendung bereit ist.

    Args:
        pimax_path: Custom path to PimaxPlay.exe (auto-detected if empty).
        steamvr_path: Custom path to vrstartup.exe (auto-detected if empty).
        skip_pimax: Set True if you don't use Pimax (e.g. Meta Quest, Valve Index).
        skip_steamvr: Set True to skip SteamVR (e.g. for Desktop-only mode).
        pimax_wait_s: Seconds to wait for Pimax Play to initialize (default 10).
        steamvr_wait_s: Seconds to wait for SteamVR to initialize (default 15).
        force_restart: If MSFS is already running, close it and restart in VR (default True).
        kill_wait_s: Seconds to wait for MSFS to fully exit before restarting (default 30).
    """
    import time
    steps: list[dict] = []

    # --- Step 0: Close MSFS if already running ---
    if _is_msfs_running():
        if not force_restart:
            return {
                "status": "already_running",
                "message": (
                    "MSFS läuft bereits. Setze force_restart=True um es "
                    "automatisch zu beenden und in VR neu zu starten."
                ),
            }
        kill_result = _kill_msfs(timeout_s=kill_wait_s)
        steps.append({
            "step": "MSFS beenden",
            "status": kill_result["status"],
            "detail": kill_result["message"],
        })
        if kill_result["status"] == "partial":
            return {
                "error": kill_result["message"],
                "steps": steps,
                "hint": "MSFS konnte nicht vollständig beendet werden. Bitte manuell schließen.",
            }

    # --- Step 1: Pimax Play ---
    if not skip_pimax:
        pimax_running = any(
            _is_running(name) for name in _PIMAX_EXE_NAMES
        )
        if pimax_running:
            steps.append({"step": "Pimax Play", "status": "already_running"})
        else:
            pimax = _find_exe(_PIMAX_PLAY_CANDIDATES, pimax_path) if pimax_path else _find_pimax()
            if pimax is None:
                return {
                    "error": (
                        "Pimax Play konnte nicht gefunden werden. "
                        "Bitte gib den Pfad an, z.B.: "
                        "'Starte MSFS VR mit pimax_path C:\\...\\PimaxPlay.exe'"
                    ),
                }
            subprocess.Popen([pimax])
            steps.append({"step": "Pimax Play", "status": "started", "path": pimax})
            time.sleep(pimax_wait_s)

    # --- Step 2: SteamVR ---
    if not skip_steamvr:
        if _is_running("vrserver.exe") or _is_running("vrstartup.exe"):
            steps.append({"step": "SteamVR", "status": "already_running"})
        else:
            steamvr = _find_exe(_STEAMVR_CANDIDATES, steamvr_path)
            if steamvr is None:
                try:
                    subprocess.Popen(["cmd", "/c", "start", "steam://run/250820"])
                    steps.append({"step": "SteamVR", "status": "started_via_steam"})
                except Exception as exc:
                    return {"error": f"Could not start SteamVR: {exc}", "steps": steps}
            else:
                subprocess.Popen([steamvr])
                steps.append({"step": "SteamVR", "status": "started", "path": steamvr})
            ready = _wait_for_process("vrserver.exe", timeout_s=steamvr_wait_s)
            if not ready:
                steps.append({"step": "SteamVR", "warning": "vrserver.exe not detected yet, continuing anyway"})

    # --- Step 3: MSFS 2024 via Steam ---
    try:
        subprocess.Popen(["cmd", "/c", "start", f"steam://run/{_MSFS2024_STEAM_APP_ID}"])
        steps.append({"step": "MSFS 2024", "status": "started", "steam_app_id": _MSFS2024_STEAM_APP_ID})
    except Exception as exc:
        return {"error": f"Could not launch MSFS 2024: {exc}", "steps": steps}

    return {
        "status": "ok",
        "launch_sequence": steps,
        "message": (
            "Startreihenfolge abgeschlossen: "
            + " → ".join(s["step"] for s in steps)
            + ". MSFS 2024 wird in VR geladen."
        ),
    }


# ---------------------------------------------------------------------------
# MSFS Stuck / Crash Recovery
# ---------------------------------------------------------------------------

# Additional processes that can block MSFS or cause hangs
_MSFS_RELATED_PROCESSES = [
    "FlightSimulator2024.exe",
    "FlightSimulator2024.Windows.exe",
    # Common helpers / launchers that may hang
    "GameBar.exe",
    "GameBarPresenceWriter.exe",
    "EAAntiCheat.GameServiceLauncher.exe",
    "EABackgroundService.exe",
]

# MSFS 2024 cache directories (shader cache + rolling cache)
_MSFS_CACHE_DIRS = [
    # MSFS 2024 shader caches
    Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator 2024" / "ShaderCache",
    Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator 2024" / "Shadercache",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
    / "Microsoft.Limitless_8wekyb3d8bbwe" / "LocalCache" / "ShaderCache",
    # NVIDIA shader cache (shared, but clears MSFS shaders too)
    Path(os.environ.get("LOCALAPPDATA", "")) / "NVIDIA" / "DXCache",
    Path(os.environ.get("LOCALAPPDATA", "")) / "NVIDIA" / "GLCache",
    # DirectX shader cache
    Path(os.environ.get("LOCALAPPDATA", "")) / "D3DSCache",
]

# Rolling cache (user-configurable, but this is the default location)
_MSFS_ROLLING_CACHE = [
    Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator 2024" / "SceneryIndexes",
]


def _safe_rmtree(folder: Path) -> dict:
    """Delete folder contents (not the folder itself). Returns stats."""
    if not folder.exists():
        return {"path": str(folder), "status": "not_found"}
    total_size = 0
    file_count = 0
    try:
        for item in folder.iterdir():
            if item.is_file():
                total_size += item.stat().st_size
                item.unlink()
                file_count += 1
            elif item.is_dir():
                for f in item.rglob("*"):
                    if f.is_file():
                        total_size += f.stat().st_size
                        file_count += 1
                shutil.rmtree(item, ignore_errors=True)
        return {
            "path": str(folder),
            "status": "cleared",
            "files_deleted": file_count,
            "size_freed_mb": round(total_size / 1024 / 1024, 1),
        }
    except Exception as exc:
        return {"path": str(folder), "status": "error", "error": str(exc)}


@mcp.tool()
def fix_msfs(
    action: Literal[
        "restart", "kill", "clear_shader_cache", "clear_rolling_cache", "full_reset"
    ] = "restart",
    restart_in_vr: bool = True,
    skip_pimax: bool = False,
    skip_steamvr: bool = False,
) -> dict:
    """Fix a stuck, frozen, or crashing MSFS 2024.

    Use this when the user says things like:
    - "MSFS steckt" / "MSFS hängt" / "MSFS reagiert nicht"
    - "MSFS ist eingefroren" / "MSFS ist abgestürzt"
    - "Schwarzer Bildschirm in MSFS"
    - "MSFS lädt nicht" / "MSFS startet nicht"
    - "MSFS Shader Cache leeren"

    Actions:
    - restart: Force-kill MSFS + related processes, then restart (default)
    - kill: Only force-kill, don't restart
    - clear_shader_cache: Kill MSFS + delete shader caches (fixes many crashes/stutters)
    - clear_rolling_cache: Kill MSFS + delete rolling cache/scenery indexes
    - full_reset: Kill + clear ALL caches + restart (last resort for persistent problems)

    Args:
        action: What to do (see above).
        restart_in_vr: If True, restart MSFS in VR mode via launch_msfs_vr (default True).
        skip_pimax: Skip Pimax when restarting in VR.
        skip_steamvr: Skip SteamVR when restarting in VR.
    """
    import time
    steps: list[dict] = []

    # --- Step 1: Force-kill MSFS and related processes ---
    killed = []
    for proc in _MSFS_RELATED_PROCESSES:
        if _is_running(proc):
            try:
                subprocess.run(
                    ["taskkill", "/IM", proc, "/F"],
                    capture_output=True, timeout=10,
                )
                killed.append(proc)
            except Exception:
                pass

    if killed:
        steps.append({
            "step": "Prozesse beendet",
            "killed": killed,
        })
        # Wait for processes to fully exit
        deadline = time.time() + 15
        while time.time() < deadline:
            still_alive = [p for p in killed if _is_running(p)]
            if not still_alive:
                break
            time.sleep(1)
        else:
            still_alive = [p for p in killed if _is_running(p)]
            if still_alive:
                # Try harder with wmic/taskkill /T (kill process tree)
                for proc in still_alive:
                    try:
                        subprocess.run(
                            ["taskkill", "/IM", proc, "/F", "/T"],
                            capture_output=True, timeout=10,
                        )
                    except Exception:
                        pass
                time.sleep(3)
                final_check = [p for p in still_alive if _is_running(p)]
                if final_check:
                    steps.append({
                        "step": "Warnung",
                        "message": f"Prozesse konnten nicht beendet werden: {final_check}",
                    })

        # Extra wait for GPU/file handles to release
        time.sleep(5)
    else:
        steps.append({"step": "Prozesse prüfen", "message": "MSFS war nicht aktiv."})

    # --- Step 2: Clear caches if requested ---
    if action in ("clear_shader_cache", "full_reset"):
        cache_results = []
        for cache_dir in _MSFS_CACHE_DIRS:
            result = _safe_rmtree(cache_dir)
            if result["status"] != "not_found":
                cache_results.append(result)
        if cache_results:
            total_freed = sum(r.get("size_freed_mb", 0) for r in cache_results)
            steps.append({
                "step": "Shader-Cache geleert",
                "caches": cache_results,
                "total_freed_mb": round(total_freed, 1),
                "hinweis": "Erster Start nach Cache-Leerung dauert länger (Shader werden neu kompiliert).",
            })
        else:
            steps.append({"step": "Shader-Cache", "message": "Keine Cache-Ordner gefunden."})

    if action in ("clear_rolling_cache", "full_reset"):
        rolling_results = []
        for rc_dir in _MSFS_ROLLING_CACHE:
            result = _safe_rmtree(rc_dir)
            if result["status"] != "not_found":
                rolling_results.append(result)
        if rolling_results:
            steps.append({
                "step": "Rolling Cache geleert",
                "caches": rolling_results,
            })

    # --- Step 3: Restart if requested ---
    if action == "kill":
        return {
            "status": "ok",
            "steps": steps,
            "message": "MSFS wurde beendet.",
        }

    if action in ("restart", "clear_shader_cache", "clear_rolling_cache", "full_reset"):
        if restart_in_vr:
            # Use launch_msfs_vr for proper VR startup sequence
            vr_result = launch_msfs_vr(
                skip_pimax=skip_pimax,
                skip_steamvr=skip_steamvr,
                force_restart=False,  # We already killed MSFS
            )
            steps.append({
                "step": "Neustart in VR",
                "result": vr_result,
            })
        else:
            # Desktop restart
            try:
                subprocess.Popen(["cmd", "/c", "start", f"steam://run/{_MSFS2024_STEAM_APP_ID}"])
                steps.append({"step": "MSFS 2024 gestartet", "mode": "Desktop"})
            except Exception as exc:
                steps.append({"step": "Fehler", "message": f"Konnte MSFS nicht starten: {exc}"})

    # Build final message
    action_labels = {
        "restart": "MSFS beendet und neu gestartet",
        "kill": "MSFS beendet",
        "clear_shader_cache": "MSFS beendet, Shader-Cache geleert und neu gestartet",
        "clear_rolling_cache": "MSFS beendet, Rolling Cache geleert und neu gestartet",
        "full_reset": "MSFS beendet, alle Caches geleert und neu gestartet",
    }

    return {
        "status": "ok",
        "action": action,
        "steps": steps,
        "message": action_labels.get(action, action) + ".",
    }


# ===================================================================
# SYSTEM ADMINISTRATION TOOLS
# ===================================================================


def _reg(args: list[str], timeout: int = 10) -> subprocess.CompletedProcess:
    """Run a reg.exe command, trying elevated PowerShell if access is denied."""
    # First try normal
    r = subprocess.run(
        ["reg"] + args,
        capture_output=True, text=True, timeout=timeout,
    )
    if r.returncode == 0:
        return r

    # Check if access denied
    err = (r.stderr + r.stdout).lower()
    if "zugriff" in err or "access" in err or "denied" in err or "verweigert" in err:
        # Retry with elevated PowerShell (Start-Process -Verb RunAs)
        # Build reg command as a single string
        reg_cmd = "reg " + " ".join(f'"{a}"' if " " in a else a for a in args)

        # Use PowerShell to run reg with elevation and capture output
        ps_script = (
            f"$p = Start-Process -FilePath 'reg.exe' "
            f"-ArgumentList '{' '.join(args)}' "
            f"-Verb RunAs -Wait -PassThru "
            f"-WindowStyle Hidden "
            f"-RedirectStandardOutput $env:TEMP\\reg_out.txt "
            f"-RedirectStandardError $env:TEMP\\reg_err.txt; "
            f"$out = Get-Content $env:TEMP\\reg_out.txt -Raw -ErrorAction SilentlyContinue; "
            f"$err = Get-Content $env:TEMP\\reg_err.txt -Raw -ErrorAction SilentlyContinue; "
            f"Write-Output $out"
        )
        try:
            r2 = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                capture_output=True, text=True, timeout=30,
            )
            r2.returncode = 0 if r2.stdout.strip() else r2.returncode
            return r2
        except Exception:
            pass

    return r  # Return original result if elevation also fails


def _reg_query(reg_path: str, timeout: int = 10) -> subprocess.CompletedProcess:
    """Query a registry path, with auto-elevation on access denied."""
    return _reg(["query", reg_path], timeout=timeout)


def _reg_query_value(reg_path: str, value_name: str, timeout: int = 10) -> subprocess.CompletedProcess:
    """Query a specific registry value, with auto-elevation."""
    return _reg(["query", reg_path, "/v", value_name], timeout=timeout)


def _reg_add(reg_path: str, value_name: str, reg_type: str, data: str,
             timeout: int = 10) -> subprocess.CompletedProcess:
    """Add/set a registry value, with auto-elevation on access denied."""
    return _reg(["add", reg_path, "/v", value_name, "/t", reg_type,
                 "/d", data, "/f"], timeout=timeout)


def _ps(script: str, timeout: int = 30) -> str:
    """Execute a PowerShell command and return stdout."""
    r = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, timeout=timeout,
    )
    if r.returncode != 0 and r.stderr.strip():
        raise RuntimeError(r.stderr.strip())
    return r.stdout.strip()


def _ps_json(script: str, timeout: int = 60) -> list[dict]:
    """Execute PowerShell, convert output to JSON, parse and return as list."""
    raw = _ps(f"({script}) | ConvertTo-Json -Depth 4 -Compress", timeout=timeout)
    if not raw:
        return []
    try:
        data = _json.loads(raw)
    except _json.JSONDecodeError:
        logger.warning("_ps_json: could not parse PowerShell output as JSON: %.200s", raw)
        return []
    if isinstance(data, dict):
        return [data]
    return data


# ---------------------------------------------------------------------------
# 1. get_system_info
# ---------------------------------------------------------------------------

@mcp.tool()
def get_system_info() -> dict:
    """Full Windows system overview: OS, CPU, RAM, disks, network, uptime.

    Returns a comprehensive snapshot useful for diagnostics and support.
    """
    try:
        os_info = _ps_json(
            "Get-CimInstance Win32_OperatingSystem "
            "| Select-Object Caption, Version, BuildNumber, "
            "OSArchitecture, LastBootUpTime, "
            "@{N='TotalRAM_GB';E={[math]::Round($_.TotalVisibleMemorySize/1MB,2)}}, "
            "@{N='FreeRAM_GB';E={[math]::Round($_.FreePhysicalMemory/1MB,2)}}"
        )
    except Exception as exc:
        return {"error": f"Failed to query OS info: {exc}"}

    try:
        cpu_info = _ps_json(
            "Get-CimInstance Win32_Processor "
            "| Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, "
            "MaxClockSpeed, LoadPercentage"
        )
    except Exception:
        cpu_info = []

    try:
        disk_info = _ps_json(
            "Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' "
            "| Select-Object DeviceID, VolumeName, "
            "@{N='SizeGB';E={[math]::Round($_.Size/1GB,2)}}, "
            "@{N='FreeGB';E={[math]::Round($_.FreeSpace/1GB,2)}}, "
            "@{N='UsedPercent';E={[math]::Round(($_.Size-$_.FreeSpace)/$_.Size*100,1)}}"
        )
    except Exception:
        disk_info = []

    try:
        net_info = _ps_json(
            "Get-NetAdapter | Where-Object Status -eq 'Up' "
            "| Select-Object Name, InterfaceDescription, MacAddress, "
            "LinkSpeed, Status"
        )
    except Exception:
        net_info = []

    return {
        "os": os_info[0] if os_info else {},
        "cpu": cpu_info,
        "disks": disk_info,
        "network_adapters": net_info,
    }


# ---------------------------------------------------------------------------
# 2. manage_processes
# ---------------------------------------------------------------------------

@mcp.tool()
def manage_processes(
    action: Literal["list", "search", "kill"],
    name: str = "",
    pid: int = 0,
    sort_by: Literal["cpu", "memory", "name"] = "memory",
    top_n: int = 25,
) -> dict:
    """List, search or kill Windows processes.

    Args:
        action: 'list' top processes, 'search' by name, or 'kill' by PID/name.
        name: Process name or pattern (for search/kill). Supports wildcards (*).
        pid: Process ID (for kill).
        sort_by: Sort order for list ('cpu', 'memory', or 'name').
        top_n: Number of processes to return (default 25).
    """
    sort_map = {
        "cpu": "CPU",
        "memory": "WorkingSet64",
        "name": "ProcessName",
    }
    sort_field = sort_map.get(sort_by, "WorkingSet64")

    if action == "list":
        procs = _ps_json(
            f"Get-Process | Sort-Object {sort_field} -Descending "
            f"| Select-Object -First {top_n} ProcessName, Id, "
            f"@{{N='CPU_s';E={{[math]::Round($_.CPU,1)}}}}, "
            f"@{{N='MemoryMB';E={{[math]::Round($_.WorkingSet64/1MB,1)}}}}"
        )
        return {"processes": procs}

    if action == "search":
        if not name:
            return {"error": "Parameter 'name' is required for search."}
        procs = _ps_json(
            f"Get-Process | Where-Object {{$_.ProcessName -like '*{name}*'}} "
            f"| Select-Object ProcessName, Id, "
            f"@{{N='CPU_s';E={{[math]::Round($_.CPU,1)}}}}, "
            f"@{{N='MemoryMB';E={{[math]::Round($_.WorkingSet64/1MB,1)}}}}"
        )
        return {"matches": procs, "count": len(procs)}

    if action == "kill":
        if pid:
            try:
                _ps(f"Stop-Process -Id {pid} -Force")
            except RuntimeError as exc:
                return {"error": f"Could not kill PID {pid}: {exc}"}
            return {"status": "ok", "killed_pid": pid}
        if name:
            try:
                _ps(f"Stop-Process -Name '{name}' -Force")
            except RuntimeError as exc:
                return {"error": f"Could not kill process '{name}': {exc}"}
            return {"status": "ok", "killed_name": name}
        return {"error": "Provide 'pid' or 'name' to kill a process."}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 3. manage_services
# ---------------------------------------------------------------------------

@mcp.tool()
def manage_services(
    action: Literal["list", "status", "start", "stop", "restart", "set_startup"],
    name: str = "",
    startup_type: Literal["Automatic", "Manual", "Disabled"] = "Manual",
    filter_running: bool = False,
) -> dict:
    """Manage Windows services: list, start, stop, restart, change startup type.

    Args:
        action: The operation to perform.
        name: Service name (required for status/start/stop/restart/set_startup).
        startup_type: For 'set_startup' – Automatic, Manual, or Disabled.
        filter_running: For 'list' – if True, show only running services.
    """
    if action == "list":
        where = "| Where-Object Status -eq 'Running' " if filter_running else ""
        svcs = _ps_json(
            f"Get-Service {where}"
            "| Select-Object Name, DisplayName, Status, StartType"
        )
        return {"services": svcs, "count": len(svcs)}

    if not name:
        return {"error": f"Parameter 'name' is required for action '{action}'."}

    if action == "status":
        svcs = _ps_json(
            f"Get-Service -Name '{name}' "
            "| Select-Object Name, DisplayName, Status, StartType"
        )
        return {"service": svcs[0] if svcs else {}}

    if action in ("start", "stop", "restart"):
        cmd_map = {
            "start": f"Start-Service -Name '{name}'",
            "stop": f"Stop-Service -Name '{name}' -Force",
            "restart": f"Restart-Service -Name '{name}' -Force",
        }
        _ps(cmd_map[action])
        new = _ps_json(
            f"Get-Service -Name '{name}' | Select-Object Name, Status, StartType"
        )
        return {"status": "ok", "service": new[0] if new else {}}

    if action == "set_startup":
        _ps(f"Set-Service -Name '{name}' -StartupType '{startup_type}'")
        return {"status": "ok", "name": name, "startup_type": startup_type}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 4. network_diagnostics
# ---------------------------------------------------------------------------

@mcp.tool()
def network_diagnostics(
    action: Literal["ipconfig", "ping", "traceroute", "dns", "connections", "wifi"],
    target: str = "",
    count: int = 4,
) -> dict:
    """Network diagnostics: IP config, ping, traceroute, DNS, connections, Wi-Fi.

    Args:
        action: The diagnostic to run.
        target: Hostname or IP (required for ping, traceroute, dns).
        count: Number of pings (default 4).
    """
    if action == "ipconfig":
        addrs = _ps_json(
            "Get-NetIPAddress -AddressFamily IPv4 "
            "| Select-Object InterfaceAlias, IPAddress, PrefixLength"
        )
        dns = _ps_json(
            "Get-DnsClientServerAddress -AddressFamily IPv4 "
            "| Select-Object InterfaceAlias, ServerAddresses"
        )
        gateway = _ps(
            "(Get-NetRoute -DestinationPrefix '0.0.0.0/0' "
            "| Select-Object -First 1).NextHop"
        )
        return {"addresses": addrs, "dns_servers": dns, "default_gateway": gateway}

    if action == "ping":
        if not target:
            return {"error": "Parameter 'target' is required for ping."}
        results = _ps_json(
            f"Test-Connection -ComputerName '{target}' -Count {count} "
            "| Select-Object Address, "
            "@{N='ResponseTimeMs';E={$_.ResponseTime}}, StatusCode"
        )
        return {"ping_results": results, "target": target}

    if action == "traceroute":
        if not target:
            return {"error": "Parameter 'target' is required for traceroute."}
        raw = _ps(f"tracert -d -w 1000 {target}", timeout=60)
        return {"traceroute": raw, "target": target}

    if action == "dns":
        if not target:
            return {"error": "Parameter 'target' is required for dns."}
        records = _ps_json(f"Resolve-DnsName '{target}'")
        return {"dns_records": records, "target": target}

    if action == "connections":
        conns = _ps_json(
            "Get-NetTCPConnection -State Established "
            "| Select-Object LocalAddress, LocalPort, "
            "RemoteAddress, RemotePort, OwningProcess "
            "| Sort-Object RemoteAddress"
        )
        return {"connections": conns, "count": len(conns)}

    if action == "wifi":
        profiles = _ps("netsh wlan show profiles", timeout=10)
        iface = _ps("netsh wlan show interfaces", timeout=10)
        return {"profiles": profiles, "interface": iface}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 5. manage_startup_programs
# ---------------------------------------------------------------------------

@mcp.tool()
def manage_startup_programs(
    action: Literal["list", "disable", "enable"],
    name: str = "",
) -> dict:
    """View and manage Windows startup programs (autostart entries).

    Reads from the Run keys in the registry (HKLM + HKCU) and the
    common Startup folder.

    Args:
        action: 'list' all, 'disable' by name, or 'enable' a previously disabled entry.
        name: Name/key of the startup entry (required for disable/enable).
    """
    hklm = r"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    hkcu = r"HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

    if action == "list":
        entries: list[dict] = []
        for hive, label in [(hklm, "HKLM"), (hkcu, "HKCU")]:
            try:
                items = _ps_json(
                    f"Get-ItemProperty '{hive}' "
                    "| ForEach-Object {{ $_.PSObject.Properties "
                    "| Where-Object {{ $_.Name -notlike 'PS*' }} "
                    "| Select-Object Name, Value }}"
                )
                for item in items:
                    item["hive"] = label
                entries.extend(items)
            except Exception:
                pass

        # Also check shell:startup folder
        try:
            startup_folder = _ps(
                "[Environment]::GetFolderPath('Startup')"
            )
            if startup_folder:
                shortcuts = _ps_json(
                    f"Get-ChildItem '{startup_folder}' "
                    "| Select-Object Name, FullName"
                )
                for s in shortcuts:
                    s["hive"] = "StartupFolder"
                entries.extend(shortcuts)
        except Exception:
            pass

        return {"startup_entries": entries, "count": len(entries)}

    if not name:
        return {"error": f"Parameter 'name' is required for '{action}'."}

    if action == "disable":
        for hive in [hklm, hkcu]:
            try:
                _ps(f"Remove-ItemProperty -Path '{hive}' -Name '{name}' -ErrorAction Stop")
                return {"status": "ok", "disabled": name, "hive": hive}
            except Exception:
                continue
        return {"error": f"Startup entry '{name}' not found in registry Run keys."}

    if action == "enable":
        return {
            "error": (
                "To re-enable a startup entry, use run_shell_command to set the "
                "registry value, e.g.: "
                f"New-ItemProperty -Path '{hkcu}' -Name '{name}' "
                "-Value 'C:\\path\\to\\program.exe'"
            )
        }

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 6. set_power_plan
# ---------------------------------------------------------------------------

@mcp.tool()
def set_power_plan(
    action: Literal["list", "set"],
    plan: Literal["high_performance", "balanced", "power_saver", "ultimate"] = "balanced",
) -> dict:
    """List or switch Windows power plans.

    Args:
        action: 'list' available plans or 'set' the active plan.
        plan: Plan to activate – 'high_performance', 'balanced',
              'power_saver', or 'ultimate' (Ultimate Performance).
    """
    if action == "list":
        raw = _ps("powercfg /list")
        return {"power_plans": raw}

    guid_map = {
        "high_performance": "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c",
        "balanced": "381b4222-f694-41f0-9685-ff5bb260df2e",
        "power_saver": "a1841308-3541-4fab-bc81-f71556f20b4a",
        "ultimate": "e9a42b02-d5df-448d-aa00-03f14749eb61",
    }
    guid = guid_map.get(plan)
    if not guid:
        return {"error": f"Unknown plan: {plan}"}

    try:
        _ps(f"powercfg /setactive {guid}")
    except RuntimeError:
        # Ultimate might not exist – try to add it first
        if plan == "ultimate":
            _ps(f"powercfg -duplicatescheme {guid}")
            _ps(f"powercfg /setactive {guid}")
        else:
            raise

    current = _ps("powercfg /getactivescheme")
    return {"status": "ok", "active_plan": current}


# ---------------------------------------------------------------------------
# 7. manage_firewall
# ---------------------------------------------------------------------------

@mcp.tool()
def manage_firewall(
    action: Literal["status", "list_rules", "add_rule", "remove_rule", "enable", "disable"],
    rule_name: str = "",
    program: str = "",
    port: int = 0,
    protocol: Literal["TCP", "UDP", "Any"] = "TCP",
    direction: Literal["Inbound", "Outbound"] = "Inbound",
    allow: bool = True,
) -> dict:
    """Manage Windows Firewall: view status, list/add/remove rules.

    Args:
        action: Operation to perform.
        rule_name: Display name for the rule (required for add/remove).
        program: Path to .exe (for add_rule – program-based rule).
        port: Port number (for add_rule – port-based rule).
        protocol: TCP, UDP, or Any (for port rules).
        direction: Inbound or Outbound.
        allow: True = Allow, False = Block (for add_rule).
    """
    if action == "status":
        profiles = _ps_json(
            "Get-NetFirewallProfile | Select-Object Name, Enabled, "
            "DefaultInboundAction, DefaultOutboundAction"
        )
        return {"firewall_profiles": profiles}

    if action == "list_rules":
        rules = _ps_json(
            f"Get-NetFirewallRule -Direction {direction} -Enabled True "
            "| Select-Object -First 50 DisplayName, Direction, Action, "
            "Enabled, Profile"
        )
        return {"rules": rules, "direction": direction, "count": len(rules)}

    if action == "add_rule":
        if not rule_name:
            return {"error": "Parameter 'rule_name' is required."}
        fw_action = "Allow" if allow else "Block"
        cmd = (
            f"New-NetFirewallRule -DisplayName '{rule_name}' "
            f"-Direction {direction} -Action {fw_action}"
        )
        if port:
            cmd += f" -Protocol {protocol} -LocalPort {port}"
        if program:
            cmd += f" -Program '{program}'"
        _ps(cmd)
        return {
            "status": "ok",
            "rule": rule_name,
            "action": fw_action,
            "direction": direction,
        }

    if action == "remove_rule":
        if not rule_name:
            return {"error": "Parameter 'rule_name' is required."}
        _ps(f"Remove-NetFirewallRule -DisplayName '{rule_name}'")
        return {"status": "ok", "removed": rule_name}

    if action in ("enable", "disable"):
        enabled = "True" if action == "enable" else "False"
        _ps(
            f"Set-NetFirewallProfile -Profile Domain,Public,Private "
            f"-Enabled {enabled}"
        )
        return {"status": "ok", "firewall_enabled": action == "enable"}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 8. disk_analysis
# ---------------------------------------------------------------------------

@mcp.tool()
def disk_analysis(
    action: Literal["usage", "large_files", "temp_cleanup"],
    path: str = "C:\\",
    top_n: int = 20,
    min_size_mb: int = 100,
) -> dict:
    """Analyse disk usage, find large files, or clean temp folders.

    Args:
        action: 'usage' for disk overview, 'large_files' to find big files,
                'temp_cleanup' to delete temp files.
        path: Root path for large_files scan (default C:\\).
        top_n: Number of large files to return.
        min_size_mb: Minimum file size in MB for large_files.
    """
    if action == "usage":
        disks = _ps_json(
            "Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' "
            "| Select-Object DeviceID, VolumeName, "
            "@{N='SizeGB';E={[math]::Round($_.Size/1GB,2)}}, "
            "@{N='FreeGB';E={[math]::Round($_.FreeSpace/1GB,2)}}, "
            "@{N='UsedPercent';E={[math]::Round(($_.Size-$_.FreeSpace)/$_.Size*100,1)}}"
        )
        return {"disks": disks}

    if action == "large_files":
        files = _ps_json(
            f"Get-ChildItem -Path '{path}' -Recurse -File -ErrorAction SilentlyContinue "
            f"| Where-Object {{$_.Length -gt {min_size_mb}MB}} "
            f"| Sort-Object Length -Descending "
            f"| Select-Object -First {top_n} FullName, "
            f"@{{N='SizeMB';E={{[math]::Round($_.Length/1MB,1)}}}}, LastWriteTime",
            timeout=120,
        )
        return {"large_files": files, "count": len(files), "scanned_path": path}

    if action == "temp_cleanup":
        result = _ps(
            "$paths = @($env:TEMP, 'C:\\Windows\\Temp');\n"
            "$before = 0; $freed = 0;\n"
            "foreach ($p in $paths) {\n"
            "  $items = Get-ChildItem $p -Recurse -ErrorAction SilentlyContinue;\n"
            "  $before += ($items | Measure-Object Length -Sum).Sum;\n"
            "  Remove-Item \"$p\\*\" -Recurse -Force -ErrorAction SilentlyContinue\n"
            "}\n"
            "$after = 0;\n"
            "foreach ($p in $paths) {\n"
            "  $after += (Get-ChildItem $p -Recurse -ErrorAction SilentlyContinue "
            "| Measure-Object Length -Sum).Sum\n"
            "}\n"
            "[math]::Round(($before - $after)/1MB, 1)",
            timeout=120,
        )
        return {"status": "ok", "freed_mb": float(result) if result else 0}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 9. windows_update_status
# ---------------------------------------------------------------------------

@mcp.tool()
def windows_update_status(
    action: Literal["check", "history"] = "history",
    count: int = 15,
) -> dict:
    """Check Windows Update status or view update history.

    Args:
        action: 'history' for recent updates, 'check' for pending updates.
        count: Number of history entries to return.
    """
    if action == "history":
        updates = _ps_json(
            f"Get-HotFix | Sort-Object InstalledOn -Descending "
            f"| Select-Object -First {count} HotFixID, Description, InstalledOn"
        )
        return {"recent_updates": updates, "count": len(updates)}

    if action == "check":
        try:
            pending = _ps(
                "$Session = New-Object -ComObject Microsoft.Update.Session;\n"
                "$Searcher = $Session.CreateUpdateSearcher();\n"
                "$Results = $Searcher.Search('IsInstalled=0');\n"
                "$Results.Updates | ForEach-Object {\n"
                "  [PSCustomObject]@{\n"
                "    Title=$_.Title;\n"
                "    IsMandatory=$_.IsMandatory;\n"
                "    IsDownloaded=$_.IsDownloaded\n"
                "  }\n"
                "} | ConvertTo-Json -Compress",
                timeout=120,
            )
            if not pending:
                return {"pending_updates": [], "message": "System is up to date."}
            data = _json.loads(pending)
            if isinstance(data, dict):
                data = [data]
            return {"pending_updates": data, "count": len(data)}
        except Exception as exc:
            return {"error": f"Could not check for updates: {exc}"}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 10. manage_installed_software
# ---------------------------------------------------------------------------

@mcp.tool()
def manage_installed_software(
    action: Literal["list", "search", "uninstall"],
    name: str = "",
) -> dict:
    """List, search or uninstall software on the system.

    Uses the Uninstall registry keys for fast enumeration (not Win32_Product).

    Args:
        action: 'list' all, 'search' by name, or 'uninstall' by name.
        name: Software name or pattern (for search/uninstall).
    """
    reg_paths = [
        r"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        r"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    ]
    query = " + ".join(
        f"@(Get-ItemProperty '{p}' -ErrorAction SilentlyContinue "
        f"| Where-Object {{ $_.DisplayName -ne $null }})"
        for p in reg_paths
    )

    if action == "list":
        sw = _ps_json(
            f"({query}) | Sort-Object DisplayName "
            "| Select-Object DisplayName, DisplayVersion, Publisher, InstallDate",
            timeout=30,
        )
        return {"software": sw, "count": len(sw)}

    if not name:
        return {"error": f"Parameter 'name' is required for '{action}'."}

    if action == "search":
        sw = _ps_json(
            f"({query}) | Where-Object {{$_.DisplayName -like '*{name}*'}} "
            "| Select-Object DisplayName, DisplayVersion, Publisher, "
            "InstallDate, UninstallString",
            timeout=30,
        )
        return {"matches": sw, "count": len(sw)}

    if action == "uninstall":
        matches = _ps_json(
            f"({query}) | Where-Object {{$_.DisplayName -like '*{name}*'}} "
            "| Select-Object DisplayName, UninstallString",
            timeout=30,
        )
        if not matches:
            return {"error": f"No software found matching '{name}'."}
        if len(matches) > 1:
            return {
                "error": "Multiple matches found. Be more specific.",
                "matches": [m.get("DisplayName") for m in matches],
            }
        uninstall_cmd = matches[0].get("UninstallString", "")
        if not uninstall_cmd:
            return {"error": "No uninstall command found for this software."}
        # Run uninstaller silently
        if "msiexec" in uninstall_cmd.lower():
            uninstall_cmd = uninstall_cmd.replace("/I", "/X") + " /quiet /norestart"
        _ps(f"Start-Process cmd -ArgumentList '/c {uninstall_cmd}' -Wait", timeout=120)
        return {
            "status": "ok",
            "uninstalled": matches[0].get("DisplayName"),
        }

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 11. manage_users
# ---------------------------------------------------------------------------

@mcp.tool()
def manage_users(
    action: Literal["list", "whoami", "add", "remove", "set_password"],
    username: str = "",
    password: str = "",
    admin: bool = False,
) -> dict:
    """Manage local Windows user accounts.

    Args:
        action: 'list', 'whoami', 'add', 'remove', or 'set_password'.
        username: Username (required for add/remove/set_password).
        password: Password (required for add/set_password).
        admin: If True, add the user to the Administrators group (for 'add').
    """
    if action == "whoami":
        info = _ps(
            "$u = [System.Security.Principal.WindowsIdentity]::GetCurrent();\n"
            "$p = New-Object System.Security.Principal.WindowsPrincipal($u);\n"
            "$isAdmin = $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator);\n"
            "\"User: $($u.Name) | Admin: $isAdmin\""
        )
        return {"info": info}

    if action == "list":
        users = _ps_json(
            "Get-LocalUser | Select-Object Name, Enabled, "
            "LastLogon, PasswordRequired, Description"
        )
        return {"users": users}

    if not username:
        return {"error": f"Parameter 'username' is required for '{action}'."}

    if action == "add":
        if not password:
            return {"error": "Parameter 'password' is required to add a user."}
        _ps(
            f"$pw = ConvertTo-SecureString '{password}' -AsPlainText -Force;\n"
            f"New-LocalUser -Name '{username}' -Password $pw -FullName '{username}'"
        )
        if admin:
            _ps(f"Add-LocalGroupMember -Group 'Administrators' -Member '{username}'")
        return {"status": "ok", "created": username, "admin": admin}

    if action == "remove":
        _ps(f"Remove-LocalUser -Name '{username}'")
        return {"status": "ok", "removed": username}

    if action == "set_password":
        if not password:
            return {"error": "Parameter 'password' is required."}
        _ps(
            f"$pw = ConvertTo-SecureString '{password}' -AsPlainText -Force;\n"
            f"Set-LocalUser -Name '{username}' -Password $pw"
        )
        return {"status": "ok", "username": username, "password_updated": True}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 12. manage_scheduled_tasks
# ---------------------------------------------------------------------------

@mcp.tool()
def manage_scheduled_tasks(
    action: Literal["list", "create", "delete", "run", "disable", "enable"],
    task_name: str = "",
    program: str = "",
    arguments: str = "",
    schedule: Literal["daily", "weekly", "hourly", "at_logon", "once"] = "daily",
    time: str = "09:00",
) -> dict:
    """Manage Windows Task Scheduler entries.

    Args:
        action: 'list', 'create', 'delete', 'run', 'disable', or 'enable'.
        task_name: Name of the task (required for all except list).
        program: Path to executable (required for create).
        arguments: Command-line arguments for the program.
        schedule: Trigger type for create.
        time: Time of day for the trigger (HH:MM format).
    """
    if action == "list":
        tasks = _ps_json(
            "Get-ScheduledTask | Where-Object {$_.TaskPath -eq '\\\\'} "
            "| Select-Object TaskName, State, "
            "@{N='Triggers';E={($_.Triggers | ForEach-Object {$_.CimClass.CimClassName}) -join ', '}}",
            timeout=30,
        )
        return {"tasks": tasks, "count": len(tasks)}

    if not task_name:
        return {"error": f"Parameter 'task_name' is required for '{action}'."}

    if action == "create":
        if not program:
            return {"error": "Parameter 'program' is required for 'create'."}
        trigger_map = {
            "daily": f"New-ScheduledTaskTrigger -Daily -At '{time}'",
            "weekly": f"New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At '{time}'",
            "hourly": "New-ScheduledTaskTrigger -Once -At '00:00' "
                      "-RepetitionInterval (New-TimeSpan -Hours 1)",
            "at_logon": "New-ScheduledTaskTrigger -AtLogon",
            "once": f"New-ScheduledTaskTrigger -Once -At '{time}'",
        }
        trigger_cmd = trigger_map.get(schedule, trigger_map["daily"])
        args_part = f" -Argument '{arguments}'" if arguments else ""
        _ps(
            f"$action = New-ScheduledTaskAction -Execute '{program}'{args_part};\n"
            f"$trigger = {trigger_cmd};\n"
            f"Register-ScheduledTask -TaskName '{task_name}' "
            f"-Action $action -Trigger $trigger -RunLevel Highest"
        )
        return {"status": "ok", "created": task_name, "schedule": schedule}

    if action == "delete":
        _ps(f"Unregister-ScheduledTask -TaskName '{task_name}' -Confirm:$false")
        return {"status": "ok", "deleted": task_name}

    if action == "run":
        _ps(f"Start-ScheduledTask -TaskName '{task_name}'")
        return {"status": "ok", "started": task_name}

    if action == "disable":
        _ps(f"Disable-ScheduledTask -TaskName '{task_name}'")
        return {"status": "ok", "disabled": task_name}

    if action == "enable":
        _ps(f"Enable-ScheduledTask -TaskName '{task_name}'")
        return {"status": "ok", "enabled": task_name}

    return {"error": f"Unknown action: {action}"}


# ---------------------------------------------------------------------------
# 13. run_shell_command  (catch-all for everything else)
# ---------------------------------------------------------------------------

@mcp.tool()
def run_shell_command(
    command: str,
    use_powershell: bool = True,
    timeout: int = 30,
) -> dict:
    """Execute an arbitrary shell command (PowerShell or CMD).

    Use this as a fallback when no other tool covers the task.
    Requires care – commands run with the same privileges as the MCP server.

    Args:
        command: The command string to execute.
        use_powershell: True for PowerShell (default), False for CMD.
        timeout: Max seconds to wait (default 30).
    """
    try:
        if use_powershell:
            r = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
                capture_output=True, text=True, timeout=timeout,
            )
        else:
            r = subprocess.run(
                ["cmd", "/c", command],
                capture_output=True, text=True, timeout=timeout,
            )
        return {
            "stdout": r.stdout.strip(),
            "stderr": r.stderr.strip(),
            "returncode": r.returncode,
        }
    except subprocess.TimeoutExpired:
        return {"error": f"Command timed out after {timeout}s."}
    except Exception as exc:
        return {"error": str(exc)}


# ---------------------------------------------------------------------------
# Browser control via Chrome DevTools Protocol (CDP)
# Same protocol used by the Chrome extension – no AppleScript needed.
# ---------------------------------------------------------------------------

import glob as _glob
import time as _time
import threading as _threading

try:
    import websocket as _websocket          # websocket-client package
    _WS_AVAILABLE = True
except ImportError:
    _WS_AVAILABLE = False

# CDP debug port – Chrome must be launched with --remote-debugging-port=9222
_CDP_PORT = 9222
_CDP_BASE = f"http://localhost:{_CDP_PORT}"

# ── Low-level CDP helpers ──────────────────────────────────────────────────────

def _cdp_ensure_chrome() -> bool:
    """Ensure Chrome is reachable on the CDP debug port (Windows).

    Logic:
    1. If CDP is already available → done.
    2. If Chrome is running WITHOUT the debug port → terminate it via taskkill,
       then relaunch with --remote-debugging-port.
    3. If Chrome is not running at all → launch it fresh with the flag.
    Returns True if CDP becomes available, False otherwise.
    """
    # ── 1. Already available? ─────────────────────────────────────────────────
    try:
        resp = httpx.get(f"{_CDP_BASE}/json/version", timeout=2)
        if resp.status_code == 200:
            return True
    except Exception:
        pass

    # ── Windows Chrome install locations ─────────────────────────────────────
    _CHROME_PATHS = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        str(Path(os.environ.get("LOCALAPPDATA", ""))
            / "Google" / "Chrome" / "Application" / "chrome.exe"),
        r"C:\Program Files\Chromium\Application\chrome.exe",
    ]
    chrome_bin = next((p for p in _CHROME_PATHS if Path(p).exists()), None)
    if not chrome_bin:
        return False   # Chrome not installed

    # ── 2. Chrome running without debug port → terminate via taskkill ────────
    try:
        chk = subprocess.run(
            ["tasklist", "/fi", "IMAGENAME eq chrome.exe", "/fo", "csv", "/nh"],
            capture_output=True, text=True,
        )
        if "chrome.exe" in chk.stdout.lower():
            # Gracefully close all Chrome windows first, then force-kill remainder
            subprocess.run(
                ["taskkill", "/im", "chrome.exe"],
                capture_output=True,
            )
            _time.sleep(1)
            subprocess.run(
                ["taskkill", "/f", "/im", "chrome.exe"],
                capture_output=True,
            )
            _time.sleep(1)
    except Exception:
        pass

    # ── 3. Launch Chrome with remote-debugging-port ───────────────────────────
    subprocess.Popen(
        [chrome_bin,
         f"--remote-debugging-port={_CDP_PORT}",
         "--no-first-run",
         "--no-default-browser-check",
         "--restore-last-session"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        **( {"creationflags": subprocess.CREATE_NO_WINDOW}
            if platform.system() == "Windows" else {} ),
    )

    # Wait up to 8 s for CDP to become available
    for _ in range(16):
        _time.sleep(0.5)
        try:
            resp = httpx.get(f"{_CDP_BASE}/json/version", timeout=1)
            if resp.status_code == 200:
                return True
        except Exception:
            pass

    return False


def _cdp_list_tabs() -> list[dict]:
    """Return all open Chrome tabs via the CDP HTTP API."""
    try:
        resp = httpx.get(f"{_CDP_BASE}/json", timeout=3)
        return [t for t in resp.json() if t.get("type") == "page"]
    except Exception:
        return []


def _cdp_new_tab(url: str) -> dict | None:
    """Open a new tab at *url* and return its CDP descriptor."""
    try:
        resp = httpx.get(f"{_CDP_BASE}/json/new?{url}", timeout=5)
        return resp.json()
    except Exception:
        return None


def _cdp_exec(ws_url: str, method: str, params: dict | None = None,
              timeout: int = 15) -> dict:
    """Send a single CDP command over WebSocket and return the result.
    Falls back to AppleScript if websocket-client is not installed."""
    if not _WS_AVAILABLE:
        raise RuntimeError(
            "websocket-client not installed. Run: pip install websocket-client"
        )
    params = params or {}
    cmd_id = 1
    payload = _json.dumps({"id": cmd_id, "method": method, "params": params})
    result: dict = {}

    def _run():
        nonlocal result
        try:
            ws = _websocket.create_connection(ws_url, timeout=timeout)
            ws.send(payload)
            deadline = _time.time() + timeout
            while _time.time() < deadline:
                raw = ws.recv()
                msg = _json.loads(raw)
                if msg.get("id") == cmd_id:
                    result = msg.get("result", {})
                    break
            ws.close()
        except Exception as exc:
            result = {"__error": str(exc)}

    t = _threading.Thread(target=_run, daemon=True)
    t.start()
    t.join(timeout + 2)
    return result


def _cdp_js(ws_url: str, js: str, timeout: int = 15) -> str:
    """Execute *js* in the tab at *ws_url* and return the string result."""
    res = _cdp_exec(ws_url, "Runtime.evaluate",
                    {"expression": js, "returnByValue": True, "awaitPromise": False},
                    timeout=timeout)
    if "__error" in res:
        return f"cdp-error:{res['__error']}"
    val = res.get("result", {})
    if val.get("type") == "string":
        return val.get("value", "")
    return str(val.get("value", ""))


def _cdp_navigate(ws_url: str, url: str, timeout: int = 15) -> None:
    """Navigate the tab to *url* via CDP."""
    _cdp_exec(ws_url, "Page.navigate", {"url": url}, timeout=timeout)
    _time.sleep(3)   # wait for DOMContentLoaded


def _get_active_tab_ws() -> str | None:
    """Return the WebSocket debugger URL of the currently active (foremost) tab."""
    tabs = _cdp_list_tabs()
    if not tabs:
        return None
    # Prefer the most recently opened visible page tab
    return tabs[0].get("webSocketDebuggerUrl")


# ── Public browser helpers (used by analyse / strategy functions) ─────────────

def _chrome_open_url(url: str) -> str | None:
    """Open *url* in Chrome via CDP. Returns the tab's WebSocket debugger URL."""
    if not _cdp_ensure_chrome():
        raise RuntimeError(
            f"Chrome is not reachable on port {_CDP_PORT}. "
            "Start Chrome with: --remote-debugging-port=9222"
        )
    tab = _cdp_new_tab(url)
    if tab and tab.get("webSocketDebuggerUrl"):
        _time.sleep(3)
        return tab["webSocketDebuggerUrl"]
    # Fallback: navigate in the first existing tab
    tabs = _cdp_list_tabs()
    if tabs:
        ws = tabs[0].get("webSocketDebuggerUrl", "")
        _cdp_navigate(ws, url)
        return ws
    return None


def _chrome_js(js: str, ws_url: str | None = None) -> str:
    """Execute *js* in the active Chrome tab via CDP.
    If *ws_url* is None the currently active tab is used."""
    if ws_url is None:
        ws_url = _get_active_tab_ws()
    if not ws_url:
        return "error:no-tab"
    return _cdp_js(ws_url, js)


def _latest_download(downloads_dir: Path, before: set[Path], timeout: int = 90) -> Path | None:
    """Poll *downloads_dir* until a new completed file appears (no .crdownload / .part)."""
    deadline = _time.time() + timeout
    while _time.time() < deadline:
        current = {
            p for p in downloads_dir.iterdir()
            if p.is_file()
            and p.suffix.lower() not in {".crdownload", ".part", ".tmp"}
            and not p.name.startswith(".")
        }
        new = current - before
        if new:
            return max(new, key=lambda p: p.stat().st_mtime)
        _time.sleep(1.5)
    return None


@mcp.tool()
def ensure_chrome_debug() -> dict:
    """Start (or restart) Google Chrome with --remote-debugging-port=9222.

    Called automatically by download_mod_via_browser, but you can also call
    this tool manually to prepare Chrome before triggering a download.

    • If Chrome is already running WITH the debug port → nothing changes.
    • If Chrome is running WITHOUT the debug port → Chrome is quit and
      relaunched with the flag (previous session is restored).
    • If Chrome is not running → launched fresh with the flag.
    """
    ok = _cdp_ensure_chrome()
    if ok:
        try:
            version = httpx.get(f"{_CDP_BASE}/json/version", timeout=2).json()
            tabs    = _cdp_list_tabs()
            return {
                "status": "ready",
                "browser": version.get("Browser", "unknown"),
                "cdp_port": _CDP_PORT,
                "open_tabs": len(tabs),
            }
        except Exception:
            return {"status": "ready", "cdp_port": _CDP_PORT}
    return {
        "status": "failed",
        "message": (
            "Chrome could not be started with --remote-debugging-port=9222. "
            r"Please make sure Google Chrome is installed (e.g. C:\Program Files\Google\Chrome\Application\chrome.exe)."
        ),
    }


@mcp.tool()
def browser_navigate(
    url: str,
    wait_seconds: int = 3,
) -> dict:
    """Open a URL in Chrome and return page info.

    Use this to navigate Chrome to any URL — e.g. to open a download page,
    login page, or any website. Chrome is started automatically with CDP
    if not already running.

    Args:
        url: The URL to open.
        wait_seconds: Seconds to wait for page load (default 3).
    """
    if not _cdp_ensure_chrome():
        return {
            "error": "Chrome konnte nicht gestartet werden. Ist Chrome installiert?"
        }

    ws = _chrome_open_url(url)
    if not ws:
        return {"error": "Tab konnte nicht geöffnet werden."}

    if wait_seconds > 3:
        _time.sleep(wait_seconds - 3)  # _chrome_open_url already waits 3s

    # Read page state
    page_info = _cdp_js(ws, """
    JSON.stringify({
        url: window.location.href,
        title: document.title,
        readyState: document.readyState,
        hasLogin: !!(document.querySelector('[href*="login"],[href*="signin"],'
            + '[href*="anmelden"],input[type="password"]')),
        bodyLength: (document.body && document.body.innerText || "").length
    })
    """)

    try:
        info = _json.loads(page_info)
    except Exception:
        info = {"url": url, "raw": page_info}

    return {
        "status": "ok",
        "tab_ws": ws,
        **info,
    }


@mcp.tool()
def browser_click(
    selector: str = "",
    text: str = "",
    tab_ws: str = "",
) -> dict:
    """Click an element in the current Chrome tab.

    Use CSS selector or text content to find the element.

    Args:
        selector: CSS selector (e.g. '#download-btn', '.btn-primary').
        text: Text content to search for (e.g. 'Download', 'Login').
              If both selector and text are given, selector takes priority.
        tab_ws: WebSocket URL of the tab (from browser_navigate). Empty = active tab.
    """
    if not selector and not text:
        return {"error": "Entweder 'selector' oder 'text' muss angegeben werden."}

    ws = tab_ws or _get_active_tab_ws()
    if not ws:
        return {"error": "Kein Chrome-Tab verfügbar. Erst browser_navigate aufrufen."}

    if selector:
        js_code = f"""
        (function(){{
            var el = document.querySelector('{selector}');
            if(el) {{ el.click(); return 'clicked: ' + (el.innerText||'').trim().substring(0,50); }}
            return 'not-found: {selector}';
        }})()
        """
    else:
        safe_text = text.replace("'", "\\'")
        js_code = f"""
        (function(){{
            var all = document.querySelectorAll('a,button,[role="button"],input[type="submit"]');
            for(var i=0; i<all.length; i++){{
                var t = (all[i].innerText||all[i].value||'').trim().toLowerCase();
                if(t.includes('{safe_text}'.toLowerCase())){{
                    all[i].click();
                    return 'clicked: ' + (all[i].innerText||all[i].value||'').trim().substring(0,50);
                }}
            }}
            return 'not-found: {safe_text}';
        }})()
        """

    result = _cdp_js(ws, js_code)
    return {"status": "ok" if "clicked" in result else "not_found", "result": result}


@mcp.tool()
def browser_type(
    selector: str,
    text: str,
    submit: bool = False,
    tab_ws: str = "",
) -> dict:
    """Type text into an input field in Chrome.

    Useful for filling login forms, search boxes, etc.

    Args:
        selector: CSS selector for the input (e.g. 'input[name="email"]', '#password').
        text: Text to type into the field.
        submit: If True, press Enter after typing (submits the form).
        tab_ws: WebSocket URL of the tab. Empty = active tab.
    """
    ws = tab_ws or _get_active_tab_ws()
    if not ws:
        return {"error": "Kein Chrome-Tab verfügbar."}

    safe_text = text.replace("\\", "\\\\").replace("'", "\\'")
    submit_code = (
        "el.form && el.form.submit();"
        if submit else ""
    )
    js_code = f"""
    (function(){{
        var el = document.querySelector('{selector}');
        if(!el) return 'not-found: {selector}';
        el.focus();
        el.value = '{safe_text}';
        el.dispatchEvent(new Event('input', {{bubbles:true}}));
        el.dispatchEvent(new Event('change', {{bubbles:true}}));
        {submit_code}
        return 'typed: ' + '{safe_text}'.substring(0,3) + '***';
    }})()
    """

    result = _cdp_js(ws, js_code)
    return {"status": "ok" if "typed" in result else "not_found", "result": result}


@mcp.tool()
def browser_read_page(
    tab_ws: str = "",
) -> dict:
    """Read the current page content from Chrome.

    Returns URL, title, visible text, and form fields.
    Useful to check if a login was successful or what's on the page.

    Args:
        tab_ws: WebSocket URL of the tab. Empty = active tab.
    """
    ws = tab_ws or _get_active_tab_ws()
    if not ws:
        return {"error": "Kein Chrome-Tab verfügbar."}

    js_code = r"""
    (function(){
        var inputs = [];
        document.querySelectorAll('input,select,textarea').forEach(function(el){
            if(el.type === 'hidden') return;
            inputs.push({
                tag: el.tagName.toLowerCase(),
                type: el.type || '',
                name: el.name || '',
                id: el.id || '',
                placeholder: el.placeholder || '',
                value: el.type === 'password' ? '***' : (el.value||'').substring(0,50)
            });
        });

        var links = [];
        document.querySelectorAll('a[href]').forEach(function(a){
            var t = (a.innerText||'').trim();
            if(t && t.length < 80) links.push({text:t, href:a.href});
        });

        return JSON.stringify({
            url: window.location.href,
            title: document.title,
            bodyText: (document.body && document.body.innerText || '').substring(0, 2000),
            inputs: inputs.slice(0, 20),
            links: links.slice(0, 30)
        });
    })()
    """

    raw = _cdp_js(ws, js_code)
    try:
        return _json.loads(raw)
    except Exception:
        return {"url": "unknown", "raw": raw}


# ---------------------------------------------------------------------------
# JS snippets used by the browser-download tools
# ---------------------------------------------------------------------------

# Clicks the first visible "Download" / "Free Download" button on the page.
_JS_CLICK_MAIN_DOWNLOAD = r"""
(function(){
  var selectors = [
    '[data-action="download"]',
    'button[id*="download"]','a[id*="download"]',
    'a.download-btn','button.download-btn',
    '.btn-download','#btn-download',
    'a[href*="/download"]','a[href*="download"]'
  ];
  for(var i=0;i<selectors.length;i++){
    var els=document.querySelectorAll(selectors[i]);
    for(var j=0;j<els.length;j++){
      var t=(els[j].textContent||"").toLowerCase().trim();
      if(t.includes("download")||t===""){els[j].click();return"clicked:"+selectors[i];}
    }
  }
  var all=document.querySelectorAll("a,button");
  for(var k=0;k<all.length;k++){
    if((all[k].innerText||"").toLowerCase().includes("download")){
      all[k].click();return"fallback:"+all[k].innerText.trim().substring(0,50);
    }
  }
  return"not-found";
})()
"""

# Detects whether a modal/dialog is currently visible and returns its info.
_JS_DETECT_MODAL = r"""
(function(){
  var modalSelectors=[
    '.modal','[role="dialog"]','[aria-modal="true"]',
    '.dialog','#download-modal','[class*="modal"]','[class*="dialog"]',
    '[class*="popup"]','[class*="overlay"]'
  ];
  for(var i=0;i<modalSelectors.length;i++){
    var els=document.querySelectorAll(modalSelectors[i]);
    for(var j=0;j<els.length;j++){
      var el=els[j];
      var style=window.getComputedStyle(el);
      if(style.display!=="none"&&style.visibility!=="hidden"&&style.opacity!=="0"){
        // Count download buttons inside the modal
        var btns=el.querySelectorAll("a,button");
        var dlBtns=[];
        for(var k=0;k<btns.length;k++){
          if((btns[k].innerText||btns[k].href||"").toLowerCase().includes("download"))
            dlBtns.push(btns[k].innerText.trim().substring(0,60)+"||"+(btns[k].href||""));
        }
        return JSON.stringify({found:true,selector:modalSelectors[i],downloadButtons:dlBtns});
      }
    }
  }
  return JSON.stringify({found:false});
})()
"""

# Clicks ALL download buttons inside any visible modal, one per row/item.
# Returns JSON array of what was clicked.
_JS_CLICK_ALL_MODAL_DOWNLOADS = r"""
(function(){
  var modalSelectors=[
    '.modal','[role="dialog"]','[aria-modal="true"]',
    '.dialog','[class*="modal"]','[class*="dialog"]',
    '[class*="popup"]','[class*="overlay"]'
  ];
  var clicked=[];
  for(var i=0;i<modalSelectors.length;i++){
    var modals=document.querySelectorAll(modalSelectors[i]);
    for(var j=0;j<modals.length;j++){
      var modal=modals[j];
      var style=window.getComputedStyle(modal);
      if(style.display==="none"||style.visibility==="hidden"||style.opacity==="0") continue;
      // Find all download anchors/buttons in this modal
      var btns=modal.querySelectorAll("a,button");
      for(var k=0;k<btns.length;k++){
        var btn=btns[k];
        var txt=(btn.innerText||btn.textContent||btn.value||"").toLowerCase();
        var href=(btn.href||"").toLowerCase();
        if(txt.includes("download")||href.includes("download")||href.includes(".zip")||
           href.includes(".7z")||href.includes(".rar")){
          clicked.push((btn.innerText||btn.href||"btn").trim().substring(0,60));
          btn.click();
        }
      }
    }
  }
  return JSON.stringify(clicked.length>0?clicked:["no-modal-buttons-found"]);
})()
"""


def _install_archives(archives: list[Path], community: Path) -> dict:
    """Extract each archive and copy mod roots into *community*. Returns summary dict."""
    all_installed: list[str] = []
    errors: list[str] = []
    for archive in archives:
        if archive.suffix.lower() not in SUPPORTED_ARCHIVES:
            errors.append(f"{archive.name}: unsupported format {archive.suffix}")
            continue
        with tempfile.TemporaryDirectory(prefix="msfs_mod_") as tmp:
            tmp_path = Path(tmp)
            extract_dir = tmp_path / "extracted"
            extract_dir.mkdir()
            try:
                _extract(archive, extract_dir)
            except Exception as exc:
                errors.append(f"{archive.name}: extraction failed – {exc}")
                continue
            mod_roots = _find_mod_roots(extract_dir)
            if not mod_roots:
                errors.append(f"{archive.name}: no mod folders found")
                continue
            for root in mod_roots:
                dest = community / root.name
                if dest.exists():
                    shutil.rmtree(dest)
                shutil.copytree(root, dest)
                all_installed.append(root.name)
    return {"installed_mods": all_installed, "errors": errors}


# ---------------------------------------------------------------------------
# Intelligent page analysis – figures out HOW to download before acting
# ---------------------------------------------------------------------------

# JS that reads the full page state and returns a structured analysis JSON.
_JS_ANALYSE_PAGE = r"""
(function(){
  var result = {
    url: window.location.href,
    title: document.title,
    strategy: "unknown",
    directLinks: [],
    downloadButtons: [],
    modalTriggers: [],
    captcha: false,
    loginRequired: false,
    countdown: null,
    notes: []
  };

  // 1. Detect login wall
  var bodyText = (document.body && document.body.innerText || "").toLowerCase();
  if(bodyText.includes("sign in") || bodyText.includes("log in") || bodyText.includes("anmelden"))
    result.loginRequired = true;

  // 2. Detect captcha / human check
  if(bodyText.includes("captcha") || bodyText.includes("human") || document.querySelector("iframe[src*='recaptcha']"))
    result.captcha = true;

  // 3. Detect countdown timer (some sites delay the download button)
  var timerEl = document.querySelector("[id*='timer'],[class*='timer'],[id*='countdown'],[class*='countdown']");
  if(timerEl) result.countdown = (timerEl.innerText||"").trim().substring(0,20);

  // 4. Find all <a> tags with direct archive links
  var anchors = document.querySelectorAll("a[href]");
  for(var i=0;i<anchors.length;i++){
    var href = anchors[i].href || "";
    if(/\.(zip|7z|rar|exe|msi)(\?|$)/i.test(href)){
      result.directLinks.push({label:(anchors[i].innerText||"").trim().substring(0,60), href:href});
    }
  }

  // 5. Find all clickable download elements (buttons + anchors with download text)
  var clickables = document.querySelectorAll("a,button,[role='button']");
  for(var j=0;j<clickables.length;j++){
    var el = clickables[j];
    var txt = (el.innerText||el.textContent||el.value||"").toLowerCase().trim();
    var elHref = (el.href||"").toLowerCase();
    var isDownload = txt.includes("download") || txt.includes("herunterladen") ||
                     elHref.includes("download") || el.hasAttribute("download");
    if(!isDownload) continue;
    var style = window.getComputedStyle(el);
    var visible = style.display!=="none" && style.visibility!=="hidden" && style.opacity!=="0";
    var info = {
      tag: el.tagName,
      text: (el.innerText||"").trim().substring(0,60),
      href: el.href||"",
      id: el.id||"",
      cls: (el.className||"").substring(0,60),
      visible: visible,
      isModalTrigger: !!(el.getAttribute("data-bs-toggle")||el.getAttribute("data-toggle")||
                         el.getAttribute("data-target")||el.getAttribute("data-modal"))
    };
    if(info.isModalTrigger) result.modalTriggers.push(info);
    else result.downloadButtons.push(info);
  }

  // 6. Decide best strategy
  if(result.captcha){
    result.strategy = "captcha_blocked";
    result.notes.push("CAPTCHA detected – manual interaction needed");
  } else if(result.directLinks.length > 0){
    result.strategy = "direct_link";
    result.notes.push("Direct archive link(s) found – no click needed");
  } else if(result.countdown){
    result.strategy = "countdown_then_click";
    result.notes.push("Countdown timer found: " + result.countdown);
  } else if(result.modalTriggers.length > 0){
    result.strategy = "click_opens_modal";
    result.notes.push("Download button opens a modal with multiple files");
  } else if(result.downloadButtons.length > 0){
    result.strategy = "direct_click";
    result.notes.push("Download button found – single click");
  } else {
    result.strategy = "unknown";
    result.notes.push("No download element found on this page state");
  }

  return JSON.stringify(result);
})()
"""

# JS that waits for a countdown to reach zero then returns.
_JS_WAIT_COUNTDOWN = r"""
(function(){
  var max = 30, waited = 0;
  var check = setInterval(function(){
    waited++;
    var timers = document.querySelectorAll("[id*='timer'],[class*='timer'],[id*='countdown'],[class*='countdown']");
    var active = false;
    for(var i=0;i<timers.length;i++){
      var t = (timers[i].innerText||"").trim();
      if(t && /\d/.test(t) && parseInt(t) > 0){ active = true; break; }
    }
    if(!active || waited >= max){ clearInterval(check); }
  }, 1000);
  return "countdown-wait-started";
})()
"""


def _analyse_page(url: str) -> dict:
    """Open *url* in Chrome via CDP, analyse the page and return a strategy dict.
    The returned dict includes a '_ws_url' key so callers can reuse the same tab."""
    try:
        ws_url = _chrome_open_url(url)  # opens tab, waits 3 s
    except Exception as exc:
        return {"error": f"Cannot open Chrome via CDP: {exc}", "strategy": "error"}

    if not ws_url:
        return {"error": "CDP tab could not be opened", "strategy": "error"}

    for attempt in range(3):
        try:
            raw = _cdp_js(ws_url, _JS_ANALYSE_PAGE)
            if raw.startswith("{"):
                result = _json.loads(raw)
                result["_ws_url"] = ws_url   # carry tab reference forward
                return result
        except Exception:
            pass
        _time.sleep(1)

    return {"strategy": "unknown", "notes": ["Page analysis JS failed"], "_ws_url": ws_url}


def _execute_strategy(analysis: dict, dl_dir: Path, modal_timeout: int, timeout: int) -> dict:
    """Given a page analysis dict (with _ws_url), execute the download strategy.
    All JS is sent to the same CDP tab. Returns {"collected": [Path, ...], "log": [...]}"""
    strategy = analysis.get("strategy", "unknown")
    ws = analysis.get("_ws_url")            # CDP WebSocket URL of the open tab
    log: list[str] = [f"Strategy chosen: {strategy}", f"CDP tab: {ws or 'none'}"]
    log += analysis.get("notes", [])

    def js(code: str) -> str:
        return _cdp_js(ws, code) if ws else "error:no-ws"

    def snapshot() -> set[Path]:
        return {
            p for p in dl_dir.iterdir()
            if p.is_file() and not p.name.startswith(".")
            and p.suffix.lower() not in {".crdownload", ".part", ".tmp"}
        }

    before = snapshot()

    # ── Strategy: direct archive link ────────────────────────────────────────
    if strategy == "direct_link":
        for link in analysis.get("directLinks", []):
            href = link.get("href", "")
            if href:
                log.append(f"CDP navigate to direct link: {href[:80]}")
                try:
                    if ws:
                        _cdp_navigate(ws, href)
                    else:
                        _chrome_open_url(href)
                except Exception as e:
                    log.append(f"  error: {e}")

    # ── Strategy: countdown then click ───────────────────────────────────────
    elif strategy == "countdown_then_click":
        log.append("Waiting for countdown via CDP…")
        js(_JS_WAIT_COUNTDOWN)
        _time.sleep(32)
        r = js(_JS_CLICK_MAIN_DOWNLOAD)
        log.append(f"Post-countdown click: {r}")

    # ── Strategy: click opens modal ───────────────────────────────────────────
    elif strategy == "click_opens_modal":
        r = js(_JS_CLICK_MAIN_DOWNLOAD)
        log.append(f"Modal trigger clicked via CDP: {r}")
        # Wait for modal, then click all items inside
        deadline = _time.time() + modal_timeout
        modal_found = False
        while _time.time() < deadline:
            try:
                raw = js(_JS_DETECT_MODAL)
                info = _json.loads(raw) if raw.startswith("{") else {}
                if info.get("found"):
                    modal_found = True
                    break
            except Exception:
                pass
            _time.sleep(0.8)
        if modal_found:
            log.append("Modal detected via CDP – clicking all download buttons inside")
            raw_clicks = js(_JS_CLICK_ALL_MODAL_DOWNLOADS)
            clicks = _json.loads(raw_clicks) if raw_clicks.startswith("[") else []
            log.append(f"Modal clicks: {clicks}")
        else:
            log.append("Modal did not appear – trying direct click via CDP as fallback")
            r = js(_JS_CLICK_MAIN_DOWNLOAD)
            log.append(f"Fallback click: {r}")

    # ── Strategy: direct click ────────────────────────────────────────────────
    elif strategy == "direct_click":
        r = js(_JS_CLICK_MAIN_DOWNLOAD)
        log.append(f"Download button clicked via CDP: {r}")
        # Also check if a modal pops up after click
        _time.sleep(1.5)
        try:
            raw = js(_JS_DETECT_MODAL)
            info = _json.loads(raw) if raw.startswith("{") else {}
            if info.get("found"):
                log.append("Unexpected modal appeared – clicking all buttons via CDP")
                raw_clicks = js(_JS_CLICK_ALL_MODAL_DOWNLOADS)
                log.append(f"Modal clicks: {raw_clicks}")
        except Exception:
            pass

    # ── Strategy: captcha / unknown – best-effort ─────────────────────────────
    else:
        log.append("Unknown strategy – attempting blind click via CDP")
        r = js(_JS_CLICK_MAIN_DOWNLOAD)
        log.append(f"Blind click: {r}")

    # ── Wait for files ────────────────────────────────────────────────────────
    collected: list[Path] = []
    deadline_dl = _time.time() + timeout
    while _time.time() < deadline_dl:
        current = snapshot()
        new = current - before - set(collected)
        if new:
            collected.extend(new)
            # Give a bit more time for additional files (multi-file modal)
            deadline_dl = max(deadline_dl, _time.time() + 8)
        _time.sleep(1.5)

    return {"collected": collected, "log": log}


@mcp.tool()
def download_mod_via_browser(
    url: str,
    downloads_dir: str = "",
    community_path: str = "",
    auto_install: bool = True,
    timeout: int = 120,
    modal_timeout: int = 8,
) -> dict:
    """Smart browser-assisted mod downloader for flightsim.to (and similar sites).

    Before clicking anything the tool ANALYSES the page and chooses the best
    download strategy automatically:

      • direct_link       – Archive link found in HTML → navigate directly, no click
      • direct_click      – Single visible Download button → click it
      • click_opens_modal – Button opens a modal with multiple files → click ALL items
      • countdown_then_click – Timer delays the button → wait, then click
      • captcha_blocked   – CAPTCHA detected → reports back, lets user know
      • unknown           – Nothing found → blind click as last resort

    After downloading, every archive (.zip / .7z / .rar) is extracted and
    copied into the MSFS Community folder (when auto_install=True).

    Works on Windows via Chrome DevTools Protocol (CDP). Chrome is started automatically.

    Args:
        url:            flightsim.to addon/file page URL.
        downloads_dir:  Chrome download directory (default: ~/Downloads).
        community_path: MSFS 2024 Community folder – leave empty for auto-detect.
        auto_install:   Extract & install into Community when True (default).
        timeout:        Seconds to wait for all files (default 120).
        modal_timeout:  Seconds to wait for a modal to appear (default 8).
    """
    dl_dir = Path(downloads_dir).expanduser() if downloads_dir else Path.home() / "Downloads"
    if not dl_dir.exists():
        return {"error": f"Downloads directory not found: {dl_dir}"}

    # ── 1. Analyse the page ───────────────────────────────────────────────────
    analysis = _analyse_page(url)

    if analysis.get("loginRequired"):
        return {
            "status": "login_required",
            "url": url,
            "strategy": "login_required",
            "message": (
                "Die Seite erfordert einen Login. Chrome wurde geöffnet — "
                "bitte logge dich im Browser ein. Danach ruf dieses Tool "
                "erneut mit derselben URL auf."
            ),
            "page_analysis": analysis,
        }

    if analysis.get("strategy") == "captcha_blocked":
        return {
            "status": "blocked",
            "strategy": "captcha_blocked",
            "message": "CAPTCHA erkannt. Bitte im Browser lösen, dann erneut aufrufen.",
            "page_analysis": analysis,
        }
    if "error" in analysis:
        return {"status": "error", "message": analysis["error"]}

    # ── 2. Execute strategy + collect files ───────────────────────────────────
    exec_result = _execute_strategy(analysis, dl_dir, modal_timeout, timeout)
    collected: list[Path] = exec_result["collected"]
    log: list[str] = exec_result["log"]

    if not collected:
        return {
            "status": "timeout",
            "strategy": analysis.get("strategy"),
            "log": log,
            "page_analysis": {
                "directLinks": analysis.get("directLinks", []),
                "downloadButtons": len(analysis.get("downloadButtons", [])),
                "modalTriggers": len(analysis.get("modalTriggers", [])),
                "notes": analysis.get("notes", []),
            },
            "message": f"No new files appeared in {dl_dir} within {timeout}s.",
        }

    result: dict = {
        "status": "downloaded",
        "strategy_used": analysis.get("strategy"),
        "files": [str(p) for p in collected],
        "filenames": [p.name for p in collected],
        "log": log,
    }

    # ── 3. Install ────────────────────────────────────────────────────────────
    if auto_install:
        community = _detect_community_folder(community_path)
        if community is None:
            result["install_status"] = "skipped – Community folder not found. Pass community_path."
            return result

        archives = [p for p in collected if p.suffix.lower() in SUPPORTED_ARCHIVES]
        if not archives:
            result["install_status"] = (
                f"skipped – none of the downloaded files are supported archives "
                f"({', '.join(SUPPORTED_ARCHIVES)})."
            )
            return result

        install_result = _install_archives(archives, community)
        result["install_status"] = "ok" if not install_result["errors"] else "partial"
        result["community_folder"] = str(community)
        result["installed_mods"] = install_result["installed_mods"]
        result["install_errors"] = install_result["errors"]

    return result


@mcp.tool()
def list_modal_downloads(url: str) -> dict:
    """Open a flightsim.to page, click the Download button, wait for a modal,
    and return a list of all downloadable items found in the modal WITHOUT
    downloading anything yet.  Useful to preview what a multi-file mod contains.

    Args:
        url: flightsim.to addon/file page URL.
    """
    try:
        ws = _chrome_open_url(url)   # CDP: open tab, return ws URL
    except Exception as exc:
        return {"error": f"Could not open Chrome via CDP: {exc}"}
    if not ws:
        return {"error": "CDP tab could not be opened"}

    _cdp_js(ws, _JS_CLICK_MAIN_DOWNLOAD)

    for _ in range(10):
        _time.sleep(1)
        try:
            raw = _cdp_js(ws, _JS_DETECT_MODAL)
            info = _json.loads(raw) if raw.startswith("{") else {"found": False}
        except Exception:
            info = {"found": False}
        if info.get("found"):
            return {
                "modal_found": True,
                "download_buttons": info.get("downloadButtons", []),
                "count": len(info.get("downloadButtons", [])),
                "cdp_tab": ws,
            }

    return {"modal_found": False, "message": "No modal appeared within 10 seconds."}


# ---------------------------------------------------------------------------
# Mod search helper – queries flightsim.to search API + page scrape
# ---------------------------------------------------------------------------

def _search_flightsim(query: str, max_results: int = 8) -> list[dict]:
    """Search flightsim.to for *query* and return a list of mod dicts.
    Each dict has keys: id, title, author, category, url, thumbnail.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/124.0 Safari/537.36",
        "Accept": "application/json, text/html, */*",
        "Referer": "https://flightsim.to/",
    }
    results: list[dict] = []

    with httpx.Client(follow_redirects=True, timeout=20, headers=headers) as client:

        # ── Try 1: JSON API ───────────────────────────────────────────────────
        try:
            resp = client.get(
                "https://flightsim.to/api/v1/search",
                params={"q": query, "per_page": max_results},
            )
            if resp.status_code == 200:
                data = resp.json()
                items = data if isinstance(data, list) else data.get("data", [])
                for item in items[:max_results]:
                    addon_id = item.get("id") or item.get("addon_id", "")
                    slug     = item.get("slug") or item.get("name", str(addon_id))
                    results.append({
                        "id":        str(addon_id),
                        "title":     item.get("title") or item.get("name", "Unknown"),
                        "author":    (item.get("user") or {}).get("username", ""),
                        "category":  item.get("category", ""),
                        "url":       f"https://flightsim.to/file/{addon_id}/{slug}",
                        "thumbnail": item.get("thumbnail") or item.get("image", ""),
                    })
                if results:
                    return results
        except Exception:
            pass

        # ── Try 2: Scrape the search results page ─────────────────────────────
        try:
            resp = client.get(
                "https://flightsim.to/search",
                params={"q": query},
            )
            if resp.status_code == 200:
                html = resp.text
                # Extract addon cards – pattern: /file/{id}/{slug}
                pattern = r'href="https://flightsim\.to/file/(\d+)/([^"]+)"[^>]*>.*?<[^>]+class="[^"]*title[^"]*"[^>]*>([^<]+)<'
                for m in re.finditer(pattern, html, re.S):
                    addon_id, slug, title = m.group(1), m.group(2), m.group(3).strip()
                    if not any(r["id"] == addon_id for r in results):
                        results.append({
                            "id":        addon_id,
                            "title":     title,
                            "author":    "",
                            "category":  "",
                            "url":       f"https://flightsim.to/file/{addon_id}/{slug}",
                            "thumbnail": "",
                        })
                    if len(results) >= max_results:
                        break

                # Fallback: simpler href scan
                if not results:
                    seen: set[str] = set()
                    for m in re.finditer(
                        r'href="(https://flightsim\.to/file/(\d+)/([^"?#]+))"', html
                    ):
                        href, addon_id, slug = m.group(1), m.group(2), m.group(3)
                        if addon_id in seen or not slug or slug.endswith("/download"):
                            continue
                        seen.add(addon_id)
                        results.append({
                            "id":        addon_id,
                            "title":     slug.replace("-", " ").title(),
                            "author":    "",
                            "category":  "",
                            "url":       href,
                            "thumbnail": "",
                        })
                        if len(results) >= max_results:
                            break
        except Exception:
            pass

    # ── Try 3: CDP – open Chrome, let JS render, read live DOM ───────────────
    if not results:
        results = _search_flightsim_via_cdp(query, max_results)

    return results


# JS run inside Chrome after the search page has rendered –
# extracts every addon card that is actually visible in the DOM.
_JS_EXTRACT_SEARCH_RESULTS = r"""
(function(){
  var cards = [];
  // flightsim.to renders cards as <a href="/file/{id}/{slug}"> with a title inside
  var selectors = [
    'a[href*="/file/"]',
    '[class*="card"] a[href*="/file/"]',
    '[class*="addon"] a[href*="/file/"]',
    '[class*="result"] a[href*="/file/"]',
  ];
  var seen = {};
  selectors.forEach(function(sel){
    document.querySelectorAll(sel).forEach(function(el){
      var href = el.href || "";
      var m = href.match(/\/file\/(\d+)\/([^/?#]+)/);
      if (!m || seen[m[1]]) return;
      seen[m[1]] = true;

      // Title: prefer dedicated title element, fall back to link text
      var titleEl = el.querySelector('[class*="title"],[class*="name"],h2,h3,h4,strong');
      var title = (titleEl ? titleEl.innerText : el.innerText || "").trim().substring(0,100);
      if (!title) title = m[2].replace(/-/g," ");

      // Author
      var authorEl = el.querySelector('[class*="author"],[class*="user"],[class*="username"]');
      var author = authorEl ? authorEl.innerText.trim() : "";

      // Category / tag
      var catEl = el.querySelector('[class*="category"],[class*="tag"],[class*="type"]');
      var category = catEl ? catEl.innerText.trim() : "";

      // Downloads / rating  (extra context)
      var dlEl = el.querySelector('[class*="download"],[class*="count"]');
      var downloads = dlEl ? dlEl.innerText.trim() : "";

      cards.push({
        id: m[1],
        title: title,
        author: author,
        category: category,
        downloads: downloads,
        url: "https://flightsim.to/file/" + m[1] + "/" + m[2]
      });
    });
  });
  return JSON.stringify(cards.slice(0,12));
})()
"""


def _search_flightsim_via_cdp(query: str, max_results: int = 8) -> list[dict]:
    """Open flightsim.to/search in Chrome via CDP, wait for JS to render,
    then read the fully populated DOM to extract search results."""
    try:
        search_url = f"https://flightsim.to/search?q={query.replace(' ', '+')}"
        ws = _chrome_open_url(search_url)   # opens tab, waits 3 s
        if not ws:
            return []

        # Give JS-rendered content a bit more time to appear
        _time.sleep(2)

        raw = _cdp_js(ws, _JS_EXTRACT_SEARCH_RESULTS, timeout=15)
        if not raw.startswith("["):
            return []

        items = _json.loads(raw)
        return [
            {
                "id":        item["id"],
                "title":     item["title"],
                "author":    item.get("author", ""),
                "category":  item.get("category", ""),
                "downloads": item.get("downloads", ""),
                "url":       item["url"],
                "thumbnail": "",
                "source":    "cdp",     # so caller knows this came from live DOM
            }
            for item in items[:max_results]
        ]
    except Exception:
        return []


@mcp.tool()
def find_and_install_mod(
    mod_name: str,
    choice_index: int = -1,
    community_path: str = "",
    downloads_dir: str = "",
    auto_install: bool = True,
) -> dict:
    """Find and install an MSFS mod by name from flightsim.to.

    Workflow:
      1. Search flightsim.to for *mod_name*.
      2. No results → tell the user, done.
      3. Exactly 1 result → open the page and download + install automatically.
      4. Multiple results → return the list so the user can pick one.
         Then call again with the same *mod_name* and *choice_index* (0-based)
         to proceed with the chosen entry.

    Args:
        mod_name:       Name or keyword to search for (e.g. "EDDF Frankfurt").
        choice_index:   0-based index from a previous multi-result call.
                        Pass -1 (default) to trigger a new search.
        community_path: MSFS 2024 Community folder – leave empty for auto-detect.
        downloads_dir:  Chrome download directory – leave empty for ~/Downloads.
        auto_install:   Extract & install after download (default True).
    """

    # ── A. User is confirming a previous multi-result search ─────────────────
    if choice_index >= 0:
        # Re-run the search to get the same list
        hits = _search_flightsim(mod_name)
        if not hits:
            return {"status": "no_results",
                    "message": f"No results found for '{mod_name}'."}
        if choice_index >= len(hits):
            return {
                "status": "invalid_choice",
                "message": f"Index {choice_index} is out of range (0–{len(hits)-1}).",
                "results": hits,
            }
        chosen = hits[choice_index]
        # Fall through to download
        return _trigger_download(chosen["url"], community_path, downloads_dir, auto_install)

    # ── B. Fresh search ───────────────────────────────────────────────────────
    hits = _search_flightsim(mod_name)

    if not hits:
        return {
            "status": "no_results",
            "message": (
                f"Keine Ergebnisse für '{mod_name}' auf flightsim.to gefunden. "
                "Bitte prüfe den Mod-Namen oder suche manuell auf https://flightsim.to"
            ),
        }

    if len(hits) == 1:
        # Single hit → download directly, no confirmation needed
        return _trigger_download(hits[0]["url"], community_path, downloads_dir, auto_install)

    # Multiple hits → return list and ask user to choose
    choices = [
        {"index": i, "title": h["title"], "author": h["author"],
         "category": h["category"], "url": h["url"]}
        for i, h in enumerate(hits)
    ]
    return {
        "status": "multiple_results",
        "message": (
            f"{len(hits)} Ergebnisse für '{mod_name}' gefunden. "
            "Bitte ruf das Tool erneut mit choice_index=N auf (0-basiert)."
        ),
        "results": choices,
    }


def _trigger_download(url: str, community_path: str,
                      downloads_dir: str, auto_install: bool) -> dict:
    """Internal helper – runs the full smart download flow for a known URL."""
    dl_dir = (Path(downloads_dir).expanduser()
              if downloads_dir else Path.home() / "Downloads")

    if not dl_dir.exists():
        return {"error": f"Downloads directory not found: {dl_dir}"}

    analysis = _analyse_page(url)

    # Handle login-required pages
    if analysis.get("loginRequired"):
        return {
            "status": "login_required",
            "url": url,
            "message": (
                "Die Seite erfordert einen Login. Chrome wurde geöffnet — "
                "bitte logge dich im Browser ein. Danach ruf dieses Tool "
                "erneut mit derselben URL auf."
            ),
        }

    if analysis.get("strategy") == "captcha_blocked":
        return {
            "status": "blocked",
            "url": url,
            "message": "CAPTCHA erkannt. Bitte im Browser lösen, dann erneut aufrufen.",
        }
    if "error" in analysis:
        return {"status": "error", "url": url, "message": analysis["error"]}

    exec_result = _execute_strategy(analysis, dl_dir, modal_timeout=8, timeout=120)
    collected: list[Path] = exec_result["collected"]
    log: list[str]        = exec_result["log"]

    if not collected:
        return {
            "status": "timeout",
            "url": url,
            "strategy": analysis.get("strategy"),
            "log": log,
            "message": f"Keine Datei in {dl_dir} innerhalb von 120s erschienen.",
        }

    result: dict = {
        "status": "downloaded",
        "url": url,
        "strategy_used": analysis.get("strategy"),
        "filenames": [p.name for p in collected],
        "log": log,
    }

    if auto_install:
        community = _detect_community_folder(community_path)
        if community is None:
            result["install_status"] = "skipped – Community-Ordner nicht gefunden."
            return result
        archives = [p for p in collected if p.suffix.lower() in SUPPORTED_ARCHIVES]
        if not archives:
            result["install_status"] = "skipped – keine unterstützten Archive gefunden."
            return result
        ir = _install_archives(archives, community)
        result["install_status"]  = "ok" if not ir["errors"] else "partial"
        result["community_folder"] = str(community)
        result["installed_mods"]   = ir["installed_mods"]
        result["install_errors"]   = ir["errors"]

    return result


# ===========================================================================
# ReShade Integration (reshade.me)
# ===========================================================================

import configparser as _configparser

# ── ReShade config search paths ───────────────────────────────────────────

# Common MSFS install locations (ReShade.ini lives next to the game .exe)
_RESHADE_GAME_DIRS: list[Path] = [
    # MSFS 2024 (Steam)
    Path(r"C:\Program Files (x86)\Steam\steamapps\common\Microsoft Flight Simulator 2024"),
    Path(r"D:\Steam\steamapps\common\Microsoft Flight Simulator 2024"),
    Path(r"D:\SteamLibrary\steamapps\common\Microsoft Flight Simulator 2024"),
    Path(r"E:\SteamLibrary\steamapps\common\Microsoft Flight Simulator 2024"),
    # MSFS 2024 (MS Store)
    Path(os.environ.get("LOCALAPPDATA", ""))
    / "Packages" / "Microsoft.Limitless_8wekyb3d8bbwe" / "LocalCache" / "Local",
    # Generic game dirs
    Path(r"C:\Program Files (x86)\Steam\steamapps\common"),
    Path(r"D:\Steam\steamapps\common"),
    Path(r"D:\Games"),
    Path(r"E:\Games"),
]

_RESHADE_INI_NAMES = [
    "ReShade.ini", "reshade.ini", "Reshade.ini",
    "dxgi.ini", "d3d11.ini", "d3d12.ini", "opengl32.ini",
]

_RESHADE_PRESET_EXTENSIONS = {".ini", ".txt"}

# ── ReShade effect definitions ────────────────────────────────────────────

RESHADE_EFFECTS: dict[str, dict] = {
    # LUT / Color
    "LiftGammaGain": {
        "label": "Lift Gamma Gain",
        "category": "Color",
        "keys": {
            "RGB_Lift": {"label": "Lift (Schatten)", "default": "1.0,1.0,1.0"},
            "RGB_Gamma": {"label": "Gamma", "default": "1.0,1.0,1.0"},
            "RGB_Gain": {"label": "Gain (Lichter)", "default": "1.0,1.0,1.0"},
        },
    },
    "Tonemap": {
        "label": "Tonemap",
        "category": "Color",
        "keys": {
            "Gamma": {"label": "Gamma", "default": "1.0", "range": "0.0-2.0"},
            "Exposure": {"label": "Belichtung", "default": "0.0", "range": "-1.0-1.0"},
            "Saturation": {"label": "Sättigung", "default": "0.0", "range": "-1.0-1.0"},
            "Bleach": {"label": "Bleach", "default": "0.0", "range": "0.0-1.0"},
            "Defog": {"label": "Entnebelung", "default": "0.0", "range": "0.0-1.0"},
        },
    },
    "Vibrance": {
        "label": "Vibrance",
        "category": "Color",
        "keys": {
            "Vibrance": {"label": "Vibrance", "default": "0.15", "range": "-1.0-1.0"},
        },
    },
    "Colorfulness": {
        "label": "Colorfulness",
        "category": "Color",
        "keys": {
            "colorfullness": {"label": "Farbintensität", "default": "0.4", "range": "0.0-2.0"},
            "lim_luma": {"label": "Luminanz-Limit", "default": "0.7"},
        },
    },
    "Levels": {
        "label": "Levels",
        "category": "Color",
        "keys": {
            "BlackPoint": {"label": "Schwarzpunkt", "default": "16", "range": "0-255"},
            "WhitePoint": {"label": "Weißpunkt", "default": "235", "range": "0-255"},
        },
    },
    # Sharpening
    "CAS": {
        "label": "AMD CAS (Contrast Adaptive Sharpening)",
        "category": "Sharpening",
        "keys": {
            "Sharpness": {"label": "Schärfe", "default": "0.4", "range": "0.0-1.0"},
        },
    },
    "AdaptiveSharpen": {
        "label": "Adaptive Sharpen",
        "category": "Sharpening",
        "keys": {
            "curve_height": {"label": "Schärfe-Intensität", "default": "0.3", "range": "0.0-2.0"},
        },
    },
    "LumaSharpen": {
        "label": "LumaSharpen",
        "category": "Sharpening",
        "keys": {
            "sharp_strength": {"label": "Schärfe", "default": "0.65", "range": "0.1-3.0"},
            "sharp_clamp": {"label": "Clamp", "default": "0.035"},
        },
    },
    # Clarity / Detail
    "Clarity": {
        "label": "Clarity",
        "category": "Detail",
        "keys": {
            "ClarityRadius": {"label": "Radius", "default": "3"},
            "ClarityBlendMode": {"label": "Blend-Modus", "default": "2"},
            "ClarityOffset": {"label": "Offset", "default": "2.0"},
            "ClarityStrength": {"label": "Stärke", "default": "0.4", "range": "0.0-1.0"},
        },
    },
    # Bloom / Glow
    "MagicBloom": {
        "label": "Magic Bloom",
        "category": "Bloom",
        "keys": {
            "f2Sigma": {"label": "Bloom-Breite", "default": "0.5"},
            "f2Intensity": {"label": "Intensität", "default": "1.0", "range": "0.0-10.0"},
        },
    },
    # Anti-Aliasing
    "SMAA": {
        "label": "SMAA Anti-Aliasing",
        "category": "AA",
        "keys": {
            "EdgeDetectionType": {"label": "Kantenerkennung", "default": "1",
                                  "values": {"0": "Luminance", "1": "Color", "2": "Depth"}},
            "EdgeDetectionThreshold": {"label": "Schwellwert", "default": "0.1"},
        },
    },
    "FXAA": {
        "label": "FXAA Anti-Aliasing",
        "category": "AA",
        "keys": {
            "Subpix": {"label": "Sub-Pixel AA", "default": "0.25", "range": "0.0-1.0"},
            "EdgeThreshold": {"label": "Schwellwert", "default": "0.125"},
        },
    },
    # Depth effects (Vorsicht in VR — Performance-intensiv!)
    "MXAO": {
        "label": "MXAO (Ambient Occlusion)",
        "category": "Depth",
        "keys": {
            "MXAO_SAMPLE_COUNT": {"label": "Sample-Anzahl", "default": "24"},
            "MXAO_SAMPLE_RADIUS": {"label": "Radius", "default": "2.5"},
            "MXAO_AMOUNT": {"label": "Intensität", "default": "2.0"},
        },
    },
}

# ── ReShade VR presets ────────────────────────────────────────────────────

RESHADE_VR_PRESETS: dict[str, dict] = {
    "VR_Performance": {
        "description": "Minimale Last — nur Schärfe",
        "effects": {
            "CAS": {"enabled": True, "Sharpness": "0.5"},
            "Vibrance": {"enabled": True, "Vibrance": "0.15"},
        },
        "disable": ["MXAO", "MagicBloom", "SMAA", "FXAA", "Clarity", "AdaptiveSharpen"],
    },
    "VR_Balanced": {
        "description": "Gute Optik bei moderater Last",
        "effects": {
            "CAS": {"enabled": True, "Sharpness": "0.4"},
            "Clarity": {"enabled": True, "ClarityStrength": "0.4"},
            "Vibrance": {"enabled": True, "Vibrance": "0.2"},
            "Tonemap": {"enabled": True, "Gamma": "1.0", "Exposure": "0.0",
                        "Saturation": "0.05"},
        },
        "disable": ["MXAO", "MagicBloom", "FXAA"],
    },
    "VR_Quality": {
        "description": "Maximale Bildqualität",
        "effects": {
            "AdaptiveSharpen": {"enabled": True, "curve_height": "0.4"},
            "Clarity": {"enabled": True, "ClarityStrength": "0.5"},
            "Vibrance": {"enabled": True, "Vibrance": "0.25"},
            "Tonemap": {"enabled": True, "Gamma": "1.0", "Exposure": "0.0",
                        "Saturation": "0.1"},
            "Colorfulness": {"enabled": True, "colorfullness": "0.5"},
            "Levels": {"enabled": True, "BlackPoint": "16", "WhitePoint": "235"},
            "SMAA": {"enabled": True},
        },
        "disable": ["MXAO", "MagicBloom"],
    },
    "VR_MSFS_Optimized": {
        "description": "Speziell für MSFS 2024 in VR",
        "effects": {
            "CAS": {"enabled": True, "Sharpness": "0.6"},
            "Clarity": {"enabled": True, "ClarityStrength": "0.35",
                        "ClarityRadius": "3"},
            "Vibrance": {"enabled": True, "Vibrance": "0.2"},
            "Tonemap": {"enabled": True, "Gamma": "1.02", "Exposure": "0.05",
                        "Saturation": "0.05", "Defog": "0.0"},
            "Colorfulness": {"enabled": True, "colorfullness": "0.3"},
        },
        "disable": ["MXAO", "MagicBloom", "SMAA", "FXAA", "AdaptiveSharpen"],
    },
}


# ── ReShade helpers (old _find/_read/_write helpers removed — use _rs_locate_ini / _rs_load_ini / _rs_save_ini) ───


def _find_preset_file(ini_path: Path) -> Path | None:
    """Find the currently active preset file referenced by ReShade.ini."""
    config, _ = _rs_load_ini()
    preset_path = None
    for section in config.sections():
        if config.has_option(section, "PresetPath"):
            preset_path = config.get(section, "PresetPath")
            break
    if not preset_path:
        return None

    # Resolve relative to ReShade.ini directory
    p = Path(preset_path)
    if not p.is_absolute():
        p = ini_path.parent / p
    return p if p.is_file() else None


def _read_preset(preset_path: Path) -> _configparser.ConfigParser:
    """Read a ReShade preset .ini file."""
    config = _configparser.ConfigParser(
        interpolation=None,
        comment_prefixes=(";",),
        inline_comment_prefixes=(";",),
    )
    config.optionxform = str
    config.read(str(preset_path), encoding="utf-8")
    return config


def _list_presets(ini_path: Path) -> list[dict]:
    """Find all preset files near the ReShade installation."""
    presets: list[dict] = []
    base = ini_path.parent

    # Check "reshade-presets", "Presets", "presets" folders
    for folder_name in ["reshade-presets", "Presets", "presets", "ReShade", "reshade"]:
        folder = base / folder_name
        if folder.is_dir():
            for f in folder.iterdir():
                if f.suffix.lower() in _RESHADE_PRESET_EXTENSIONS:
                    presets.append({"name": f.stem, "path": str(f)})

    # Also check files directly next to ReShade.ini
    for f in base.iterdir():
        if (f.suffix.lower() in _RESHADE_PRESET_EXTENSIONS
                and f.name.lower() not in {"reshade.ini", "dxgi.ini", "d3d11.ini",
                                            "d3d12.ini", "opengl32.ini"}
                and "preset" in f.name.lower()):
            presets.append({"name": f.stem, "path": str(f)})

    return presets


# ── ReShade MCP Tools ─────────────────────────────────────────────────────

@mcp.tool()
def analyze_reshade(
    game_dir: str = "",
) -> dict:
    """Analyze current ReShade installation and settings.

    Finds ReShade.ini, reads the active preset, lists all effects and their
    settings, and shows available presets.

    Present the results as a markdown table with columns:
    Effekt | Status | Einstellung | Wert | Beschreibung

    Args:
        game_dir: Game directory containing ReShade.ini (auto-detected if empty).
    """
    ini = _rs_locate_ini()
    if ini is None:
        return {
            "error": (
                "ReShade.ini nicht gefunden. Ist ReShade installiert? "
                "Gib den Spielordner über game_dir an "
                "(z.B. den MSFS-Installationsordner)."
            ),
            "searched": [str(p) for p in _RESHADE_GAME_DIRS[:6]],
        }

    config, ini = _rs_load_ini()

    # Active preset
    preset_path = _find_preset_file(ini)
    preset_data = _read_preset(preset_path) if preset_path else None

    # Gather effect status and settings
    effects_table: list[dict] = []
    if preset_data:
        for section in preset_data.sections():
            # Enabled/disabled status
            enabled = True
            technique_section = section
            # Check technique list or per-effect toggle
            for opt in preset_data.options(section):
                if opt.lower() in ("enabled", "active"):
                    enabled = preset_data.getboolean(section, opt, fallback=True)

            # Check known effect definitions
            defn = RESHADE_EFFECTS.get(section, {})
            label = defn.get("label", section)
            category = defn.get("category", "")

            settings: dict[str, str] = {}
            for opt in preset_data.options(section):
                if opt.lower() not in ("enabled", "active"):
                    settings[opt] = preset_data.get(section, opt)

            effects_table.append({
                "effect": label,
                "category": category,
                "enabled": enabled,
                "settings": settings,
            })

    # Available presets
    available_presets = _list_presets(ini)

    # ReShade.ini sections
    ini_sections = {}
    for section in config.sections():
        ini_sections[section] = dict(config.items(section))

    return {
        "reshade_ini": str(ini),
        "active_preset": str(preset_path) if preset_path else None,
        "effects": effects_table,
        "available_presets": available_presets,
        "ini_sections": ini_sections,
        "game_dir": str(ini.parent),
        "display_instructions": (
            "Show effects as a markdown table with columns: "
            "Effekt | Kategorie | Status | Einstellungen. "
            "Use green checkmarks for enabled, red X for disabled. "
            "Answer in German."
        ),
    }


@mcp.tool()
def set_reshade_effect(
    effect: str,
    enabled: bool | None = None,
    settings: dict | None = None,
    game_dir: str = "",
) -> dict:
    """Enable/disable a ReShade effect or change its parameters.

    Use this when the user says things like "Schärfe erhöhen",
    "Bloom ausschalten", "Vibrance auf 0.3", etc.

    Args:
        effect: Effect name (e.g. "CAS", "Clarity", "Vibrance", "Tonemap",
                "SMAA", "FXAA", "MagicBloom", "MXAO", "AdaptiveSharpen",
                "LumaSharpen", "Colorfulness", "Levels", "LiftGammaGain").
        enabled: True to enable, False to disable, None to keep current.
        settings: Dict of parameter→value (e.g. {"Sharpness": "0.5"}).
        game_dir: Game directory with ReShade.ini (auto-detected if empty).
    """
    # Resolve German aliases
    _EFFECT_ALIASES = {
        "schärfe": "CAS",
        "sharpness": "CAS",
        "sharpen": "CAS",
        "cas": "CAS",
        "adaptive sharpen": "AdaptiveSharpen",
        "klarheit": "Clarity",
        "clarity": "Clarity",
        "farbe": "Vibrance",
        "vibrance": "Vibrance",
        "sättigung": "Vibrance",
        "saturation": "Tonemap",
        "tonemap": "Tonemap",
        "tonemapping": "Tonemap",
        "gamma": "Tonemap",
        "belichtung": "Tonemap",
        "exposure": "Tonemap",
        "bloom": "MagicBloom",
        "glow": "MagicBloom",
        "blüte": "MagicBloom",
        "aa": "SMAA",
        "anti-aliasing": "SMAA",
        "antialiasing": "SMAA",
        "smaa": "SMAA",
        "fxaa": "FXAA",
        "ao": "MXAO",
        "ambient occlusion": "MXAO",
        "mxao": "MXAO",
        "farbintensität": "Colorfulness",
        "colorfulness": "Colorfulness",
        "levels": "Levels",
        "luma": "LumaSharpen",
        "lumasharpen": "LumaSharpen",
        "lift gamma gain": "LiftGammaGain",
        "liftgammagain": "LiftGammaGain",
    }
    resolved = _EFFECT_ALIASES.get(effect.lower().strip(), effect.strip())

    ini = _rs_locate_ini()
    if ini is None:
        return {"error": "ReShade.ini nicht gefunden. Gib game_dir an."}

    preset_path = _find_preset_file(ini)
    if not preset_path:
        return {"error": "Kein aktives ReShade-Preset gefunden."}

    # Backup
    backup = preset_path.parent / f"{preset_path.stem}_backup{preset_path.suffix}"
    if not backup.exists():
        shutil.copy2(preset_path, backup)

    config = _read_preset(preset_path)

    if not config.has_section(resolved):
        config.add_section(resolved)

    applied: dict = {}

    if enabled is not None:
        # ReShade uses Techniques list or per-section keys
        # We update both approaches
        config.set(resolved, "Enabled", "1" if enabled else "0")
        applied["enabled"] = enabled

    if settings:
        for key, val in settings.items():
            config.set(resolved, key, str(val))
            applied[key] = val

    with open(preset_path, "w", encoding="utf-8") as f:
        config.write(f)

    defn = RESHADE_EFFECTS.get(resolved, {})

    # ReShade reads preset files live — but only if it detects a change.
    # Touch the file to force a reload.
    try:
        preset_path.touch()
    except Exception:
        pass

    return {
        "status": "ok",
        "effect": defn.get("label", resolved),
        "category": defn.get("category", ""),
        "applied": applied,
        "preset_file": str(preset_path),
        "hinweis": (
            "Preset-Datei aktualisiert. ReShade übernimmt Änderungen "
            "normalerweise live im laufenden Spiel. "
            "Falls nicht sichtbar: Home-Taste drücken → ReShade-Overlay → "
            "Preset neu laden. Oder MSFS neu starten."
        ),
    }


@mcp.tool()
def apply_reshade_preset(
    preset: Literal[
        "VR_Performance", "VR_Balanced", "VR_Quality", "VR_MSFS_Optimized", "custom"
    ] = "VR_Balanced",
    custom_preset_path: str = "",
    game_dir: str = "",
) -> dict:
    """Apply a VR-optimized ReShade preset.

    Presets:
    - VR_Performance: Nur Schärfe — minimale FPS-Last
    - VR_Balanced: Schärfe + Clarity + Vibrance — gute Optik
    - VR_Quality: Volle Bildverbesserung — maximale Qualität
    - VR_MSFS_Optimized: Speziell für MSFS 2024 in VR
    - custom: Eigene Preset-Datei laden (custom_preset_path nötig)

    Args:
        preset: Which preset to apply.
        custom_preset_path: Path to a custom .ini preset (only for preset="custom").
        game_dir: Game directory with ReShade.ini (auto-detected if empty).
    """
    ini = _rs_locate_ini()
    if ini is None:
        return {"error": "ReShade.ini nicht gefunden. Gib game_dir an."}

    # Custom preset: just switch the PresetPath
    if preset == "custom":
        if not custom_preset_path:
            return {"error": "custom_preset_path muss angegeben werden."}
        cp = Path(custom_preset_path)
        if not cp.is_file():
            return {"error": f"Preset-Datei nicht gefunden: {custom_preset_path}"}
        config, ini = _rs_load_ini()
        for section in config.sections():
            if config.has_option(section, "PresetPath"):
                config.set(section, "PresetPath", str(cp))
                break
        else:
            if not config.has_section("GENERAL"):
                config.add_section("GENERAL")
            config.set("GENERAL", "PresetPath", str(cp))
        _rs_save_ini(config, ini)
        return {
            "status": "ok",
            "preset": cp.stem,
            "preset_file": str(cp),
        }

    # Built-in VR presets
    preset_def = RESHADE_VR_PRESETS.get(preset)
    if not preset_def:
        return {"error": f"Unbekanntes Preset: {preset}"}

    preset_path = _find_preset_file(ini)
    if not preset_path:
        # Create new preset file
        preset_path = ini.parent / "reshade-presets" / f"{preset}.ini"
        preset_path.parent.mkdir(exist_ok=True)

    # Backup
    if preset_path.exists():
        backup = preset_path.parent / f"{preset_path.stem}_before_{preset}{preset_path.suffix}"
        if not backup.exists():
            shutil.copy2(preset_path, backup)

    config = _read_preset(preset_path) if preset_path.exists() else _configparser.ConfigParser(
        interpolation=None)
    config.optionxform = str

    # Disable effects listed in 'disable'
    for eff_name in preset_def.get("disable", []):
        if config.has_section(eff_name):
            config.set(eff_name, "Enabled", "0")

    # Enable and configure effects
    applied: dict = {}
    for eff_name, eff_settings in preset_def.get("effects", {}).items():
        if not config.has_section(eff_name):
            config.add_section(eff_name)
        for key, val in eff_settings.items():
            if key == "enabled":
                config.set(eff_name, "Enabled", "1" if val else "0")
            else:
                config.set(eff_name, key, str(val))
        applied[eff_name] = eff_settings

    with open(preset_path, "w", encoding="utf-8") as f:
        config.write(f)

    # Point ReShade.ini to this preset
    ini_config, ini = _rs_load_ini()
    for section in ini_config.sections():
        if ini_config.has_option(section, "PresetPath"):
            ini_config.set(section, "PresetPath", str(preset_path))
            break
    _rs_save_ini(ini_config, ini)

    # Touch to help ReShade detect the change
    try:
        preset_path.touch()
    except Exception:
        pass

    return {
        "status": "ok",
        "preset": preset,
        "description": preset_def["description"],
        "effects_enabled": list(preset_def.get("effects", {}).keys()),
        "effects_disabled": preset_def.get("disable", []),
        "preset_file": str(preset_path),
        "hinweis": (
            f"Preset '{preset}' angewendet. ReShade lädt Änderungen "
            "normalerweise live. Falls nicht sichtbar: Home-Taste → "
            "ReShade-Overlay → Preset wechseln/neu laden."
        ),
    }


@mcp.tool()
def list_reshade_presets(
    game_dir: str = "",
) -> dict:
    """List all available ReShade presets for the game.

    Shows the active preset and all preset files found near the installation.

    Args:
        game_dir: Game directory with ReShade.ini (auto-detected if empty).
    """
    ini = _rs_locate_ini()
    if ini is None:
        return {"error": "ReShade.ini nicht gefunden. Gib game_dir an."}

    active = _find_preset_file(ini)
    available = _list_presets(ini)

    # Add built-in VR presets
    builtin = [
        {
            "name": name,
            "type": "builtin_vr",
            "description": defn["description"],
        }
        for name, defn in RESHADE_VR_PRESETS.items()
    ]

    return {
        "reshade_ini": str(ini),
        "active_preset": str(active) if active else None,
        "file_presets": available,
        "builtin_vr_presets": builtin,
    }


@mcp.tool()
def uninstall_reshade(
    game_dir: str = "",
    backup: bool = True,
) -> dict:
    """Completely uninstall/remove ReShade from a game (MSFS 2024).

    Use this when the user says:
    - "ReShade deinstallieren" / "ReShade entfernen"
    - "ReShade komplett löschen"
    - "ReShade ausschalten"

    Removes all ReShade files:
    - DLL files (dxgi.dll, d3d11.dll, d3d12.dll, opengl32.dll, ReShade64.dll, etc.)
    - INI files (ReShade.ini, dxgi.ini, etc.)
    - Shader folders (reshade-shaders)
    - Preset files (optional)
    - Log files (ReShade.log)

    Creates a backup before deleting (default).

    Args:
        game_dir: Game directory containing ReShade (auto-detected if empty).
        backup: Create a backup folder before deleting (default True).
    """
    ini = _rs_locate_ini()
    if ini is None:
        return {
            "status": "not_found",
            "message": (
                "ReShade nicht gefunden. Entweder bereits deinstalliert "
                "oder game_dir angeben."
            ),
        }

    base = ini.parent

    # All ReShade-related files and folders to remove
    reshade_dlls = [
        "dxgi.dll", "d3d11.dll", "d3d12.dll", "opengl32.dll",
        "ReShade64.dll", "ReShade32.dll", "ReShade.dll",
    ]
    reshade_inis = [
        "ReShade.ini", "reshade.ini", "Reshade.ini",
        "dxgi.ini", "d3d11.ini", "d3d12.ini", "opengl32.ini",
        "ReShadePreset.ini",
    ]
    reshade_logs = ["ReShade.log", "reshade.log", "dxgi.log"]
    reshade_dirs = [
        "reshade-shaders", "reshade-presets", "ReShade",
        "reshade", "Presets", "Shaders",
    ]

    # Collect all files/dirs that exist
    to_delete_files: list[Path] = []
    to_delete_dirs: list[Path] = []

    for name in reshade_dlls + reshade_inis + reshade_logs:
        p = base / name
        if p.exists():
            to_delete_files.append(p)

    for name in reshade_dirs:
        p = base / name
        if p.is_dir():
            to_delete_dirs.append(p)

    # Also find preset .ini files (not game configs)
    for f in base.iterdir():
        if (f.suffix.lower() == ".ini"
                and f.name.lower() not in {"usercfg.opt", "flight.ini", "config.ini"}
                and any(kw in f.name.lower() for kw in ["preset", "reshade", "shader"])):
            to_delete_files.append(f)

    if not to_delete_files and not to_delete_dirs:
        return {
            "status": "clean",
            "message": "Keine ReShade-Dateien gefunden. Bereits sauber.",
            "game_dir": str(base),
        }

    # Backup
    backup_path = None
    if backup:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = base / f"reshade_backup_{ts}"
        backup_dir.mkdir(exist_ok=True)
        for f in to_delete_files:
            try:
                shutil.copy2(f, backup_dir / f.name)
            except Exception:
                pass
        for d in to_delete_dirs:
            try:
                shutil.copytree(d, backup_dir / d.name, dirs_exist_ok=True)
            except Exception:
                pass
        backup_path = str(backup_dir)

    # Delete
    deleted_files: list[str] = []
    deleted_dirs: list[str] = []
    errors: list[str] = []

    for f in to_delete_files:
        try:
            f.unlink()
            deleted_files.append(f.name)
        except Exception as exc:
            errors.append(f"{f.name}: {exc}")

    for d in to_delete_dirs:
        try:
            shutil.rmtree(d)
            deleted_dirs.append(d.name)
        except Exception as exc:
            errors.append(f"{d.name}/: {exc}")

    return {
        "status": "ok" if not errors else "partial",
        "game_dir": str(base),
        "gelöschte_dateien": deleted_files,
        "gelöschte_ordner": deleted_dirs,
        "fehler": errors if errors else None,
        "backup": backup_path,
        "message": (
            f"ReShade deinstalliert: {len(deleted_files)} Dateien, "
            f"{len(deleted_dirs)} Ordner gelöscht."
            + (f" Backup unter: {backup_path}" if backup_path else "")
        ),
    }


# ===========================================================================
# OpenXR Toolkit Integration — Registry-based VR graphics settings
# ===========================================================================

# OpenXR Toolkit stores ALL settings as DWORD values in the Windows Registry:
#   HKCU\SOFTWARE\OpenXR_Toolkit\<game.exe>
# Changes take effect on the next VR session (no restart of the game needed,
# but the OpenXR API layer re-reads on session start).

_OPENXR_REG_BASE = r"HKCU\SOFTWARE\OpenXR_Toolkit"
_OPENXR_MSFS2024_EXE = "FlightSimulator2024.exe"

# All known OpenXR Toolkit setting keys with metadata
OPENXR_SETTING_DEFS: dict[str, dict] = {
    # ── Upscaling & Schärfe ──
    "scaling_type": {
        "label": "Upscaler",
        "values": {"0": "Off", "1": "NIS", "2": "FSR", "3": "CAS"},
        "category": "upscaling",
    },
    "scaling": {
        "label": "Skalierung (%)",
        "category": "upscaling",
    },
    "anamorphic": {
        "label": "Anamorphic",
        "category": "upscaling",
    },
    "sharpness": {
        "label": "Schärfe",
        "category": "upscaling",
    },
    "mipmap_bias": {
        "label": "MipMap Bias",
        "category": "upscaling",
    },
    # ── Post-Processing / Farbe ──
    "post_process": {
        "label": "Post-Processing",
        "values": {"0": "Off", "1": "On"},
        "category": "color",
    },
    "post_brightness": {"label": "Helligkeit", "category": "color"},
    "post_contrast": {"label": "Kontrast", "category": "color"},
    "post_exposure": {"label": "Belichtung", "category": "color"},
    "post_saturation": {"label": "Sättigung", "category": "color"},
    "post_vibrance": {"label": "Vibrance", "category": "color"},
    "post_gain_r": {"label": "Farbe Rot", "category": "color"},
    "post_gain_g": {"label": "Farbe Grün", "category": "color"},
    "post_gain_b": {"label": "Farbe Blau", "category": "color"},
    "post_highlights": {"label": "Lichter", "category": "color"},
    "post_shadows": {"label": "Schatten", "category": "color"},
    "post_sunglasses": {
        "label": "Sonnenbrille",
        "values": {"0": "Off", "1": "Light", "2": "Dark", "3": "Night"},
        "category": "color",
    },
    # ── Variable Rate Shading (Foveated Rendering) ──
    "vrs": {
        "label": "Foveated Rendering",
        "values": {"0": "Off", "1": "Preset", "2": "Custom"},
        "category": "vrs",
    },
    "vrs_inner": {
        "label": "VRS Innen",
        "values": {"0": "1x", "1": "1/2", "2": "1/4", "3": "1/8", "4": "1/16"},
        "category": "vrs",
    },
    "vrs_middle": {
        "label": "VRS Mitte",
        "values": {"0": "1x", "1": "1/2", "2": "1/4", "3": "1/8", "4": "1/16"},
        "category": "vrs",
    },
    "vrs_outer": {
        "label": "VRS Außen",
        "values": {"0": "1x", "1": "1/2", "2": "1/4", "3": "1/8", "4": "1/16"},
        "category": "vrs",
    },
    "vrs_inner_radius": {"label": "VRS Innen-Radius", "category": "vrs"},
    "vrs_outer_radius": {"label": "VRS Außen-Radius", "category": "vrs"},
    "vrs_x_offset": {"label": "VRS X-Offset", "category": "vrs"},
    "vrs_y_offset": {"label": "VRS Y-Offset", "category": "vrs"},
    "vrs_cull_mask": {
        "label": "VRS HAM Culling",
        "values": {"0": "Off", "1": "On"},
        "category": "vrs",
    },
    # ── World / FOV ──
    "world_scale": {"label": "Weltgröße (IPD)", "category": "world"},
    "fov": {"label": "FOV", "category": "world"},
    "fov_up": {"label": "FOV Oben", "category": "world"},
    "fov_down": {"label": "FOV Unten", "category": "world"},
    "zoom": {"label": "Zoom", "category": "world"},
    # ── Performance ──
    "turbo": {
        "label": "Turbo Mode",
        "values": {"0": "Off", "1": "On"},
        "category": "performance",
    },
    "motion_reprojection": {
        "label": "Motion Reprojection",
        "values": {"0": "Default", "1": "Off", "2": "On"},
        "category": "performance",
    },
    "motion_reprojection_rate": {
        "label": "Reprojection Rate",
        "values": {"1": "Off", "2": "45 Hz", "3": "30 Hz", "4": "22 Hz"},
        "category": "performance",
    },
    "frame_throttle": {"label": "Frame Throttle", "category": "performance"},
    "target_rate": {"label": "Ziel-FPS", "category": "performance"},
    "override_resolution": {
        "label": "Auflösung überschreiben",
        "values": {"0": "Off", "1": "On"},
        "category": "performance",
    },
    "resolution_height": {"label": "Auflösung Höhe", "category": "performance"},
}

# German aliases → registry key name
_OPENXR_ALIASES: dict[str, str] = {
    # Upscaling
    "upscaler": "scaling_type",
    "upscaling": "scaling_type",
    "nis": "scaling_type",
    "fsr": "scaling_type",
    "cas": "scaling_type",
    "skalierung": "scaling",
    "scale": "scaling",
    "schärfe": "sharpness",
    "sharpness": "sharpness",
    "sharpen": "sharpness",
    "mipmap": "mipmap_bias",
    # Post-Processing
    "post processing": "post_process",
    "postprocessing": "post_process",
    "nachbearbeitung": "post_process",
    "helligkeit": "post_brightness",
    "brightness": "post_brightness",
    "kontrast": "post_contrast",
    "contrast": "post_contrast",
    "belichtung": "post_exposure",
    "exposure": "post_exposure",
    "sättigung": "post_saturation",
    "saturation": "post_saturation",
    "vibrance": "post_vibrance",
    "farbintensität": "post_vibrance",
    "rot": "post_gain_r",
    "grün": "post_gain_g",
    "blau": "post_gain_b",
    "lichter": "post_highlights",
    "highlights": "post_highlights",
    "schatten": "post_shadows",
    "shadows": "post_shadows",
    "sonnenbrille": "post_sunglasses",
    "sunglasses": "post_sunglasses",
    # VRS / Foveated
    "foveated": "vrs",
    "foveated rendering": "vrs",
    "variable rate shading": "vrs",
    "vrs innen": "vrs_inner",
    "vrs mitte": "vrs_middle",
    "vrs außen": "vrs_outer",
    # World
    "weltgröße": "world_scale",
    "world scale": "world_scale",
    "ipd": "world_scale",
    "sichtfeld": "fov",
    "field of view": "fov",
    "zoom": "zoom",
    # Performance
    "turbo": "turbo",
    "reprojection": "motion_reprojection",
    "motion reprojection": "motion_reprojection",
    "reprojection rate": "motion_reprojection_rate",
    "frame throttle": "frame_throttle",
    "fps limit": "target_rate",
    "ziel fps": "target_rate",
    "target fps": "target_rate",
    "auflösung": "override_resolution",
}

# Named value aliases for enum-type settings
_OPENXR_VALUE_ALIASES: dict[str, dict[str, str]] = {
    "scaling_type": {
        "off": "0", "aus": "0", "kein": "0",
        "nis": "1", "nvidia": "1",
        "fsr": "2", "amd": "2", "fidelityfx": "2",
        "cas": "3",
    },
    "post_sunglasses": {
        "off": "0", "aus": "0",
        "light": "1", "hell": "1", "leicht": "1",
        "dark": "2", "dunkel": "2",
        "night": "3", "nacht": "3",
    },
    "vrs": {
        "off": "0", "aus": "0",
        "preset": "1", "voreinstellung": "1",
        "custom": "2", "benutzerdefiniert": "2",
    },
    "motion_reprojection": {
        "default": "0", "standard": "0",
        "off": "1", "aus": "1",
        "on": "2", "an": "2", "ein": "2",
    },
    "motion_reprojection_rate": {
        "off": "1", "aus": "1",
        "45": "2", "45hz": "2",
        "30": "3", "30hz": "3",
        "22": "4", "22hz": "4",
    },
    "_on_off": {
        "off": "0", "aus": "0", "nein": "0",
        "on": "1", "an": "1", "ein": "1", "ja": "1",
    },
}

_OPENXR_ON_OFF_KEYS = {
    "post_process", "turbo", "override_resolution", "vrs_cull_mask",
}


def _openxr_reg_path(game_exe: str = "") -> str:
    """Return the full registry path for a game's OpenXR Toolkit settings."""
    exe = game_exe or _OPENXR_MSFS2024_EXE
    return f"{_OPENXR_REG_BASE}\\{exe}"


def _read_openxr_settings(game_exe: str = "") -> dict[str, int]:
    """Read all OpenXR Toolkit settings from the registry for a game."""
    reg_path = _openxr_reg_path(game_exe)
    result: dict[str, int] = {}
    try:
        r = _reg_query(reg_path)
        if r.returncode != 0:
            return result
        for line in r.stdout.splitlines():
            m = re.match(r"\s+(\S+)\s+REG_DWORD\s+0x([0-9a-fA-F]+)", line)
            if m:
                key = m.group(1)
                val = int(m.group(2), 16)
                result[key] = val
    except Exception:
        pass
    return result


def _write_openxr_setting(key: str, value: int, game_exe: str = "") -> bool:
    """Write a single DWORD value to the OpenXR Toolkit registry."""
    reg_path = _openxr_reg_path(game_exe)
    try:
        r = _reg_add(reg_path, key, "REG_DWORD", str(value))
        return r.returncode == 0
    except Exception:
        return False


# ── OpenXR Runtime Management ─────────────────────────────────────────────

# Registry path where the active OpenXR runtime is stored
_OPENXR_RUNTIME_REG = r"HKLM\SOFTWARE\Khronos\OpenXR\1"
_OPENXR_RUNTIME_VALUE = "ActiveRuntime"

# Known OpenXR runtime JSON manifest paths
_OPENXR_RUNTIMES: dict[str, list[str]] = {
    "pimax": [
        r"C:\Program Files\Pimax\Runtime\PiOpenXR_64.json",
        r"C:\Program Files (x86)\Pimax\Runtime\PiOpenXR_64.json",
        r"D:\Program Files\Pimax\Runtime\PiOpenXR_64.json",
    ],
    "steamvr": [
        r"C:\Program Files (x86)\Steam\steamapps\common\SteamVR\steamxr_win64.json",
        r"D:\Steam\steamapps\common\SteamVR\steamxr_win64.json",
        r"D:\SteamLibrary\steamapps\common\SteamVR\steamxr_win64.json",
        r"E:\SteamLibrary\steamapps\common\SteamVR\steamxr_win64.json",
    ],
    "oculus": [
        r"C:\Program Files\Oculus\Support\oculus-runtime\oculus_openxr_64.json",
    ],
    "wmr": [
        r"C:\WINDOWS\system32\MixedRealityRuntime.json",
    ],
}

# German aliases for runtime names
_RUNTIME_ALIASES: dict[str, str] = {
    "pimax": "pimax",
    "pimaxplay": "pimax",
    "pimax play": "pimax",
    "steamvr": "steamvr",
    "steam vr": "steamvr",
    "steam": "steamvr",
    "oculus": "oculus",
    "meta": "oculus",
    "quest": "oculus",
    "meta quest": "oculus",
    "wmr": "wmr",
    "windows mixed reality": "wmr",
    "mixed reality": "wmr",
}


def _get_active_runtime() -> dict:
    """Read the currently active OpenXR runtime from the registry."""
    try:
        r = _reg(["query", _OPENXR_RUNTIME_REG, "/v", _OPENXR_RUNTIME_VALUE])
        for line in r.stdout.splitlines():
            m = re.search(r"ActiveRuntime\s+REG_SZ\s+(.+)", line)
            if m:
                path = m.group(1).strip()
                # Identify which runtime this is
                path_lower = path.lower()
                name = "unbekannt"
                for rt_name, candidates in _OPENXR_RUNTIMES.items():
                    for c in candidates:
                        if c.lower() in path_lower or Path(c).name.lower() in path_lower:
                            name = rt_name
                            break
                    if name != "unbekannt":
                        break
                return {"path": path, "name": name}
    except Exception:
        pass
    return {"path": None, "name": None}


def _find_runtime_path(runtime: str) -> str | None:
    """Find the JSON manifest path for a given runtime name."""
    rt = _RUNTIME_ALIASES.get(runtime.lower().strip(), runtime.lower().strip())
    candidates = _OPENXR_RUNTIMES.get(rt, [])
    for c in candidates:
        if Path(c).exists():
            return c

    # Search registry for installed runtimes
    try:
        r = _reg(["query", r"HKLM\SOFTWARE\Khronos\OpenXR\1\AvailableRuntimes"])
        for line in r.stdout.splitlines():
            if rt in line.lower():
                m = re.match(r"\s+(.+\.json)\s+REG_DWORD", line)
                if m:
                    return m.group(1).strip()
    except Exception:
        pass

    # Search filesystem as last resort
    if rt == "pimax":
        for root in [r"C:\Program Files", r"C:\Program Files (x86)", r"D:\Program Files"]:
            rp = Path(root)
            if rp.exists():
                for hit in rp.glob("**/PiOpenXR_64.json"):
                    return str(hit)

    return None


@mcp.tool()
def set_openxr_runtime(
    runtime: str = "pimax",
) -> dict:
    """Set the active OpenXR runtime (e.g. to Pimax, SteamVR, Oculus).

    Use this when the user says things like:
    - "OpenXR Runtime auf Pimax stellen"
    - "OpenXR auf Pimax umstellen"
    - "Pimax als OpenXR Runtime"
    - "SteamVR Runtime aktivieren"
    - "OpenXR Runtime wurde verstellt"

    Writes to: HKLM\\SOFTWARE\\Khronos\\OpenXR\\1\\ActiveRuntime
    (requires Admin-Rechte — wird automatisch elevated).

    Args:
        runtime: Runtime name — "pimax", "steamvr", "oculus", "wmr",
                 or full path to the runtime JSON manifest.
    """
    # Check current runtime
    current = _get_active_runtime()

    # If user gave a full path, use it directly
    if runtime.lower().endswith(".json") or "\\" in runtime or "/" in runtime:
        target_path = runtime
        target_name = "custom"
        if not Path(target_path).exists():
            return {
                "error": f"Runtime-Datei nicht gefunden: {target_path}",
            }
    else:
        # Resolve name → path
        rt_key = _RUNTIME_ALIASES.get(runtime.lower().strip())
        if rt_key is None:
            return {
                "error": f"Unbekannte Runtime: '{runtime}'",
                "verfügbar": ["pimax", "steamvr", "oculus", "wmr"],
            }
        target_name = rt_key
        target_path = _find_runtime_path(rt_key)
        if target_path is None:
            return {
                "error": (
                    f"Runtime '{rt_key}' nicht gefunden. "
                    f"Gesucht: {_OPENXR_RUNTIMES.get(rt_key, [])}"
                ),
                "tipp": "Gib den vollständigen Pfad zur .json Datei an.",
            }

    # Already set?
    if current["path"] and Path(current["path"]).resolve() == Path(target_path).resolve():
        return {
            "status": "already_set",
            "runtime": target_name,
            "path": target_path,
            "message": f"OpenXR Runtime ist bereits auf {target_name} eingestellt.",
        }

    # Set the new runtime (HKLM — needs admin)
    r = _reg_add(_OPENXR_RUNTIME_REG, _OPENXR_RUNTIME_VALUE, "REG_SZ", target_path)

    # Verify
    new_current = _get_active_runtime()
    verified = (new_current["path"] and
                Path(new_current["path"]).resolve() == Path(target_path).resolve())

    if not verified:
        return {
            "error": (
                f"Konnte Runtime nicht umstellen. "
                f"Registry-Schreibzugriff auf HKLM benötigt Admin-Rechte."
            ),
            "vorher": current,
            "versucht": target_path,
            "tipp": (
                "Starte den MCP-Server als Administrator, oder führe manuell aus:\n"
                f'reg add "HKLM\\SOFTWARE\\Khronos\\OpenXR\\1" '
                f'/v ActiveRuntime /t REG_SZ /d "{target_path}" /f'
            ),
        }

    return {
        "status": "ok",
        "vorher": {
            "runtime": current["name"],
            "path": current["path"],
        },
        "nachher": {
            "runtime": target_name,
            "path": target_path,
        },
        "verifiziert": verified,
        "message": (
            f"OpenXR Runtime umgestellt: {current['name']} → {target_name}. "
            f"Wird beim nächsten VR-Start aktiv."
        ),
    }


@mcp.tool()
def get_openxr_runtime() -> dict:
    """Show the currently active OpenXR runtime and all available runtimes.

    Use this when the user asks:
    - "Welche OpenXR Runtime ist aktiv?"
    - "Zeige OpenXR Runtime"
    - "Ist Pimax als Runtime eingestellt?"

    Reads from: HKLM\\SOFTWARE\\Khronos\\OpenXR\\1
    """
    current = _get_active_runtime()

    # Find all available runtimes
    available: list[dict] = []
    for rt_name, candidates in _OPENXR_RUNTIMES.items():
        for c in candidates:
            if Path(c).exists():
                is_active = (current["path"] and
                             Path(current["path"]).resolve() == Path(c).resolve())
                available.append({
                    "runtime": rt_name,
                    "path": c,
                    "aktiv": is_active,
                })
                break  # Only first existing path per runtime

    # Also check AvailableRuntimes registry
    try:
        r = _reg(["query", r"HKLM\SOFTWARE\Khronos\OpenXR\1\AvailableRuntimes"])
        for line in r.stdout.splitlines():
            m = re.match(r"\s+(.+\.json)\s+REG_DWORD\s+0x0", line)
            if m:
                rt_path = m.group(1).strip()
                # Skip if already in our list
                if not any(a["path"] == rt_path for a in available):
                    is_active = (current["path"] and
                                 Path(current["path"]).resolve() == Path(rt_path).resolve())
                    available.append({
                        "runtime": Path(rt_path).stem,
                        "path": rt_path,
                        "aktiv": is_active,
                    })
    except Exception:
        pass

    return {
        "status": "ok",
        "aktive_runtime": current,
        "verfügbare_runtimes": available,
        "registry_path": _OPENXR_RUNTIME_REG,
        "hinweis": (
            "Zum Umstellen: set_openxr_runtime(runtime='pimax') verwenden. "
            "Oder sage einfach 'OpenXR auf Pimax stellen'."
        ),
    }


# ── OpenXR Toolkit MCP Tools ─────────────────────────────────────────────

@mcp.tool()
def analyze_openxr(
    game_exe: str = "",
) -> dict:
    """Analyze current OpenXR Toolkit VR graphics settings (Registry).

    WICHTIG: Dies ist das OpenXR Toolkit — NICHT die MSFS-Grafikeinstellungen!
    - OpenXR Toolkit = Upscaling (NIS/FSR/CAS), Schärfe, Foveated Rendering,
      Post-Processing (Helligkeit/Kontrast/Sättigung), Motion Reprojection
    - MSFS Grafik = DLSS, Wolken, Terrain LoD, Schatten (→ analyze_msfs_graphics)

    Reads all OpenXR Toolkit settings from the Windows Registry at:
      HKCU\\SOFTWARE\\OpenXR_Toolkit\\FlightSimulator2024.exe

    Use this tool when the user mentions:
    - "OpenXR" / "OpenXR Toolkit"
    - "VR Schärfe" / "VR Upscaling" / "NIS" / "FSR"
    - "Foveated Rendering" / "VRS"
    - "VR Helligkeit" / "VR Kontrast" (Post-Processing)
    - "Motion Reprojection"

    Args:
        game_exe: Game executable name (default: FlightSimulator2024.exe).
    """
    settings = _read_openxr_settings(game_exe)
    exe = game_exe or _OPENXR_MSFS2024_EXE

    if not settings:
        return {
            "error": (
                f"Keine OpenXR Toolkit Einstellungen gefunden für '{exe}'. "
                "Ist OpenXR Toolkit installiert und wurde MSFS 2024 schon "
                "einmal in VR gestartet?"
            ),
            "registry_path": _openxr_reg_path(game_exe),
        }

    # Group by category
    categories: dict[str, list[dict]] = {}
    for key, val in sorted(settings.items()):
        defn = OPENXR_SETTING_DEFS.get(key, {})
        label = defn.get("label", key)
        category = defn.get("category", "sonstige")
        value_map = defn.get("values", {})
        display = value_map.get(str(val), str(val))

        categories.setdefault(category, []).append({
            "key": key,
            "label": label,
            "value": val,
            "display": display,
        })

    return {
        "status": "ok",
        "game": exe,
        "registry_path": _openxr_reg_path(game_exe),
        "categories": categories,
        "display_instructions": (
            "Zeige die Einstellungen als Markdown-Tabellen gruppiert nach Kategorie. "
            "Kategorien: Upscaling, Farbe/Post-Processing, Foveated Rendering (VRS), "
            "World/FOV, Performance, Sonstige. Antworte auf Deutsch."
        ),
    }


@mcp.tool()
def set_openxr_setting(
    setting: str,
    value: str,
    game_exe: str = "",
) -> dict:
    """Change an OpenXR Toolkit VR setting in the Windows Registry.

    WICHTIG: Dies ändert OpenXR Toolkit Einstellungen — NICHT MSFS-Grafik!
    Für MSFS-Grafik (DLSS, Wolken, Terrain) → set_msfs_setting verwenden.

    OpenXR Toolkit steuert: Upscaling (NIS/FSR/CAS), Schärfe, Foveated
    Rendering (VRS), Post-Processing (Helligkeit/Kontrast), Reprojection.

    Writes directly to: HKCU\\SOFTWARE\\OpenXR_Toolkit\\FlightSimulator2024.exe
    Changes take effect at the next VR session start.

    Use this when the user mentions "OpenXR" and wants to change:
    - "OpenXR Schärfe auf 80" / "NIS Upscaling aktivieren"
    - "OpenXR Foveated Rendering einschalten"
    - "OpenXR Helligkeit erhöhen" / "OpenXR Kontrast auf 500"
    - "Motion Reprojection an" / "Turbo Mode aus"
    - "Sonnenbrille auf Dark"

    Supports relative values: "+10", "-20" to adjust from current.

    Args:
        setting: Setting name (German or English), e.g. "schärfe",
                 "upscaler", "helligkeit", "foveated", "turbo".
        value: New value — name ("nis", "aus", "dark") or number ("80"),
               or relative ("+10", "-5").
        game_exe: Game executable (default: FlightSimulator2024.exe).
    """
    exe = game_exe or _OPENXR_MSFS2024_EXE

    # Resolve alias → registry key
    setting_lower = setting.strip().lower()
    reg_key = _OPENXR_ALIASES.get(setting_lower)

    if reg_key is None:
        # Direct match against known keys
        if setting_lower in OPENXR_SETTING_DEFS:
            reg_key = setting_lower
        else:
            for k in OPENXR_SETTING_DEFS:
                if setting_lower in k or k in setting_lower:
                    reg_key = k
                    break

    if reg_key is None:
        return {
            "error": f"Unbekannte OpenXR-Einstellung: '{setting}'",
            "verfügbar": sorted(_OPENXR_ALIASES.keys()),
        }

    defn = OPENXR_SETTING_DEFS.get(reg_key, {})
    label = defn.get("label", reg_key)
    value_map = defn.get("values", {})

    # Read current value
    current = _read_openxr_settings(exe)
    old_val = current.get(reg_key, 0)

    # Resolve value
    value_lower = value.strip().lower()
    resolved: int | None = None

    # 1. Check named value aliases
    alias_group = _OPENXR_VALUE_ALIASES.get(reg_key)
    if alias_group is None and reg_key in _OPENXR_ON_OFF_KEYS:
        alias_group = _OPENXR_VALUE_ALIASES["_on_off"]

    if alias_group and value_lower in alias_group:
        resolved = int(alias_group[value_lower])

    # 2. Special: if user says "NIS"/"FSR"/"CAS" as value for scaling_type
    if resolved is None and reg_key == "scaling_type":
        for name, num in _OPENXR_VALUE_ALIASES["scaling_type"].items():
            if value_lower == name:
                resolved = int(num)
                break

    # 3. Relative adjustment: "+10", "-5"
    if resolved is None:
        v_stripped = value.strip()
        if v_stripped.startswith("+") or (v_stripped.startswith("-") and len(v_stripped) > 1):
            try:
                delta = int(v_stripped)
                resolved = old_val + delta
            except ValueError:
                pass

    # 4. Direct integer
    if resolved is None:
        try:
            resolved = int(value.strip())
        except ValueError:
            pass

    if resolved is None:
        return {
            "error": f"Kann Wert '{value}' nicht interpretieren für '{label}'.",
            "erlaubte_werte": value_map if value_map else "Ganzzahl",
        }

    # Write to registry
    ok = _write_openxr_setting(reg_key, resolved, exe)

    old_display = value_map.get(str(old_val), str(old_val))
    new_display = value_map.get(str(resolved), str(resolved))

    if not ok:
        return {"error": f"Konnte '{reg_key}' nicht in Registry schreiben."}

    # Verify
    verify = _read_openxr_settings(exe)
    verified = verify.get(reg_key) == resolved

    return {
        "status": "ok",
        "einstellung": label,
        "key": reg_key,
        "vorher": old_display,
        "nachher": new_display,
        "raw_value": resolved,
        "registry_path": f"{_openxr_reg_path(exe)}\\{reg_key}",
        "verifiziert": verified,
        "hinweis": (
            "Einstellung in Registry geschrieben. Wird bei der nächsten "
            "VR-Session aktiv. Falls MSFS in VR läuft: VR-Modus beenden "
            "und neu starten (Ctrl+Tab oder MSFS neu starten)."
        ),
    }


@mcp.tool()
def apply_openxr_preset(
    preset: Literal["Performance", "Balanced", "Quality", "Auto"] = "Auto",
    game_exe: str = "",
) -> dict:
    """Apply a GPU-aware OpenXR Toolkit VR preset for MSFS 2024.

    Detects your GPU automatically and selects appropriate OpenXR Toolkit
    settings. "Auto" (default) picks the best preset for your GPU:
    - Flagship (5090/4090) → Quality
    - High-end (5080/4080/3090) → Balanced
    - Mid/Entry (4070 and below) → Performance

    Manual override: Performance / Balanced / Quality.

    Steuert: NIS/FSR Upscaling, Schärfe, Foveated Rendering (VRS),
    Post-Processing, Turbo Mode, Motion Reprojection.

    Args:
        preset: "Auto" (GPU-basiert), "Performance", "Balanced", oder "Quality".
        game_exe: Game executable (default: FlightSimulator2024.exe).
    """
    # GPU detection for Auto mode
    gpu_name, gpu_tier, vram_mb = _detect_gpu_tier()

    if preset == "Auto":
        _TIER_TO_OPENXR = {
            "flagship": "Quality",
            "high_end": "Balanced",
            "mid_high": "Performance",
            "mid_range": "Performance",
            "entry": "Performance",
            "unknown": "Balanced",
        }
        preset = _TIER_TO_OPENXR.get(gpu_tier, "Balanced")

    presets: dict[str, dict[str, int]] = {
        "Performance": {
            "scaling_type": 1,    # NIS
            "sharpness": 70,
            "post_process": 0,    # Off (spart FPS)
            "vrs": 1,             # Preset
            "vrs_inner": 0,       # 1x (scharf in der Mitte)
            "vrs_middle": 2,      # 1/4
            "vrs_outer": 3,       # 1/8
            "turbo": 1,           # On
            "motion_reprojection": 2,  # On
        },
        "Balanced": {
            "scaling_type": 1,    # NIS
            "sharpness": 50,
            "post_process": 1,    # On
            "post_contrast": 500,
            "post_saturation": 300,
            "vrs": 1,             # Preset
            "vrs_inner": 0,       # 1x
            "vrs_middle": 1,      # 1/2
            "vrs_outer": 2,       # 1/4
            "turbo": 0,           # Off
            "motion_reprojection": 0,  # Default
        },
        "Quality": {
            "scaling_type": 3,    # CAS (nur Schärfe, kein Upscaling)
            "sharpness": 40,
            "post_process": 1,    # On
            "post_contrast": 400,
            "post_saturation": 200,
            "post_vibrance": 200,
            "vrs": 0,             # Off
            "turbo": 0,           # Off
            "motion_reprojection": 0,  # Default
        },
    }

    preset_def = presets.get(preset)
    if not preset_def:
        return {"error": f"Unbekanntes Preset: {preset}"}

    exe = game_exe or _OPENXR_MSFS2024_EXE

    # Backup current settings
    current = _read_openxr_settings(exe)

    # Apply all settings
    applied: dict[str, dict] = {}
    failed: list[str] = []
    for key, new_val in preset_def.items():
        old_val = current.get(key, 0)
        defn = OPENXR_SETTING_DEFS.get(key, {})
        value_map = defn.get("values", {})

        ok = _write_openxr_setting(key, new_val, exe)
        if ok:
            applied[key] = {
                "label": defn.get("label", key),
                "vorher": value_map.get(str(old_val), str(old_val)),
                "nachher": value_map.get(str(new_val), str(new_val)),
            }
        else:
            failed.append(key)

    return {
        "status": "ok" if not failed else "partial",
        "preset": preset,
        "gpu": gpu_name,
        "gpu_tier": gpu_tier,
        "game": exe,
        "geändert": applied,
        "fehlgeschlagen": failed if failed else None,
        "hinweis": (
            f"OpenXR Toolkit '{preset}'-Preset angewendet "
            f"(GPU: {gpu_name}, Tier: {gpu_tier}). "
            "Wird bei der nächsten VR-Session aktiv. "
            "MSFS VR beenden und neu starten damit es wirkt."
        ),
    }


# ---------------------------------------------------------------------------
# Hardware Profile & VR Performance Tier
# ---------------------------------------------------------------------------

# VR performance tiers based on combined CPU+GPU+RAM score
_VR_TIERS = [
    ("ultra",   "RTX 4090 / RX 7900 XTX level — maximale VR-Qualität möglich"),
    ("high",    "RTX 4080/4070 Ti / RX 7900 XT — sehr gute VR-Qualität"),
    ("mid_high","RTX 4070/3080 / RX 6800 XT — gute VR-Qualität mit Kompromissen"),
    ("mid",     "RTX 3070/3060 Ti / RX 6700 XT — mittlere VR-Qualität"),
    ("low",     "RTX 3060 oder schwächer — Performance-Modus empfohlen"),
]


def _classify_vr_tier(vram_gb: float, gpu_name: str, cpu_cores: int, ram_gb: float) -> str:
    name = gpu_name.upper()
    if vram_gb >= 20 or any(x in name for x in ["4090", "4080", "7900 XTX", "7900 XT"]):
        return "ultra"
    if vram_gb >= 16 or any(x in name for x in ["4070 TI", "4070TI", "3090", "6900", "7900"]):
        return "high"
    if vram_gb >= 10 or any(x in name for x in ["4070", "3080", "6800", "7800", "4060 TI"]):
        return "mid_high"
    if vram_gb >= 8 or any(x in name for x in ["3070", "3060 TI", "6700", "6750", "4060"]):
        return "mid"
    return "low"


@mcp.tool()
def get_detailed_hardware_profile() -> dict:
    """Liest alle Hardware-Details aus (GPU, CPU, RAM, Speicher) und berechnet den VR-Performance-Tier.

    Rufe dieses Tool ZUERST auf bevor du VR-Einstellungen optimierst.
    Gibt GPU-VRAM, CPU-Kerne, RAM, Speicher-Typ und den VR-Tier zurück.
    """
    # GPU via pynvml
    gpu_info = {}
    try:
        pynvml.nvmlInit()
        try:
            h = pynvml.nvmlDeviceGetHandleByIndex(0)
            name = pynvml.nvmlDeviceGetName(h)
            if isinstance(name, bytes):
                name = name.decode()
            mem = pynvml.nvmlDeviceGetMemoryInfo(h)
            temp = pynvml.nvmlDeviceGetTemperature(h, pynvml.NVML_TEMPERATURE_GPU)
            clock = pynvml.nvmlDeviceGetClockInfo(h, pynvml.NVML_CLOCK_GRAPHICS)
            max_clock = pynvml.nvmlDeviceGetMaxClockInfo(h, pynvml.NVML_CLOCK_GRAPHICS)
            util = pynvml.nvmlDeviceGetUtilizationRates(h)
            gpu_info = {
                "name": name,
                "vram_total_gb": round(mem.total / 1024**3, 1),
                "vram_free_gb": round(mem.free / 1024**3, 1),
                "temperature_c": temp,
                "clock_mhz": clock,
                "max_clock_mhz": max_clock,
                "utilization_gpu_pct": util.gpu,
                "utilization_mem_pct": util.memory,
            }
        finally:
            pynvml.nvmlShutdown()
    except Exception as exc:
        gpu_info = {"error": str(exc)}

    # CPU
    try:
        cpu_raw = _ps_json(
            "Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, "
            "NumberOfLogicalProcessors, MaxClockSpeed, CurrentClockSpeed, L2CacheSize, L3CacheSize"
        )
        cpu_info = cpu_raw[0] if cpu_raw else {}
    except Exception:
        cpu_info = {}

    # RAM
    try:
        ram_raw = _ps_json(
            "Get-CimInstance Win32_PhysicalMemory | Select-Object "
            "Capacity, Speed, MemoryType, SMBIOSMemoryType, Manufacturer, PartNumber"
        )
        total_ram_bytes = sum(int(r.get("Capacity", 0) or 0) for r in ram_raw)
        ram_gb = round(total_ram_bytes / 1024**3, 1)
        mem_type_map = {24: "DDR3", 26: "DDR4", 34: "DDR5"}
        mem_type = mem_type_map.get(int(ram_raw[0].get("SMBIOSMemoryType", 0) or 0), "DDR?") if ram_raw else "?"
        ram_speed = ram_raw[0].get("Speed", "?") if ram_raw else "?"
        ram_info = {
            "total_gb": ram_gb,
            "type": mem_type,
            "speed_mhz": ram_speed,
            "modules": len(ram_raw),
        }
    except Exception as exc:
        ram_info = {"error": str(exc)}
        ram_gb = 16.0

    # Disk (MSFS-relevant)
    try:
        disk_raw = _ps_json(
            "Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size, BusType"
        )
        disks = [
            {
                "name": d.get("FriendlyName", ""),
                "type": d.get("MediaType", ""),
                "bus": d.get("BusType", ""),
                "size_gb": round(int(d.get("Size", 0) or 0) / 1024**3, 0),
            }
            for d in disk_raw
        ]
    except Exception:
        disks = []

    # Compute VR tier
    vram_gb = gpu_info.get("vram_total_gb", 8.0) if "error" not in gpu_info else 8.0
    cpu_cores = int(cpu_info.get("NumberOfCores", 8) or 8)
    gpu_name = gpu_info.get("name", "") if "error" not in gpu_info else ""
    tier = _classify_vr_tier(vram_gb, gpu_name, cpu_cores, ram_gb)
    tier_desc = dict(_VR_TIERS).get(tier, "")

    return {
        "gpu": gpu_info,
        "cpu": cpu_info,
        "ram": ram_info,
        "disks": disks,
        "vr_tier": tier,
        "vr_tier_beschreibung": tier_desc,
        "empfehlung": (
            f"Dein System ({gpu_name}, {vram_gb:.0f}GB VRAM, {cpu_cores} CPU-Kerne, {ram_gb:.0f}GB RAM) "
            f"ist im VR-Tier '{tier}'. Sage 'VR optimieren' für hardware-spezifische Einstellungen."
        ),
    }


# ---------------------------------------------------------------------------
# Master VR Optimizer — sets everything based on hardware
# ---------------------------------------------------------------------------

# Preset tables per tier: MSFS GraphicsVR settings
_VR_MSFS_PRESETS: dict[str, dict[str, str]] = {
    "ultra": {
        "GraphicsVR.TerrainLoD": "2.5",
        "GraphicsVR.ObjectsLoD": "2.5",
        "GraphicsVR.CloudsQuality": "3",
        "GraphicsVR.TextureResolution": "3",
        "GraphicsVR.AnisotropicFilter": "4",
        "GraphicsVR.ShadowQuality": "3",
        "GraphicsVR.Reflections": "4",
        "GraphicsVR.SSContact": "1",
        "GraphicsVR.MotionBlur": "0",
        "Graphics.TerrainLoD": "2.5",
        "Graphics.ObjectsLoD": "2.5",
        "Graphics.CloudsQuality": "3",
        "Graphics.TextureResolution": "3",
        "Graphics.AnisotropicFilter": "4",
        "Graphics.ShadowQuality": "3",
        "Video.DLSS": "1",
        "Video.DLSSG": "1",
        "Video.RenderScale": "100",
        "RayTracing.Enabled": "0",
    },
    "high": {
        "GraphicsVR.TerrainLoD": "2.0",
        "GraphicsVR.ObjectsLoD": "2.0",
        "GraphicsVR.CloudsQuality": "3",
        "GraphicsVR.TextureResolution": "3",
        "GraphicsVR.AnisotropicFilter": "4",
        "GraphicsVR.ShadowQuality": "2",
        "GraphicsVR.Reflections": "2",
        "GraphicsVR.SSContact": "1",
        "GraphicsVR.MotionBlur": "0",
        "Graphics.TerrainLoD": "2.0",
        "Graphics.ObjectsLoD": "2.0",
        "Graphics.CloudsQuality": "3",
        "Graphics.TextureResolution": "3",
        "Graphics.AnisotropicFilter": "4",
        "Graphics.ShadowQuality": "2",
        "Video.DLSS": "2",
        "Video.DLSSG": "0",
        "Video.RenderScale": "100",
        "RayTracing.Enabled": "0",
    },
    "mid_high": {
        "GraphicsVR.TerrainLoD": "1.5",
        "GraphicsVR.ObjectsLoD": "1.5",
        "GraphicsVR.CloudsQuality": "2",
        "GraphicsVR.TextureResolution": "2",
        "GraphicsVR.AnisotropicFilter": "4",
        "GraphicsVR.ShadowQuality": "2",
        "GraphicsVR.Reflections": "1",
        "GraphicsVR.SSContact": "1",
        "GraphicsVR.MotionBlur": "0",
        "Graphics.TerrainLoD": "1.5",
        "Graphics.ObjectsLoD": "1.5",
        "Graphics.CloudsQuality": "2",
        "Graphics.TextureResolution": "2",
        "Graphics.AnisotropicFilter": "4",
        "Graphics.ShadowQuality": "2",
        "Video.DLSS": "3",
        "Video.DLSSG": "0",
        "Video.RenderScale": "100",
        "RayTracing.Enabled": "0",
    },
    "mid": {
        "GraphicsVR.TerrainLoD": "1.0",
        "GraphicsVR.ObjectsLoD": "1.0",
        "GraphicsVR.CloudsQuality": "1",
        "GraphicsVR.TextureResolution": "2",
        "GraphicsVR.AnisotropicFilter": "3",
        "GraphicsVR.ShadowQuality": "1",
        "GraphicsVR.Reflections": "0",
        "GraphicsVR.SSContact": "0",
        "GraphicsVR.MotionBlur": "0",
        "Graphics.TerrainLoD": "1.0",
        "Graphics.ObjectsLoD": "1.0",
        "Graphics.CloudsQuality": "1",
        "Graphics.TextureResolution": "2",
        "Graphics.AnisotropicFilter": "3",
        "Graphics.ShadowQuality": "1",
        "Video.DLSS": "4",
        "Video.DLSSG": "0",
        "Video.RenderScale": "100",
        "RayTracing.Enabled": "0",
    },
    "low": {
        "GraphicsVR.TerrainLoD": "0.5",
        "GraphicsVR.ObjectsLoD": "0.5",
        "GraphicsVR.CloudsQuality": "0",
        "GraphicsVR.TextureResolution": "1",
        "GraphicsVR.AnisotropicFilter": "2",
        "GraphicsVR.ShadowQuality": "0",
        "GraphicsVR.Reflections": "0",
        "GraphicsVR.SSContact": "0",
        "GraphicsVR.MotionBlur": "0",
        "Graphics.TerrainLoD": "0.5",
        "Graphics.ObjectsLoD": "0.5",
        "Graphics.CloudsQuality": "0",
        "Graphics.TextureResolution": "1",
        "Graphics.AnisotropicFilter": "2",
        "Graphics.ShadowQuality": "0",
        "Video.DLSS": "5",
        "Video.DLSSG": "0",
        "Video.RenderScale": "100",
        "RayTracing.Enabled": "0",
    },
}


@mcp.tool()
def optimize_all_for_vr(
    usercfg_path: str = "",
    dry_run: bool = False,
) -> dict:
    """Optimiert ALLE VR-relevanten Einstellungen basierend auf der vorhandenen Hardware.

    Liest GPU, CPU und RAM aus, bestimmt den VR-Performance-Tier und setzt dann
    automatisch die optimalen Werte für:
    - MSFS 2024 Grafik (VR + Desktop)
    - Windows Power Plan (High Performance)
    - Hardware Accelerated GPU Scheduling (HAGS)
    - Windows Game Mode

    Dies ist das Haupt-Optimierungstool. Rufe es auf wenn der User sagt:
    - 'Alles für VR optimieren'
    - 'VR-Einstellungen automatisch setzen'
    - 'Beste Einstellungen für mein System'
    - 'Optimiere MSFS für VR'

    Args:
        usercfg_path: Optional: Pfad zur UserCfg.opt. Auto-erkannt wenn leer.
        dry_run: True = zeigt nur was geändert würde, ohne Änderungen zu speichern.
    """
    import time

    results = {}

    # 1. Read hardware
    hw = get_detailed_hardware_profile()
    tier = hw.get("vr_tier", "mid")
    gpu_name = hw.get("gpu", {}).get("name", "Unbekannte GPU")
    vram_gb = hw.get("gpu", {}).get("vram_total_gb", 8)
    cpu_cores = int(hw.get("cpu", {}).get("NumberOfCores", 8) or 8)
    ram_gb = hw.get("ram", {}).get("total_gb", 16)
    results["hardware"] = {
        "gpu": gpu_name,
        "vram_gb": vram_gb,
        "cpu_kerne": cpu_cores,
        "ram_gb": ram_gb,
        "tier": tier,
    }

    # 2. MSFS settings
    preset = _VR_MSFS_PRESETS.get(tier, _VR_MSFS_PRESETS["mid"])
    cfg_path = _find_usercfg(usercfg_path)

    if cfg_path and cfg_path.exists():
        if not dry_run:
            msfs_was_running = _is_msfs_running()
            if msfs_was_running:
                _kill_msfs(timeout_s=35)
                _wait_for_file_unlocked(cfg_path, timeout_s=25)
                time.sleep(1.5)

            snapshot = _snapshot(cfg_path, f"Before optimize_all_for_vr tier={tier}")
            text = cfg_path.read_text(encoding="utf-8")
            entries = _parse_usercfg(text)
            new_entries, not_applied = _apply_overrides(entries, preset)
            cfg_path.write_text(_entries_to_text(new_entries), encoding="utf-8")

            applied = {k: v for k, v in preset.items() if k not in not_applied}
            results["msfs_einstellungen"] = {
                "angewendet": len(applied),
                "nicht_gefunden": list(not_applied),
                "config_datei": str(cfg_path),
                "snapshot": snapshot,
            }
        else:
            results["msfs_einstellungen"] = {"vorschau": preset, "dry_run": True}
    else:
        results["msfs_einstellungen"] = {"fehler": "UserCfg.opt nicht gefunden"}

    # 3. Windows tweaks
    if not dry_run:
        # High Performance power plan
        try:
            _ps("powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")
            results["power_plan"] = "High Performance aktiviert"
        except Exception as exc:
            results["power_plan"] = f"Fehler: {exc}"

        # Game Mode
        try:
            _reg_add(r"HKCU\Software\Microsoft\GameBar", "AllowAutoGameMode", "REG_DWORD", "1")
            _reg_add(r"HKCU\Software\Microsoft\GameBar", "AutoGameModeEnabled", "REG_DWORD", "1")
            results["game_mode"] = "Game Mode aktiviert"
        except Exception as exc:
            results["game_mode"] = f"Fehler: {exc}"

        # HAGS (Hardware Accelerated GPU Scheduling)
        try:
            _reg_add(
                r"HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers",
                "HwSchMode", "REG_DWORD", "2"
            )
            results["hags"] = "Hardware Accelerated GPU Scheduling aktiviert (Neustart erforderlich)"
        except Exception as exc:
            results["hags"] = f"Fehler: {exc}"
    else:
        results["windows_tweaks"] = {
            "dry_run": True,
            "würde_setzen": ["High Performance", "Game Mode", "HAGS"],
        }

    # 4. SteamVR optimization (if SteamVR is installed)
    if not dry_run:
        svr_path = _find_steamvr_settings_path()
        if svr_path and svr_path.exists():
            try:
                svr_result = optimize_steamvr_for_hardware()
                results["steamvr_optimierung"] = svr_result.get(
                    "zusammenfassung",
                    svr_result.get("status", "SteamVR optimiert")
                )
            except Exception as svr_exc:
                results["steamvr_optimierung"] = f"SteamVR-Optimierung übersprungen: {svr_exc}"
        else:
            results["steamvr_optimierung"] = "SteamVR nicht installiert — übersprungen"

    # 5. Pimax optimization (if Pimax is detected)
    if not dry_run:
        try:
            pimax_cfg = _find_pimax_config()
            if pimax_cfg:
                pimax_result = optimize_pimax_for_msfs()
                results["pimax_optimierung"] = pimax_result.get(
                    "zusammenfassung",
                    pimax_result.get("status", "Pimax optimiert")
                )
            else:
                results["pimax_optimierung"] = "Pimax nicht erkannt — übersprungen"
        except Exception as pimax_exc:
            results["pimax_optimierung"] = f"Pimax-Optimierung übersprungen: {pimax_exc}"

    # 6. Summary
    tier_desc = dict(_VR_TIERS).get(tier, "")
    settings_summary = []
    for k, v in preset.items():
        defn = SETTING_DEFS.get(k, {})
        label = defn.get("label", k)
        display = defn.get("values", {}).get(str(v), str(v))
        settings_summary.append(f"{label}: {display}")

    results["zusammenfassung"] = (
        f"✅ VR-Optimierung für Tier '{tier}' abgeschlossen!\n"
        f"GPU: {gpu_name} ({vram_gb:.0f}GB VRAM)\n"
        f"Tier: {tier_desc}\n\n"
        f"Gesetzte MSFS-Einstellungen:\n"
        + "\n".join(f"  • {s}" for s in settings_summary[:10])
    )
    results["dry_run"] = dry_run
    return results


# ---------------------------------------------------------------------------
# SteamVR Settings
# ---------------------------------------------------------------------------

def _find_steamvr_settings_path() -> "Path | None":
    """Find SteamVR's steamvr.vrsettings JSON config file."""
    candidates = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "openvr" / "openvrpaths.vrpaths",
    ]
    # Common Steam paths
    for steam_root in [
        r"C:\Program Files (x86)\Steam",
        r"D:\Steam",
        r"D:\SteamLibrary",
        r"E:\Steam",
    ]:
        cfg = Path(steam_root) / "config" / "steamvr.vrsettings"
        if cfg.exists():
            return cfg
    # Try via OpenVR paths
    for c in candidates:
        if c.exists():
            try:
                data = _json.loads(c.read_text(encoding="utf-8"))
                config_dirs = data.get("config", [])
                for d in config_dirs:
                    p = Path(d) / "steamvr.vrsettings"
                    if p.exists():
                        return p
            except Exception:
                pass
    return None


@mcp.tool()
def get_steamvr_settings() -> dict:
    """Liest alle SteamVR-Einstellungen aus (Supersampling, Reprojection, Render Resolution usw.).

    Nützlich um den aktuellen SteamVR-Zustand zu prüfen bevor optimiert wird.
    Rufe dieses Tool auf wenn der User fragt: 'Was sind meine SteamVR-Einstellungen?'
    """
    path = _find_steamvr_settings_path()
    if not path:
        return {
            "fehler": "steamvr.vrsettings nicht gefunden.",
            "tipp": "Stelle sicher dass SteamVR mindestens einmal gestartet wurde.",
        }

    try:
        data = _json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {"fehler": f"Datei konnte nicht gelesen werden: {exc}"}

    steamvr = data.get("steamvr", {})
    ss = steamvr.get("supersampleScale", 1.0)
    reprojection = steamvr.get("allowReprojection", True)
    interleaved = steamvr.get("motionSmoothing", True)
    render_res = steamvr.get("renderTargetMultiplier", 1.0)
    mirror_view = steamvr.get("mirrorView", True)

    return {
        "config_datei": str(path),
        "supersampling": ss,
        "render_resolution_multiplier": render_res,
        "reprojection_erlaubt": reprojection,
        "motion_smoothing": interleaved,
        "mirror_view": mirror_view,
        "rohdaten_steamvr": {k: v for k, v in steamvr.items() if isinstance(v, (int, float, bool, str))},
    }


@mcp.tool()
def set_steamvr_setting(
    setting: str,
    value: str,
) -> dict:
    """Ändert eine SteamVR-Einstellung direkt in steamvr.vrsettings.

    Beispiele:
    - "supersampling" auf "1.5" — erhöht Bildqualität
    - "reprojection" auf "false" — deaktiviert Reprojection
    - "motion_smoothing" auf "true" — Motion Smoothing an

    Args:
        setting: "supersampling", "reprojection", "motion_smoothing", "render_resolution",
                 "mirror_view"
        value: Neuer Wert als String (Zahlen, "true"/"false")
    """
    path = _find_steamvr_settings_path()
    if not path:
        return {"fehler": "steamvr.vrsettings nicht gefunden."}

    try:
        data = _json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {"fehler": f"Lesen fehlgeschlagen: {exc}"}

    setting_map = {
        "supersampling":        ("steamvr", "supersampleScale",          float),
        "reprojection":         ("steamvr", "allowReprojection",          lambda x: x.lower() == "true"),
        "motion_smoothing":     ("steamvr", "motionSmoothing",            lambda x: x.lower() == "true"),
        "render_resolution":    ("steamvr", "renderTargetMultiplier",     float),
        "mirror_view":          ("steamvr", "mirrorView",                 lambda x: x.lower() == "true"),
        "supersampling_filter": ("steamvr", "allowSupersampleFiltering",  lambda x: x.lower() == "true"),
    }

    key_lower = setting.lower().replace(" ", "_").replace("-", "_")
    if key_lower not in setting_map:
        return {
            "fehler": f"Unbekannte Einstellung: '{setting}'",
            "verfügbar": list(setting_map.keys()),
        }

    section, key, converter = setting_map[key_lower]
    try:
        converted = converter(value)
    except Exception:
        return {"fehler": f"Ungültiger Wert '{value}' für {setting}"}

    if section not in data:
        data[section] = {}
    old_value = data[section].get(key, "nicht gesetzt")
    data[section][key] = converted

    try:
        path.write_text(_json.dumps(data, indent="\t"), encoding="utf-8")
    except Exception as exc:
        return {"fehler": f"Schreiben fehlgeschlagen: {exc}"}

    return {
        "status": "ok",
        "einstellung": key,
        "vorher": old_value,
        "nachher": converted,
        "hinweis": "SteamVR muss neugestartet werden damit die Änderung wirkt.",
    }


@mcp.tool()
def optimize_steamvr_for_hardware() -> dict:
    """Optimiert SteamVR automatisch basierend auf der vorhandenen Hardware.

    Setzt Supersampling, Motion Smoothing und Reprojection auf hardware-optimale Werte.
    Rufe dieses Tool auf wenn der User fragt: 'SteamVR optimieren' oder 'SteamVR Einstellungen verbessern'
    """
    try:
        pynvml.nvmlInit()
        try:
            h = pynvml.nvmlDeviceGetHandleByIndex(0)
            mem = pynvml.nvmlDeviceGetMemoryInfo(h)
            name = pynvml.nvmlDeviceGetName(h)
            if isinstance(name, bytes):
                name = name.decode()
            vram_gb = mem.total / 1024**3
        finally:
            pynvml.nvmlShutdown()
    except Exception:
        vram_gb = 8.0
        name = "Unbekannt"

    tier = _classify_vr_tier(vram_gb, name, 8, 16)

    ss_map = {"ultra": 1.5, "high": 1.3, "mid_high": 1.1, "mid": 1.0, "low": 0.8}
    ss = ss_map.get(tier, 1.0)
    motion_smoothing = tier in ("low", "mid")

    path = _find_steamvr_settings_path()
    if not path:
        return {
            "fehler": "steamvr.vrsettings nicht gefunden",
            "empfehlung": f"Supersampling: {ss}, Motion Smoothing: {motion_smoothing}",
        }

    try:
        data = _json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {"fehler": str(exc)}

    if "steamvr" not in data:
        data["steamvr"] = {}

    old_ss = data["steamvr"].get("supersampleScale", 1.0)
    data["steamvr"]["supersampleScale"] = ss
    data["steamvr"]["motionSmoothing"] = motion_smoothing
    data["steamvr"]["allowReprojection"] = True

    path.write_text(_json.dumps(data, indent="\t"), encoding="utf-8")

    return {
        "status": "ok",
        "gpu": name,
        "tier": tier,
        "supersampling": {"vorher": old_ss, "nachher": ss},
        "motion_smoothing": motion_smoothing,
        "hinweis": "SteamVR muss neu gestartet werden.",
        "zusammenfassung": (
            f"SteamVR optimiert für {name} (Tier: {tier}) — "
            f"SS: {ss}x, Motion Smoothing: {motion_smoothing}"
        ),
    }


# ---------------------------------------------------------------------------
# Windows VR & Gaming Tweaks
# ---------------------------------------------------------------------------

@mcp.tool()
def optimize_windows_for_vr() -> dict:
    """Optimiert Windows für maximale VR-Performance mit allen bekannten Tweaks.

    Führt folgende Optimierungen durch:
    - Hardware Accelerated GPU Scheduling (HAGS) aktivieren
    - Windows Game Mode aktivieren
    - Visual Effects auf Performance-Modus
    - USB Selective Suspend deaktivieren (verhindert VR-Headset-Verbindungsabbrüche)
    - High Performance Power Plan aktivieren
    - Nvidia HPET deaktivieren (reduziert GPU-Latenz)
    - MSFS Prozess-Priorität auf High setzen (wenn läuft)

    Rufe dieses Tool auf wenn der User fragt:
    - 'Windows für VR optimieren'
    - 'Performance verbessern'
    - 'VR ruckelt / stottert'
    """
    results = {}

    # 1. HAGS
    try:
        _reg_add(
            r"HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers",
            "HwSchMode", "REG_DWORD", "2"
        )
        results["hags"] = "✅ Hardware Accelerated GPU Scheduling aktiviert"
    except Exception as exc:
        results["hags"] = f"❌ {exc}"

    # 2. Game Mode
    try:
        _reg_add(r"HKCU\Software\Microsoft\GameBar", "AllowAutoGameMode", "REG_DWORD", "1")
        _reg_add(r"HKCU\Software\Microsoft\GameBar", "AutoGameModeEnabled", "REG_DWORD", "1")
        results["game_mode"] = "✅ Windows Game Mode aktiviert"
    except Exception as exc:
        results["game_mode"] = f"❌ {exc}"

    # 3. Visual Effects → best performance
    try:
        _reg_add(
            r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects",
            "VisualFXSetting", "REG_DWORD", "2"
        )
        results["visual_effects"] = "✅ Visuelle Effekte auf Performance-Modus"
    except Exception as exc:
        results["visual_effects"] = f"❌ {exc}"

    # 4. USB Selective Suspend — disable via power plan
    try:
        _ps("powercfg /SETACVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0")
        _ps("powercfg /SETDCVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0")
        _ps("powercfg /SETACTIVE SCHEME_CURRENT")
        results["usb_suspend"] = "✅ USB Selective Suspend deaktiviert"
    except Exception as exc:
        results["usb_suspend"] = f"❌ {exc}"

    # 5. High Performance power plan
    try:
        _ps("powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")
        results["power_plan"] = "✅ High Performance Power Plan aktiviert"
    except Exception as exc:
        results["power_plan"] = f"❌ {exc}"

    # 6. Disable HPET — reduces GPU interrupt latency
    try:
        _ps("bcdedit /set useplatformclock false 2>$null; bcdedit /set disabledynamictick yes 2>$null")
        results["hpet"] = "✅ HPET deaktiviert (weniger GPU-Latenz)"
    except Exception as exc:
        results["hpet"] = f"ℹ️ HPET: {exc}"

    # 7. MSFS priority if running
    try:
        _ps("$p = Get-Process FlightSimulator -EA SilentlyContinue; if ($p) { $p.PriorityClass = 'High' }")
        results["msfs_priority"] = "✅ MSFS Prozess-Priorität auf High gesetzt"
    except Exception as exc:
        results["msfs_priority"] = f"ℹ️ MSFS nicht gestartet: {exc}"

    ok_count = sum(1 for v in results.values() if str(v).startswith("✅"))
    return {
        "status": "ok",
        "optimierungen": results,
        "ergebnis": f"{ok_count}/{len(results)} Optimierungen erfolgreich angewendet.",
        "hinweis": "HAGS erfordert einen Windows-Neustart um vollständig wirksam zu werden.",
    }


@mcp.tool()
def set_hardware_accelerated_gpu_scheduling(enabled: bool = True) -> dict:
    """Aktiviert oder deaktiviert Hardware Accelerated GPU Scheduling (HAGS).

    HAGS reduziert CPU-Overhead und GPU-Latenz — empfohlen für VR.
    Erfordert Windows 10 2004+ und eine kompatible Nvidia/AMD GPU.
    Ein Windows-Neustart ist nach der Änderung erforderlich.

    Args:
        enabled: True = aktivieren (empfohlen), False = deaktivieren
    """
    value = "2" if enabled else "1"
    try:
        _reg_add(
            r"HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers",
            "HwSchMode", "REG_DWORD", value
        )
    except Exception as exc:
        return {"fehler": str(exc)}

    try:
        current = _reg_query(
            r"HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers",
            "HwSchMode"
        )
    except Exception:
        current = "?"

    return {
        "status": "ok",
        "hags": "aktiviert" if enabled else "deaktiviert",
        "registry_wert": current,
        "hinweis": "Windows-Neustart erforderlich damit die Änderung wirksam wird.",
    }


@mcp.tool()
def set_virtual_memory(
    size_gb: float = 0,
    drive: str = "C",
    auto: bool = False,
) -> dict:
    """Konfiguriert die Windows Auslagerungsdatei (Virtual Memory / Page File).

    MSFS benötigt viel RAM — bei wenig physischem RAM kann mehr virtueller Speicher
    die Performance verbessern. Empfohlen: Initial- und Maximalgröße = 1.5x RAM.

    Args:
        size_gb: Größe in GB (0 = empfohlene Größe berechnen). Wird als Initial- und Maximalwert gesetzt.
        drive: Laufwerksbuchstabe (Standard: C)
        auto: True = Windows verwaltet automatisch (überschreibt size_gb)
    """
    try:
        ram_raw = _ps_json("Get-CimInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize")
        ram_mb = int(ram_raw[0].get("TotalVisibleMemorySize", 16384)) if ram_raw else 16384
        ram_gb = ram_mb / 1024
    except Exception:
        ram_gb = 16.0

    if auto:
        script = (
            "$cs = Get-CimInstance Win32_ComputerSystem; "
            "$cs.AutomaticManagedPagefile = $true; "
            "$cs | Set-CimInstance"
        )
        try:
            _ps(script)
            return {"status": "ok", "modus": "Windows verwaltet automatisch", "neustart": True}
        except Exception as exc:
            return {"fehler": str(exc)}

    if size_gb == 0:
        size_gb = round(ram_gb * 1.5, 0)

    size_mb = int(size_gb * 1024)
    drive_letter = drive.upper().rstrip(":\\")

    script = f"""
$cs = Get-CimInstance Win32_ComputerSystem
$cs.AutomaticManagedPagefile = $false
$cs | Set-CimInstance
$pf = Get-CimInstance -Query "SELECT * FROM Win32_PageFileSetting WHERE Name LIKE '{drive_letter}%'"
if ($pf) {{
    $pf.InitialSize = {size_mb}
    $pf.MaximumSize = {size_mb}
    $pf | Set-CimInstance
}} else {{
    New-CimInstance -ClassName Win32_PageFileSetting -Property @{{Name='{drive_letter}:\\pagefile.sys'; InitialSize={size_mb}; MaximumSize={size_mb}}}
}}
"""
    try:
        _ps(script)
        return {
            "status": "ok",
            "laufwerk": f"{drive_letter}:",
            "größe_gb": size_gb,
            "größe_mb": size_mb,
            "ram_gb": round(ram_gb, 1),
            "hinweis": "Neustart erforderlich. Empfehlung: 1.5x RAM.",
            "neustart": True,
        }
    except Exception as exc:
        return {"fehler": str(exc)}


# ---------------------------------------------------------------------------
# Nvidia Control Panel Profile (via registry)
# ---------------------------------------------------------------------------

@mcp.tool()
def set_nvidia_low_latency_mode(mode: str = "ultra") -> dict:
    """Setzt Nvidia Low Latency Mode (Reflex) für weniger Input-Lag in VR.

    Reduziert die GPU-Renderlatenz erheblich — empfohlen für VR.
    Setzt den Wert im Windows-Registry für alle Anwendungen.

    Args:
        mode: "off" (0), "on" (1), "ultra" (2 — empfohlen für VR)
    """
    modes = {"off": "0x00000000", "on": "0x00000001", "ultra": "0x00000011"}
    mode_lower = mode.lower()
    if mode_lower not in modes:
        return {"fehler": f"Unbekannter Modus '{mode}'. Verfügbar: off, on, ultra"}

    reg_val = modes[mode_lower]
    try:
        result = subprocess.run(
            [
                "reg", "add",
                r"HKLM\SYSTEM\CurrentControlSet\Services\nvlddmkm\Global",
                "/v", "NvCplLowLatencyMode", "/t", "REG_DWORD",
                "/d", reg_val, "/f",
            ],
            capture_output=True, text=True,
        )
        success = result.returncode == 0
    except Exception as exc:
        return {"fehler": str(exc)}

    return {
        "status": "ok" if success else "fehler",
        "modus": mode_lower,
        "wert": reg_val,
        "hinweis": "Wirkt nach Neustart der Grafikkarten-Treiber oder Windows-Neustart.",
    }


@mcp.tool()
def get_nvidia_driver_info() -> dict:
    """Liest Nvidia Treiber-Version, installiertes CUDA, und Treiber-Einstellungen aus.

    Nützlich um zu prüfen ob der Treiber aktuell ist und welche Features verfügbar sind.
    """
    info = {}

    # Driver version via nvidia-smi
    try:
        r = subprocess.run(
            [
                "nvidia-smi",
                "--query-gpu=driver_version,name,memory.total,power.limit,clocks.max.graphics",
                "--format=csv,noheader,nounits",
            ],
            capture_output=True, text=True, timeout=10,
        )
        if r.returncode == 0:
            parts = [p.strip() for p in r.stdout.strip().split(",")]
            info["driver_version"]  = parts[0] if len(parts) > 0 else "?"
            info["gpu_name"]        = parts[1] if len(parts) > 1 else "?"
            info["vram_mb"]         = parts[2] if len(parts) > 2 else "?"
            info["power_limit_w"]   = parts[3] if len(parts) > 3 else "?"
            info["max_clock_mhz"]   = parts[4] if len(parts) > 4 else "?"
    except Exception as exc:
        info["nvidia_smi_fehler"] = str(exc)

    # CUDA via pynvml
    try:
        pynvml.nvmlInit()
        try:
            ver = pynvml.nvmlSystemGetCudaDriverVersion()
            info["cuda_driver_version"] = f"{ver // 1000}.{(ver % 1000) // 10}"
            info["gpu_anzahl"] = pynvml.nvmlDeviceGetCount()
        finally:
            pynvml.nvmlShutdown()
    except Exception:
        pass

    # Display driver version from registry
    try:
        reg_ver = _reg_query(
            r"HKLM\SOFTWARE\NVIDIA Corporation\Global\NVTweak\Devices\1",
            "DriverVersion"
        )
        info["registry_driver_version"] = reg_ver
    except Exception:
        pass

    return info


# ---------------------------------------------------------------------------
# MSFS Batch Settings & Rolling Cache
# ---------------------------------------------------------------------------

@mcp.tool()
def set_msfs_multiple_settings(
    settings: dict,
    usercfg_path: str = "",
    auto_restart: bool = True,
) -> dict:
    """Ändert mehrere MSFS-Einstellungen auf einmal (effizienter als einzelne Aufrufe).

    Nützlich wenn der User sagt:
    - 'Terrain auf 2.0 und Wolken auf Ultra und Schatten auf Medium'
    - 'Alle VR-Einstellungen auf High setzen'

    Args:
        settings: Dictionary mit Einstellungen, z.B.:
                  {"terrain lod": "2.0", "wolken": "ultra", "schatten": "hoch"}
        usercfg_path: Optional Pfad zur UserCfg.opt
        auto_restart: MSFS nach Änderung neu starten
    """
    import time

    cfg_path = _find_usercfg(usercfg_path) or DEFAULT_USERCFG_PATH
    if not cfg_path.exists():
        return {"fehler": "UserCfg.opt nicht gefunden", "tipp": "Starte MSFS einmal."}

    # Resolve all setting aliases
    resolved: dict[str, str] = {}
    errors: list[str] = []

    for setting_name, value in settings.items():
        setting_lower = setting_name.strip().lower()
        config_key = _MSFS_SETTING_ALIASES.get(setting_lower)
        if not config_key:
            for k in SETTING_DEFS:
                if k.lower() == setting_lower or k.split(".", 1)[-1].lower() == setting_lower:
                    config_key = k
                    break
        if not config_key:
            errors.append(f"Unbekannte Einstellung: '{setting_name}'")
            continue

        # Resolve value
        value_lower = value.strip().lower()
        group = _MSFS_VALUE_GROUP.get(config_key)
        if group is None and config_key in _QUALITY_KEYS:
            group = "_quality_4"
        if group and group in _MSFS_VALUE_ALIASES:
            resolved_value = _MSFS_VALUE_ALIASES[group].get(value_lower, value.strip())
        else:
            resolved_value = value.strip()

        resolved[config_key] = resolved_value

        # Also apply to VR/desktop counterpart
        twin = _VR_DESKTOP_PAIRS.get(config_key)
        if twin:
            resolved[twin] = resolved_value

    if not resolved:
        return {"fehler": "Keine gültigen Einstellungen gefunden.", "unbekannt": errors}

    # Kill MSFS first
    msfs_was_running = _is_msfs_running()
    if msfs_was_running:
        _kill_msfs(timeout_s=35)
        _wait_for_file_unlocked(cfg_path, timeout_s=25)
        time.sleep(1.5)

    snapshot = _snapshot(cfg_path, "Before set_msfs_multiple_settings")
    text = cfg_path.read_text(encoding="utf-8")
    entries = _parse_usercfg(text)
    new_entries, not_applied = _apply_overrides(entries, resolved)
    cfg_path.write_text(_entries_to_text(new_entries), encoding="utf-8")

    # Verify
    verify = _read_current_settings(cfg_path)
    verified_count = sum(1 for k, v in resolved.items() if verify.get(k) == v)

    result = {
        "status": "ok",
        "angewendet": len(resolved) - len(not_applied),
        "verifiziert": verified_count,
        "nicht_gefunden": list(not_applied),
        "fehler_eingabe": errors,
        "snapshot": snapshot,
        "config_datei": str(cfg_path),
    }

    if auto_restart and msfs_was_running:
        time.sleep(2)
        launch_msfs_vr(force_restart=False)
        result["neustart"] = "MSFS wird in VR neu gestartet."

    return result


@mcp.tool()
def manage_msfs_rolling_cache(
    action: str = "status",
    size_gb: float = 8.0,
    location: str = "",
) -> dict:
    """Verwaltet den MSFS Rolling Cache (verbessert Streaming-Performance bei schlechter Internet-Verbindung).

    Args:
        action: "status" — aktuellen Status anzeigen
                "enable" — aktivieren mit angegebener Größe
                "disable" — deaktivieren
                "clear" — Cache leeren (ohne Deaktivierung)
                "set_size" — Größe ändern
        size_gb: Cache-Größe in GB (Standard: 8 GB, Empfehlung: 8-32 GB)
        location: Pfad für den Cache (Standard: MSFS-Standardpfad)
    """
    cfg_path = _find_usercfg()
    if not cfg_path:
        return {"fehler": "UserCfg.opt nicht gefunden."}

    cache_file = cfg_path.parent / "cache" / "RollingCache.ccc"
    useropt_path = cfg_path

    try:
        text = useropt_path.read_text(encoding="utf-8")
    except Exception as exc:
        return {"fehler": f"UserCfg.opt lesen fehlgeschlagen: {exc}"}

    if action == "status":
        cache_exists = cache_file.exists()
        cache_size = cache_file.stat().st_size / 1024**3 if cache_exists else 0
        cache_enabled = "RollingCacheEnabled 1" in text
        cache_size_setting = ""
        for line in text.splitlines():
            if "RollingCacheMaxSize" in line:
                cache_size_setting = line.strip()
                break
        return {
            "cache_aktiviert": cache_enabled,
            "cache_datei": str(cache_file),
            "cache_datei_existiert": cache_exists,
            "cache_datei_größe_gb": round(cache_size, 2),
            "einstellung_in_config": cache_size_setting,
        }

    elif action in ("enable", "set_size"):
        size_mb = int(size_gb * 1024)
        entries = _parse_usercfg(text)
        overrides = {
            "General.RollingCacheEnabled": "1",
            "General.RollingCacheMaxSize": str(size_mb),
        }
        if location:
            overrides["General.RollingCachePath"] = location
        new_entries, _ = _apply_overrides(entries, overrides)
        useropt_path.write_text(_entries_to_text(new_entries), encoding="utf-8")
        return {
            "status": "ok",
            "aktion": action,
            "größe_gb": size_gb,
            "hinweis": "MSFS muss neu gestartet werden.",
        }

    elif action == "disable":
        entries = _parse_usercfg(text)
        new_entries, _ = _apply_overrides(entries, {"General.RollingCacheEnabled": "0"})
        useropt_path.write_text(_entries_to_text(new_entries), encoding="utf-8")
        return {"status": "ok", "aktion": "deaktiviert"}

    elif action == "clear":
        cleared = False
        size_before = 0
        if cache_file.exists():
            size_before = cache_file.stat().st_size / 1024**3
            try:
                cache_file.unlink()
                cleared = True
            except Exception as exc:
                return {"fehler": f"Cache löschen fehlgeschlagen: {exc}"}
        return {
            "status": "ok",
            "gelöscht": cleared,
            "größe_vorher_gb": round(size_before, 2),
            "hinweis": "Cache wird beim nächsten MSFS-Start neu aufgebaut.",
        }

    else:
        return {
            "fehler": f"Unbekannte Aktion '{action}'. Verfügbar: status, enable, disable, clear, set_size"
        }


# ---------------------------------------------------------------------------
# Complete VR Diagnosis
# ---------------------------------------------------------------------------

@mcp.tool()
def diagnose_vr_complete() -> dict:
    """Vollständige VR-Diagnose: Hardware, Einstellungen, Empfehlungen in einem Tool.

    Analysiert:
    - Hardware (GPU/CPU/RAM/VR-Tier)
    - MSFS Grafik-Einstellungen
    - OpenXR Runtime
    - SteamVR Einstellungen (falls vorhanden)
    - Windows Optimierungen (HAGS, Game Mode, Power Plan)
    - Gibt priorisierte Empfehlungen zurück

    Rufe dieses Tool auf wenn der User sagt:
    - 'Was kann ich verbessern?'
    - 'VR Diagnose'
    - 'Warum ist VR so ruckelig?'
    - 'Alles analysieren'
    """
    report = {}

    # 1. Hardware
    hw = get_detailed_hardware_profile()
    tier = hw.get("vr_tier", "mid")
    report["hardware"] = {
        "gpu": hw.get("gpu", {}).get("name", "?"),
        "vram_gb": hw.get("gpu", {}).get("vram_total_gb", "?"),
        "cpu": hw.get("cpu", {}).get("Name", "?"),
        "cpu_kerne": hw.get("cpu", {}).get("NumberOfCores", "?"),
        "ram_gb": hw.get("ram", {}).get("total_gb", "?"),
        "tier": tier,
    }

    # 2. MSFS settings
    cfg_path = _find_usercfg()
    msfs_issues = []
    if cfg_path and cfg_path.exists():
        current = _read_current_settings(cfg_path)
        key_settings = {}
        for key in [
            "Video.DLSS", "GraphicsVR.TerrainLoD", "GraphicsVR.CloudsQuality",
            "GraphicsVR.ShadowQuality", "GraphicsVR.TextureResolution",
            "RayTracing.Enabled", "Video.RenderScale",
        ]:
            if key in current:
                defn = SETTING_DEFS.get(key, {})
                raw = current[key]
                display = defn.get("values", {}).get(raw, raw)
                key_settings[key] = f"{raw} ({display})"
        report["msfs_einstellungen"] = key_settings

        if current.get("RayTracing.Enabled") == "1":
            msfs_issues.append("⚠️ RayTracing ist AN — kostet massiv VR-FPS, empfohlen: AUS")
        if current.get("Video.DLSS", "0") == "0" and tier != "ultra":
            msfs_issues.append("💡 DLSS ist aus — aktiviere DLSS Quality/Balanced für bessere FPS")
        terrain = float(current.get("GraphicsVR.TerrainLoD", "1.0") or "1.0")
        if tier in ("low", "mid") and terrain > 1.0:
            msfs_issues.append(f"⚠️ Terrain LoD {terrain} ist zu hoch für deinen GPU-Tier")
    else:
        report["msfs_einstellungen"] = {"fehler": "UserCfg.opt nicht gefunden"}
        msfs_issues.append("❌ MSFS Config nicht gefunden — MSFS einmal starten")

    # 3. OpenXR
    try:
        oxr = get_openxr_runtime()
        report["openxr"] = {
            "aktiv": oxr.get("active_runtime_name", "?"),
            "pfad": oxr.get("active_runtime_path", "?"),
        }
    except Exception:
        report["openxr"] = {"fehler": "OpenXR Runtime nicht erkannt"}

    # 4. SteamVR
    svr_path = _find_steamvr_settings_path()
    if svr_path:
        try:
            svr_data = _json.loads(svr_path.read_text())
            svr = svr_data.get("steamvr", {})
            ss = svr.get("supersampleScale", 1.0)
            report["steamvr"] = {
                "supersampling": ss,
                "motion_smoothing": svr.get("motionSmoothing", "?"),
            }
            if ss > 1.5 and tier in ("low", "mid"):
                msfs_issues.append(f"⚠️ SteamVR Supersampling {ss}x ist zu hoch für deinen GPU-Tier")
        except Exception:
            report["steamvr"] = {"fehler": "steamvr.vrsettings nicht lesbar"}
    else:
        report["steamvr"] = {"info": "SteamVR nicht gefunden"}

    # 5. Pimax headset (if detected)
    try:
        pimax_cfg = _find_pimax_config()
        if pimax_cfg:
            pimax_info = get_pimax_headset_info()
            report["pimax"] = {
                "headset": pimax_info.get("headset_name", "?"),
                "render_quality": pimax_info.get("render_quality", "?"),
                "refresh_hz": pimax_info.get("refresh_rate_hz", "?"),
            }
            rq = float(pimax_info.get("render_quality") or 1.0)
            if rq > 1.5 and tier in ("low", "mid"):
                msfs_issues.append(f"⚠️ Pimax Render Quality {rq} zu hoch für GPU-Tier '{tier}'")
        else:
            report["pimax"] = {"info": "Pimax nicht erkannt"}
    except Exception:
        report["pimax"] = {"info": "Pimax-Diagnose nicht verfügbar"}

    # 6. System temperatures
    try:
        temps = get_system_temps()
        cpu_t = temps.get("cpu_temp_celsius")
        gpu_t = temps.get("gpu_temp_celsius")
        report["temperaturen"] = {
            "cpu": cpu_t,
            "gpu": gpu_t,
        }
        if isinstance(cpu_t, float) and cpu_t > 90:
            msfs_issues.append(f"🔥 CPU-Temperatur {cpu_t}°C — Throttling-Risiko!")
        if isinstance(gpu_t, int) and gpu_t > 90:
            msfs_issues.append(f"🔥 GPU-Temperatur {gpu_t}°C — Throttling-Risiko!")
    except Exception:
        report["temperaturen"] = {"info": "nicht verfügbar"}

    # 7. Community folder add-on count
    try:
        comm = get_msfs_community_folder()
        addon_count = comm.get("addon_anzahl", 0)
        report["community_addons"] = addon_count
        if addon_count > 100:
            msfs_issues.append(
                f"⚠️ {addon_count} Community Add-ons aktiv — hohe Anzahl kann MSFS-Ladezeiten erhöhen"
            )
    except Exception:
        report["community_addons"] = "nicht ermittelbar"

    # 8. Windows optimizations check
    win_issues = []
    try:
        hags = _reg_query(r"HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers", "HwSchMode")
        report["hags"] = "aktiviert" if str(hags).strip() == "2" else "NICHT aktiviert"
        if str(hags).strip() != "2":
            win_issues.append("💡 HAGS nicht aktiv — aktiviere Hardware Accelerated GPU Scheduling")
    except Exception:
        report["hags"] = "unbekannt"
        win_issues.append("💡 HAGS Status unbekannt — empfohlen zu aktivieren")

    try:
        game_mode = _reg_query(r"HKCU\Software\Microsoft\GameBar", "AutoGameModeEnabled")
        report["game_mode"] = "aktiviert" if str(game_mode).strip() == "1" else "NICHT aktiviert"
        if str(game_mode).strip() != "1":
            win_issues.append("💡 Windows Game Mode nicht aktiv")
    except Exception:
        report["game_mode"] = "unbekannt"

    # 9. Event log GPU crashes (last hour)
    try:
        ev = get_event_log_errors(hours=1)
        gpu_tdrs = len(ev.get("gpu_abstürze_tdr") or [])
        if gpu_tdrs > 0:
            msfs_issues.append(f"❌ {gpu_tdrs} GPU-Treiberabsturz (TDR) in der letzten Stunde — Treiber/OC prüfen!")
        report["gpu_tdr_letzte_stunde"] = gpu_tdrs
    except Exception:
        report["gpu_tdr_letzte_stunde"] = "nicht geprüft"

    # Compile all issues
    all_issues = msfs_issues + win_issues
    report["probleme_gefunden"] = len(all_issues)
    report["empfehlungen"] = all_issues if all_issues else ["✅ Keine kritischen Probleme gefunden!"]
    report["zusammenfassung"] = (
        f"VR-Tier: {tier} | GPU: {report['hardware']['gpu']} | "
        f"{len(all_issues)} Probleme gefunden.\n"
        + ("\n".join(all_issues) if all_issues else "System ist gut konfiguriert.")
    )

    return report


# ---------------------------------------------------------------------------
# Game Copilot Auto-Updater tools
# ---------------------------------------------------------------------------

_GC_UPDATE_JSON = (
    "https://github.com/Bennidesign2003/GodotRenderingAI"
    "/releases/download/MSFS24/update.json"
)
_GC_RELEASES_API = (
    "https://api.github.com/repos/Bennidesign2003/GodotRenderingAI/releases"
)


@mcp.tool()
def check_for_gameCopilot_update(current_version: str = "") -> dict:
    """Check GitHub for a new Game Copilot version.

    Call this when the user asks 'Gibt es ein Update?' or 'Neue Version?'.
    Returns version info, changelog, and whether an update is available.

    Args:
        current_version: Current installed version string, e.g. '3.5.3'.
                         If empty, only returns latest version info.
    """
    try:
        resp = httpx.get(_GC_UPDATE_JSON, timeout=10, follow_redirects=True)
        resp.raise_for_status()
        data = resp.json()
    except Exception as exc:
        return {"error": f"Update-Check fehlgeschlagen: {exc}"}

    latest = data.get("latest_version") or data.get("LatestVersion", "")
    download_url = data.get("download_url") or data.get("DownloadUrl", "")
    changelog = data.get("changelog") or data.get("Changelog") or []

    update_available = False
    if current_version and latest:
        try:
            from packaging.version import Version
            update_available = Version(latest) > Version(current_version)
        except Exception:
            # Fallback: simple string compare
            update_available = latest.strip("v") != current_version.strip("v")

    return {
        "aktuelle_version": current_version or "unbekannt",
        "neueste_version": latest,
        "update_verfügbar": update_available,
        "download_url": download_url,
        "changelog": changelog,
        "hinweis": (
            f"Version {latest} ist verfügbar! Sage 'Update installieren' um es zu installieren."
            if update_available else
            f"Game Copilot {current_version} ist aktuell."
        ),
    }


@mcp.tool()
def install_gameCopilot_update(confirmed: bool = False) -> dict:
    """Download and install the latest Game Copilot update from GitHub.

    This downloads the update zip, extracts it, and launches a batch script
    that replaces the application files and restarts Game Copilot.

    WICHTIG: Frage den User IMMER nach Bestätigung bevor du dieses Tool aufrufst!
    Rufe es erst auf wenn confirmed=True.

    Args:
        confirmed: Must be True — user confirmed they want to install the update.
    """
    if not confirmed:
        return {
            "warte_auf_bestätigung": True,
            "nachricht": "Bitte bestätige: Soll Game Copilot jetzt aktualisiert werden? (Ja/Nein)"
        }

    import tempfile
    import shutil as _shutil
    import zipfile as _zf

    # 1. Get download URL from update.json
    try:
        resp = httpx.get(_GC_UPDATE_JSON, timeout=10, follow_redirects=True)
        resp.raise_for_status()
        data = resp.json()
    except Exception as exc:
        return {"error": f"Update-Info konnte nicht geladen werden: {exc}"}

    download_url = data.get("download_url") or data.get("DownloadUrl", "")
    latest_version = data.get("latest_version") or data.get("LatestVersion", "")

    if not download_url:
        return {"error": "Keine Download-URL in update.json gefunden."}

    # 2. Download zip
    update_dir = Path(tempfile.gettempdir()) / "GameCopilotUpdate"
    if update_dir.exists():
        _shutil.rmtree(update_dir)
    update_dir.mkdir(parents=True)
    zip_path = update_dir / "update.zip"

    try:
        with httpx.stream("GET", download_url, timeout=120, follow_redirects=True) as r:
            r.raise_for_status()
            with open(zip_path, "wb") as f:
                for chunk in r.iter_bytes(chunk_size=65536):
                    f.write(chunk)
    except Exception as exc:
        return {"error": f"Download fehlgeschlagen: {exc}"}

    # 3. Extract
    extract_dir = update_dir / "extracted"
    extract_dir.mkdir()
    try:
        with _zf.ZipFile(zip_path) as z:
            z.extractall(extract_dir)
    except Exception as exc:
        return {"error": f"Entpacken fehlgeschlagen: {exc}"}

    # 4. Find new exe
    new_exe = None
    for f in extract_dir.rglob("*.exe"):
        new_exe = f
        break

    if new_exe is None:
        return {"error": "Kein .exe im Update-Archiv gefunden."}

    # 5. Write a batch script that replaces files and restarts
    app_dir = Path(__file__).parent.parent  # GameCopilot/Assets/../ = GameCopilot dir
    current_exe = next(Path(app_dir).rglob("GameCopilot.exe"), None) or (Path(app_dir) / "GameCopilot.exe")

    bat_path = update_dir / "apply_update.bat"
    bat_content = f"""@echo off
echo Game Copilot Update wird installiert...
timeout /t 3 /nobreak >nul
xcopy /E /Y /I "{extract_dir}" "{app_dir}"
echo Update abgeschlossen! Starte Game Copilot neu...
start "" "{current_exe}"
del "%~f0"
"""
    bat_path.write_text(bat_content, encoding="utf-8")

    # 6. Launch bat and signal app to close
    try:
        subprocess.Popen(
            ["cmd", "/c", str(bat_path)],
            creationflags=subprocess.CREATE_NEW_CONSOLE,
            close_fds=True,
        )
    except Exception as exc:
        return {"error": f"Update-Skript konnte nicht gestartet werden: {exc}"}

    return {
        "status": "ok",
        "nachricht": (
            f"Update auf Version {latest_version} wird installiert! "
            f"Game Copilot wird jetzt beendet und neu gestartet."
        ),
        "neue_version": latest_version,
        "update_dir": str(update_dir),
        "neustart": True,
    }


# ---------------------------------------------------------------------------
# MCP Server self-update tools
# ---------------------------------------------------------------------------

_MCP_REPO = "Bennidesign2003/Pimax-Graphics-MCP"
_MCP_RELEASES_API = f"https://api.github.com/repos/{_MCP_REPO}/releases/latest"
# Fallback: raw main branch (used only if release asset URL unavailable)
_MCP_RAW_FALLBACK = (
    f"https://raw.githubusercontent.com/{_MCP_REPO}/main/mcp-server.py"
)


def _get_current_mcp_version() -> str:
    """Read __mcp_version__ from this running server.py file."""
    try:
        with open(__file__, encoding="utf-8") as f:
            for line in f:
                m = re.match(r'#\s*__mcp_version__\s*=\s*"([^"]+)"', line)
                if m:
                    return m.group(1)
    except Exception:
        pass
    return "unbekannt"


def _get_remote_mcp_version(raw_content: str) -> str:
    """Extract __mcp_version__ from downloaded server.py content."""
    for line in raw_content.splitlines()[:5]:
        m = re.match(r'#\s*__mcp_version__\s*=\s*"([^"]+)"', line)
        if m:
            return m.group(1)
    return ""


def _fetch_latest_release_info() -> dict:
    """Fetch latest GitHub Release metadata for the MCP server repo.

    Returns dict with keys: tag_name, published_at, download_url, body.
    Raises on network error.
    """
    r = httpx.get(
        _MCP_RELEASES_API,
        timeout=15,
        headers={"User-Agent": "GameCopilot"},
        follow_redirects=True,
    )
    r.raise_for_status()
    data = r.json()

    # Find mcp-server.py asset download URL
    download_url = _MCP_RAW_FALLBACK
    for asset in data.get("assets", []):
        if asset.get("name", "").lower() == "mcp-server.py":
            download_url = asset["browser_download_url"]
            break

    return {
        "tag_name": data.get("tag_name", ""),
        "published_at": data.get("published_at", ""),
        "download_url": download_url,
        "body": data.get("body", ""),
        "html_url": data.get("html_url", ""),
    }


@mcp.tool()
def check_for_mcp_server_update() -> dict:
    """Prüft ob eine neuere Version des Pimax-Graphics-MCP Servers auf GitHub verfügbar ist.

    Liest das neueste GitHub Release von Bennidesign2003/Pimax-Graphics-MCP
    und vergleicht die Version mit dem laufenden Server.
    Rufe dieses Tool auf wenn der User fragt ob der MCP Server aktuell ist.
    """
    current = _get_current_mcp_version()

    try:
        release = _fetch_latest_release_info()
    except Exception as exc:
        return {"error": f"GitHub nicht erreichbar: {exc}", "aktuelle_version": current}

    # Version aus Tag-Name lesen (z.B. "v3.6.1" → "3.6.1")
    remote_version = release["tag_name"].lstrip("v") or "unbekannt"

    update_available = False
    if current != "unbekannt" and remote_version != "unbekannt":
        try:
            from packaging.version import Version as _V
            update_available = _V(remote_version) > _V(current)
        except Exception:
            update_available = remote_version != current

    return {
        "aktuelle_version": current,
        "github_version": remote_version,
        "update_verfügbar": update_available,
        "release_datum": release["published_at"],
        "release_url": release["html_url"],
        "download_url": release["download_url"],
        "server_datei": str(Path(__file__)),
        "hinweis": (
            f"MCP Server Update verfügbar: {current} → {remote_version}. "
            f"Sage 'MCP Server aktualisieren' um das Update zu installieren."
        ) if update_available else (
            f"MCP Server {current} ist aktuell."
        ),
    }


@mcp.tool()
def update_mcp_server(confirmed: bool = False) -> dict:
    """Lädt die neueste Version des Nvidia MCP Servers von GitHub herunter und ersetzt die lokale Datei.

    Nach dem Update muss Game Copilot neugestartet werden (oder sage 'MCP Server neustarten').
    Frage den User IMMER nach Bestätigung bevor du confirmed=True setzt!

    Args:
        confirmed: Muss True sein — User hat Aktualisierung bestätigt.
    """
    if not confirmed:
        current = _get_current_mcp_version()
        return {
            "warte_auf_bestätigung": True,
            "aktuelle_version": current,
            "nachricht": (
                "Soll der Nvidia MCP Server jetzt aktualisiert werden? "
                "Der Server wird danach neu gestartet. (Ja / Nein)"
            ),
        }

    # Resolve download URL from latest GitHub Release asset
    download_url = _MCP_RAW_FALLBACK
    release_tag = ""
    try:
        release = _fetch_latest_release_info()
        download_url = release["download_url"]
        release_tag = release["tag_name"]
    except Exception:
        pass  # Fall back to raw URL

    # Download full server.py from GitHub Release asset
    try:
        r = httpx.get(download_url, timeout=60,
                      headers={"User-Agent": "GameCopilot"},
                      follow_redirects=True)
        r.raise_for_status()
        new_content = r.text
    except Exception as exc:
        return {"error": f"Download fehlgeschlagen ({download_url}): {exc}"}

    new_version = _get_remote_mcp_version(new_content)
    if not new_version:
        return {
            "error": "Heruntergeladene Datei enthält keine __mcp_version__ — Update abgebrochen.",
            "grund": "Sicherheitscheck: Datei sieht nicht wie ein gültiger MCP Server aus.",
        }

    # Backup current server.py
    server_path = Path(__file__)
    backup_path = server_path.parent / (server_path.name + ".bak")
    try:
        import shutil as _sh
        _sh.copy2(server_path, backup_path)
    except Exception:
        pass  # Backup is best-effort

    # Replace server.py with new version
    try:
        server_path.write_text(new_content, encoding="utf-8")
    except Exception as exc:
        return {"error": f"Datei konnte nicht ersetzt werden: {exc}"}

    # Write restart flag so McpClientService auto-restarts
    restart_flag = server_path.parent / "__mcp_restart_requested__"
    try:
        restart_flag.write_text(new_version, encoding="utf-8")
    except Exception:
        pass

    return {
        "status": "ok",
        "neue_version": new_version,
        "release_tag": release_tag,
        "download_quelle": download_url,
        "server_datei": str(server_path),
        "backup": str(backup_path),
        "neustart_erforderlich": True,
        "nachricht": (
            f"✅ MCP Server wurde auf Version {new_version} aktualisiert! "
            f"Sage jetzt 'MCP Server neustarten' damit die neue Version aktiv wird."
        ),
    }


# ===========================================================================
# NEW TOOLS — Game Copilot v3.6.0
# ===========================================================================

# ---------------------------------------------------------------------------
# A1. OpenXR Layers
# ---------------------------------------------------------------------------

@mcp.tool()
def get_openxr_layers() -> dict:
    """List all active OpenXR API layers from the Windows registry.

    API layers can affect VR performance and compatibility (e.g. OpenXR Toolkit,
    OVR Advanced Settings, Pimax OpenXR layer). Too many layers = overhead.

    Returns:
        implicit_layers: Always-on API layers (HKLM/HKCU registry).
        explicit_layers: Explicitly enabled layers (user-activated).
        total_count: Combined layer count.
    """
    try:
        implicit: list[dict] = []
        explicit: list[dict] = []

        ps_result = _ps(
            r"""
$result = @()
$paths = @(
    @{key='HKLM:\SOFTWARE\Khronos\OpenXR\1\ApiLayers\Implicit'; type='implicit'},
    @{key='HKCU:\SOFTWARE\Khronos\OpenXR\1\ApiLayers\Implicit'; type='implicit'},
    @{key='HKLM:\SOFTWARE\Khronos\OpenXR\1\ApiLayers\Explicit'; type='explicit'},
    @{key='HKCU:\SOFTWARE\Khronos\OpenXR\1\ApiLayers\Explicit'; type='explicit'}
)
foreach ($entry in $paths) {
    if (Test-Path $entry.key) {
        $vals = Get-ItemProperty $entry.key -ErrorAction SilentlyContinue
        if ($vals) {
            $vals.PSObject.Properties |
            Where-Object { $_.Name -notlike 'PS*' } |
            ForEach-Object {
                $result += "$($entry.type)|$($_.Name)|$($_.Value)"
            }
        }
    }
}
$result -join "`n"
"""
        )

        for line in ps_result.splitlines():
            line = line.strip()
            if "|" not in line:
                continue
            parts = line.split("|", 2)
            if len(parts) < 3:
                continue
            layer_type, name, value = parts
            layer_info = {
                "path": name.strip(),
                # For implicit layers: value 0 = enabled, 1 = disabled
                "enabled": value.strip() == "0",
            }
            if layer_type == "implicit":
                implicit.append(layer_info)
            else:
                explicit.append(layer_info)

        return {
            "implicit_layers": implicit,
            "explicit_layers": explicit,
            "total_count": len(implicit) + len(explicit),
            "hinweis": (
                "Implicit layers sind immer aktiv (können VR-Performance beeinflussen). "
                "Explicit layers müssen vom Nutzer aktiviert werden. "
                "Viele aktive Layers können Latenz erhöhen."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A2. MSFS Process / Performance Monitoring
# ---------------------------------------------------------------------------

@mcp.tool()
def is_msfs_running() -> dict:
    """Prüft ob Microsoft Flight Simulator gerade läuft (einfacher Boolean-Check).

    Rufe dies auf bevor du Konfig-Änderungen machst — MSFS muss für
    UserCfg.opt-Änderungen geschlossen sein.

    Returns:
        running: True wenn MSFS aktiv ist.
        warnung: Hinweis wenn MSFS läuft und Änderungen geplant sind.
    """
    try:
        running = _is_msfs_running()
        if running:
            procs = _ps_json(
                "Get-Process | Where-Object { "
                "$_.ProcessName -like '*FlightSimulator*' -or "
                "$_.ProcessName -like '*MSFS*' "
                "} | Select-Object ProcessName, Id -First 3"
            )
            return {
                "running": True,
                "prozesse": procs,
                "warnung": (
                    "MSFS läuft! Config-Änderungen werden beim nächsten "
                    "MSFS-Start möglicherweise überschrieben. MSFS erst schließen."
                ),
            }
        return {"running": False, "hinweis": "MSFS ist nicht aktiv — Config-Änderungen können sicher gespeichert werden."}
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_msfs_process_info() -> dict:
    """Detaillierte Prozess-Info für Microsoft Flight Simulator.

    Gibt PID, CPU-Auslastung, RAM-Verbrauch und Fenstertitel zurück.
    Nützlich um zu prüfen ob MSFS hängt oder zu viel Ressourcen verbraucht.
    """
    try:
        procs = _ps_json(
            "Get-Process | Where-Object { "
            "$_.ProcessName -like '*FlightSimulator*' -or "
            "$_.ProcessName -like '*MSFS*' "
            "} | Select-Object ProcessName, Id, "
            "@{N='CPU_s';E={[math]::Round($_.CPU,1)}}, "
            "@{N='MemoryMB';E={[math]::Round($_.WorkingSet64/1MB,1)}}, "
            "@{N='MemoryGB';E={[math]::Round($_.WorkingSet64/1GB,2)}}, "
            "MainWindowTitle, Responding, PriorityClass"
        )
        if not procs:
            return {"running": False, "hinweis": "MSFS läuft nicht."}

        proc = procs[0] if isinstance(procs, list) else procs
        memory_mb = float(proc.get("MemoryMB") or 0)
        responding = proc.get("Responding", True)

        warnungen = []
        if memory_mb > 20000:
            warnungen.append(f"⚠️ MSFS nutzt {memory_mb:.0f} MB RAM — sehr hoher Verbrauch")
        if not responding:
            warnungen.append("❌ MSFS antwortet nicht (möglicherweise eingefroren)")

        return {
            "running": True,
            "prozess_info": proc,
            "warnungen": warnungen,
            "zusammenfassung": (
                f"MSFS läuft (PID {proc.get('Id', '?')}) — "
                f"RAM: {proc.get('MemoryGB', '?')} GB"
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_msfs_fps_estimate() -> dict:
    """Schätzt MSFS-Performance über CPU/GPU-Last als Proxy für FPS.

    Da MSFS keine FPS-API bereitstellt, liest dieses Tool GPU-Auslastung
    (via NVML) und CPU-Last als Performance-Proxy. Für echte FPS:
    Strg+Z in MSFS oder SteamVR-Overlay verwenden.

    Returns:
        gpu_usage_percent: GPU-Auslastung (Nvidia NVML).
        system_cpu_percent: Gesamt-CPU-Last.
        msfs_memory_gb: MSFS RAM-Verbrauch.
        fps_hinweis: Einschätzung ob System unter VR-Last steht.
    """
    try:
        result: dict = {}

        procs = _ps_json(
            "Get-Process | Where-Object { "
            "$_.ProcessName -like '*FlightSimulator*' -or "
            "$_.ProcessName -like '*MSFS*' "
            "} | Select-Object ProcessName, "
            "@{N='CPU_s';E={[math]::Round($_.CPU,1)}}, "
            "@{N='MemoryGB';E={[math]::Round($_.WorkingSet64/1GB,2)}}"
        )

        if not procs:
            return {
                "running": False,
                "hinweis": "MSFS läuft nicht. Starte MSFS für Performance-Messung.",
            }

        result["msfs_running"] = True
        proc = procs[0] if isinstance(procs, list) else procs
        result["msfs_memory_gb"] = proc.get("MemoryGB", "?")

        # System CPU
        try:
            cpu_load = _ps_json(
                "Get-CimInstance Win32_Processor "
                "| Select-Object @{N='Load';E={$_.LoadPercentage}}"
            )
            if cpu_load:
                c = cpu_load[0] if isinstance(cpu_load, list) else cpu_load
                result["system_cpu_percent"] = c.get("Load", "?")
        except Exception:
            result["system_cpu_percent"] = "nicht verfügbar"

        # GPU via NVML
        try:
            pynvml.nvmlInit()
            handle = pynvml.nvmlDeviceGetHandleByIndex(0)
            util = pynvml.nvmlDeviceGetUtilizationRates(handle)
            result["gpu_usage_percent"] = util.gpu
            result["gpu_memory_used_percent"] = util.memory
            pynvml.nvmlShutdown()
        except Exception:
            result["gpu_usage_percent"] = "nicht verfügbar (kein NVML)"

        # FPS hint
        gpu_pct = result.get("gpu_usage_percent", 0)
        if isinstance(gpu_pct, (int, float)):
            if gpu_pct > 95:
                result["fps_hinweis"] = "GPU-limitiert (>95%) — FPS durch GPU begrenzt (ideal für VR!)"
            elif gpu_pct > 70:
                result["fps_hinweis"] = "GPU gut ausgelastet (70-95%) — gesunde VR-Last"
            elif gpu_pct > 40:
                result["fps_hinweis"] = "GPU mäßig ausgelastet — möglicherweise CPU-Bottleneck"
            else:
                result["fps_hinweis"] = "GPU wenig ausgelastet — CPU-Bottleneck oder MSFS im Menü"
        else:
            result["fps_hinweis"] = "GPU-Last nicht messbar"

        result["empfehlung"] = (
            "Für echte FPS: Strg+Z in MSFS (integrierter FPS-Counter) "
            "oder SteamVR Dashboard (zeigt FPS + Frametimes)."
        )
        return result
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A3. Audio / VR Audio
# ---------------------------------------------------------------------------

@mcp.tool()
def get_audio_devices() -> dict:
    """Listet alle Windows-Audio-Geräte auf (Wiedergabe und Aufnahme).

    VR-Headsets (Pimax, Index, Quest) erscheinen oft als eigene Audio-Geräte.
    Nutze set_default_audio_device() um das VR-Headset-Audio als Standard zu setzen.

    Returns:
        geräte: Alle Sound-Geräte (WMI + PnP).
        tipp: Hinweis zum Setzen des Standard-Geräts.
    """
    try:
        # WMI Sound devices (always works)
        wmi_devices = _ps_json(
            "Get-WmiObject Win32_SoundDevice "
            "| Select-Object Name, Status, DeviceID, Manufacturer"
        )

        # PnP audio endpoints (more detailed)
        pnp_devices = _ps_json(
            "Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue "
            "| Where-Object { $_.Status -eq 'OK' } "
            "| Select-Object FriendlyName, InstanceId, Status"
        )

        return {
            "geräte_wmi": wmi_devices if isinstance(wmi_devices, list) else [],
            "geräte_pnp": pnp_devices if isinstance(pnp_devices, list) else [],
            "tipp": (
                "VR-Headset-Audio als Standard setzen: "
                "set_default_audio_device(device_name='Headset-Name') "
                "oder manuell: Systemsteuerung → Sound → Wiedergabe."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_default_audio_device(device_name: str) -> dict:
    """Setzt ein Audio-Gerät als Windows-Standard-Wiedergabegerät.

    Nützlich um das VR-Headset-Audio (Pimax, Index, Quest) als Standard
    zu setzen damit MSFS-Audio durch das Headset abgespielt wird.

    Args:
        device_name: Teil des Gerätenamens (Groß-/Kleinschreibung egal).
                     z.B. 'Pimax', 'Headphones', 'Headset', 'Index'.

    Returns:
        status: 'ok', 'nicht_gefunden', oder 'manuell_erforderlich'.
    """
    try:
        if not device_name:
            return {"error": "device_name darf nicht leer sein."}

        # Check if device exists
        check = _ps_json(
            f"Get-WmiObject Win32_SoundDevice "
            f"| Where-Object {{ $_.Name -like '*{device_name}*' }} "
            f"| Select-Object Name, DeviceID"
        )

        if not check:
            # Try PnP
            check = _ps_json(
                f"Get-PnpDevice -Class AudioEndpoint -ErrorAction SilentlyContinue "
                f"| Where-Object {{ $_.FriendlyName -like '*{device_name}*' -and $_.Status -eq 'OK' }} "
                f"| Select-Object FriendlyName, InstanceId"
            )

        if not check:
            return {
                "status": "nicht_gefunden",
                "gerät": device_name,
                "hinweis": (
                    f"Gerät '{device_name}' nicht gefunden. "
                    "Nutze get_audio_devices() um den genauen Namen zu sehen."
                ),
            }

        # Try nircmd.exe (free tool from nirsoft.net)
        nircmd_result = _ps(
            f"""
$nircmd = "$env:SystemRoot\\nircmd.exe"
$nircmd2 = "$env:ProgramFiles\\nircmd\\nircmd.exe"
if (Test-Path $nircmd) {{
    & $nircmd setdefaultsounddevice '{device_name}' 1
    Write-Output "OK_NIRCMD"
}} elseif (Test-Path $nircmd2) {{
    & $nircmd2 setdefaultsounddevice '{device_name}' 1
    Write-Output "OK_NIRCMD"
}} else {{
    Write-Output "MANUAL_REQUIRED"
}}
"""
        )

        if "OK_NIRCMD" in nircmd_result:
            return {
                "status": "ok",
                "gerät": device_name,
                "methode": "nircmd",
                "nachricht": f"'{device_name}' wurde als Standard-Audiogerät gesetzt.",
            }

        return {
            "status": "manuell_erforderlich",
            "gerät_gefunden": True,
            "gefundene_geräte": check,
            "anleitung": (
                f"Gerät '{device_name}' gefunden. "
                "Automatisches Setzen erfordert NirCmd (nirsoft.net/utils/nircmd.html). "
                "Manuell: Systemsteuerung → Sound → Wiedergabe → Gerät → 'Als Standard festlegen'."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_spatial_audio_status() -> dict:
    """Prüft ob Windows Spatial Audio (Windows Sonic oder Dolby Atmos) aktiviert ist.

    Spatial Audio verbessert VR-Immersion. Windows Sonic for Headphones ist kostenlos
    und direkt in Windows 10/11 integriert.

    Returns:
        spatial_audio_enabled: Ob Spatial Audio aktiv ist.
        empfehlung: Ob Spatial Audio für VR aktiviert werden sollte.
    """
    try:
        # Check registry for spatial audio settings
        result = _ps(
            r"""
$enabled = $false
$format = 'Keines'

$spatialFormats = @{
    '{B19349B0-0BBE-4f4a-A938-E60A5F39E5CF}' = 'Windows Sonic for Headphones'
    '{B18F4B13-6FE9-4209-8A53-79A67A1FA08B}' = 'Dolby Atmos for Headphones'
    '{D3489861-8414-429C-B3CD-84E6C3CEF5E1}' = 'DTS Headphone:X'
}

$sonicKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\SpatialAudio'
if (Test-Path $sonicKey) {
    $val = Get-ItemPropertyValue $sonicKey -Name 'SpatialAudioMode' -ErrorAction SilentlyContinue
    if ($null -ne $val) {
        $enabled = $true
        $fmt = $spatialFormats[$val]
        $format = if ($fmt) { $fmt } else { "Unbekanntes Format ($val)" }
    }
}

# Also check via device-level registry
$renderBase = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
if (-not $enabled -and (Test-Path $renderBase)) {
    $devs = Get-ChildItem $renderBase -ErrorAction SilentlyContinue | Select-Object -First 5
    foreach ($dev in $devs) {
        $propsPath = "$($dev.PSPath)\Properties"
        if (Test-Path $propsPath) {
            # Property {1DA5D803-D492-4EDD-8C23-E0C0FFEE7F0E},6 = Spatial format GUID
            $p = Get-ItemProperty $propsPath -ErrorAction SilentlyContinue
            $spatialProp = $p.PSObject.Properties |
                Where-Object { $_.Name -like '*1DA5D803*' } |
                Select-Object -First 1
            if ($spatialProp -and $spatialProp.Value) {
                $enabled = $true
                $format = "Spatial Audio (gerätespezifisch)"
            }
        }
    }
}

"$enabled|$format"
"""
        )

        parts = result.strip().split("|", 1)
        enabled = parts[0].strip().lower() == "true" if parts else False
        fmt = parts[1].strip() if len(parts) > 1 else "Keines"

        return {
            "spatial_audio_enabled": enabled,
            "format": fmt,
            "empfehlung": (
                "✅ Spatial Audio aktiv — gut für VR-Immersion." if enabled
                else (
                    "💡 Spatial Audio nicht aktiv. Für VR empfohlen: "
                    "Systemsteuerung → Sound → Wiedergabe → Gerät → Eigenschaften "
                    "→ Spatial Sound → 'Windows Sonic for Headphones' aktivieren."
                )
            ),
            "hinweis": "Windows Sonic ist kostenlos und verbessert VR-Klang erheblich.",
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A4. Display & Resolution Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def get_display_info() -> dict:
    """Listet alle Monitore: Auflösung, Refresh Rate, HDR-Status und Primary-Flag.

    Im VR-Modus den Desktop-Monitor auf niedrige Auflösung/Refresh zu setzen
    gibt GPU-Ressourcen für das Headset frei.

    Returns:
        monitore: Liste aktiver Monitore mit technischen Details.
        grafikkarte: Aktuelle Auflösung/Refresh der primären GPU.
    """
    try:
        # Screen info via .NET Forms
        screen_info = _ps_json(
            r"""
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
$screens = [System.Windows.Forms.Screen]::AllScreens
$out = @()
foreach ($s in $screens) {
    $out += @{
        DeviceName   = $s.DeviceName
        Width        = $s.Bounds.Width
        Height       = $s.Bounds.Height
        IsPrimary    = $s.Primary
        BitsPerPixel = $s.BitsPerPixel
    }
}
$out
"""
        )

        # GPU/display details via WMI
        gpu_info = _ps_json(
            "Get-CimInstance Win32_VideoController "
            "| Select-Object Name, CurrentRefreshRate, "
            "CurrentHorizontalResolution, CurrentVerticalResolution, "
            "@{N='VRAM_GB';E={[math]::Round($_.AdapterRAM/1GB,1)}}"
        )

        # HDR check via registry
        hdr_status = _ps(
            r"""
$hdrEnabled = $false
$key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\VideoSettings'
if (Test-Path $key) {
    $val = Get-ItemPropertyValue $key -Name 'EnableHDROutput' -ErrorAction SilentlyContinue
    $hdrEnabled = ($val -eq 1)
}
$hdrEnabled
"""
        )

        hdr_on = hdr_status.strip().lower() == "true"

        return {
            "monitore": screen_info if isinstance(screen_info, list) else [],
            "grafikkarte": gpu_info if isinstance(gpu_info, list) else [],
            "hdr_aktiv": hdr_on,
            "tipp_vr": (
                "Im VR-Betrieb: Desktop-Monitor auf 1920×1080 @ 60 Hz setzen "
                "(Windows Anzeigeeinstellungen) um GPU-Kapazität für das Headset freizugeben."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_vr_render_resolution() -> dict:
    """Berechnet die effektive VR-Renderauflösung aus Pimax + SteamVR Supersampling.

    Kombiniert native Headset-Auflösung, Pimax Render Quality und SteamVR SS-Faktor
    um die tatsächlich gerenderte Auflösung pro Auge zu berechnen.
    Hohe Renderauflösungen kosten überproportional viel VRAM und GPU-Zeit.

    Returns:
        render_auflösung_pro_auge: Effektiv gerenderte Auflösung.
        gesamt_mpixel: Gesamtpixel (beide Augen) in Megapixel.
        empfehlung: Ob die Kombination für die Hardware angemessen ist.
    """
    try:
        result: dict = {}

        # Pimax headset
        try:
            pimax_info = get_pimax_headset_info()
            native_w = int(pimax_info.get("display_width_per_eye") or 2160)
            native_h = int(pimax_info.get("display_height_per_eye") or 2160)
            headset_name = pimax_info.get("headset_name", "Unbekannt")
            render_quality = float(pimax_info.get("render_quality") or 1.0)
        except Exception:
            native_w, native_h = 2160, 2160
            headset_name = "Unbekannt"
            render_quality = 1.0

        result["headset"] = headset_name
        result["native_auflösung_pro_auge"] = f"{native_w}x{native_h}"
        result["pimax_render_quality"] = render_quality

        # SteamVR supersampling
        steamvr_ss = 1.0
        svr_path = _find_steamvr_settings_path()
        if svr_path and svr_path.exists():
            try:
                svr_data = _json.loads(svr_path.read_text(encoding="utf-8"))
                steamvr_ss = float(svr_data.get("steamvr", {}).get("supersampleScale") or 1.0)
                result["steamvr_supersampling"] = steamvr_ss
            except Exception:
                result["steamvr_supersampling"] = "nicht lesbar"
        else:
            result["steamvr_supersampling"] = "SteamVR nicht gefunden"

        # Calculate
        effective_ss = render_quality * steamvr_ss
        render_w = int(native_w * (effective_ss ** 0.5))
        render_h = int(native_h * (effective_ss ** 0.5))
        total_pixels = render_w * render_h * 2

        result["effektiver_ss_faktor"] = round(effective_ss, 3)
        result["render_auflösung_pro_auge"] = f"{render_w}x{render_h}"
        result["gesamt_pixel_beide_augen"] = f"{total_pixels:,}"
        result["gesamt_mpixel"] = round(total_pixels / 1_000_000, 1)

        # Hardware tier
        try:
            hw = get_detailed_hardware_profile()
            vram_gb = float(hw.get("gpu", {}).get("vram_total_gb") or 8)
            tier = hw.get("vr_tier", "mid")
        except Exception:
            vram_gb, tier = 8.0, "mid"

        mpixels = result["gesamt_mpixel"]
        if mpixels > 20:
            empf = f"⚠️ {mpixels}M Pixel sehr hoch für {tier}-Tier ({vram_gb:.0f}GB VRAM). Reduziere Pimax RQ oder SteamVR SS."
        elif mpixels > 14:
            empf = f"💡 {mpixels}M Pixel — anspruchsvoll. OK für high/ultra-Tier."
        else:
            empf = f"✅ {mpixels}M Pixel — angemessene Renderauflösung."

        result["empfehlung"] = empf
        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_windows_display_scaling(scale_percent: int = 100) -> dict:
    """Setzt die Windows DPI-Skalierung für den primären Monitor via Registry.

    Für VR empfohlen: 100% (keine Skalierung) setzt GPU-Ressourcen frei
    und reduziert Kompatibilitätsprobleme mit VR-Software.

    Args:
        scale_percent: Skalierung in Prozent. Gültig: 100, 125, 150, 175, 200, 225, 250.

    Hinweis: Vollständige Wirkung nach Neuanmeldung oder Explorer-Neustart.
    """
    try:
        valid_scales = [100, 125, 150, 175, 200, 225, 250]
        if scale_percent not in valid_scales:
            return {
                "error": f"Ungültiger Skalierungswert '{scale_percent}'. Gültig: {valid_scales}",
            }

        dpi_value = int(scale_percent * 96 / 100)

        out = _ps(
            f"""
Set-ItemProperty -Path 'HKCU:\\Control Panel\\Desktop' -Name 'LogPixels' -Value {dpi_value} -Type DWord -Force
Set-ItemProperty -Path 'HKCU:\\Control Panel\\Desktop' -Name 'Win8DpiScaling' -Value 1 -Type DWord -Force
Write-Output "DPI_SET:{dpi_value}"
"""
        )

        if f"DPI_SET:{dpi_value}" in out:
            return {
                "status": "ok",
                "skalierung_prozent": scale_percent,
                "dpi_wert": dpi_value,
                "hinweis": (
                    "DPI-Skalierung gesetzt. Effekt vollständig nach Neuanmeldung "
                    "oder Explorer-Neustart (Taskmanager → Explorer.exe neustarten)."
                ),
            }
        return {"status": "unbekannt", "ausgabe": out}
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A5. Targeted System Performance Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def get_cpu_info() -> dict:
    """Detaillierte CPU-Info: Kerne, Threads, Takt, Auslastung und Temperatur.

    MSFS VR ist stark Single-Core-limitiert — hoher Turbotakt (>5GHz)
    hat mehr Einfluss als zusätzliche Kerne.

    Returns:
        name, kerne, threads, max_takt_ghz, auslastung_prozent,
        temperatur_celsius, vr_bewertung.
    """
    try:
        cpu_raw = _ps_json(
            "Get-CimInstance Win32_Processor | Select-Object "
            "Name, NumberOfCores, NumberOfLogicalProcessors, "
            "MaxClockSpeed, CurrentClockSpeed, LoadPercentage, Manufacturer"
        )
        cpu = (cpu_raw[0] if isinstance(cpu_raw, list) and cpu_raw else cpu_raw) or {}

        # Temperature via WMI thermal zones (requires admin — may fail)
        temp = None
        try:
            temp_raw = _ps_json(
                "Get-CimInstance -Namespace root/wmi "
                "-ClassName MSAcpi_ThermalZoneTemperature "
                "-ErrorAction SilentlyContinue "
                "| Select-Object CurrentTemperature"
            )
            if temp_raw:
                items = temp_raw if isinstance(temp_raw, list) else [temp_raw]
                temps = []
                for item in items:
                    raw = item.get("CurrentTemperature") or 0
                    if raw > 2000:  # Kelvin * 10
                        temps.append(round((raw - 2732) / 10.0, 1))
                if temps:
                    temp = max(t for t in temps if t > 0)
        except Exception:
            pass

        cores = int(cpu.get("NumberOfCores") or 0)
        threads = int(cpu.get("NumberOfLogicalProcessors") or 0)
        max_clock = int(cpu.get("MaxClockSpeed") or 0)
        load = cpu.get("LoadPercentage") or 0
        name = cpu.get("Name", "Unbekannt")

        if max_clock >= 5000 and cores >= 8:
            vr_rating = "✅ Ausgezeichnet für MSFS VR"
        elif max_clock >= 4000 and cores >= 6:
            vr_rating = "✅ Gut für MSFS VR"
        elif max_clock >= 3500 and cores >= 4:
            vr_rating = "⚠️ Ausreichend — hoher Single-Core-Takt bevorzugt"
        else:
            vr_rating = "❌ CPU-Bottleneck in MSFS VR wahrscheinlich"

        temp_warnung = ""
        if isinstance(temp, float):
            if temp > 90:
                temp_warnung = f"🔥 CPU sehr heiß ({temp}°C) — Kühlung prüfen!"
            elif temp > 80:
                temp_warnung = f"⚠️ CPU warm ({temp}°C)"

        return {
            "name": name,
            "kerne_physisch": cores,
            "threads_logisch": threads,
            "max_takt_mhz": max_clock,
            "max_takt_ghz": round(max_clock / 1000, 2) if max_clock else "?",
            "auslastung_prozent": load,
            "temperatur_celsius": temp if temp is not None else "nicht verfügbar (Admin erforderlich)",
            "temp_warnung": temp_warnung,
            "vr_bewertung": vr_rating,
            "hinweis_msfs": "MSFS 2024 ist stark Single-Core-limitiert. Turbotakt > 5GHz = mehr FPS.",
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_ram_info() -> dict:
    """Detaillierte RAM-Info: Gesamt, verfügbar, Takt, Slots und Dual-Channel.

    Für MSFS VR sind 32 GB empfohlen. Dual-Channel erhöht die Speicherbandbreite
    und reduziert CPU-Stottern erheblich.

    Returns:
        total_gb, available_gb, speed_mhz, slots_belegt/gesamt,
        dual_channel, vr_bewertung.
    """
    try:
        os_ram = _ps_json(
            "Get-CimInstance Win32_OperatingSystem | Select-Object "
            "@{N='TotalGB';E={[math]::Round($_.TotalVisibleMemorySize/1MB,1)}}, "
            "@{N='FreeGB';E={[math]::Round($_.FreePhysicalMemory/1MB,1)}}"
        )
        os = (os_ram[0] if isinstance(os_ram, list) and os_ram else os_ram) or {}
        total_gb = float(os.get("TotalGB") or 0)
        free_gb = float(os.get("FreeGB") or 0)

        modules = _ps_json(
            "Get-CimInstance Win32_PhysicalMemory | Select-Object "
            "BankLabel, @{N='CapacityGB';E={[math]::Round($_.Capacity/1GB,0)}}, "
            "Speed, Manufacturer"
        )

        slots_total = 0
        try:
            slot_arr = _ps_json(
                "Get-CimInstance Win32_PhysicalMemoryArray "
                "| Select-Object MemoryDevices"
            )
            if slot_arr:
                s = slot_arr[0] if isinstance(slot_arr, list) else slot_arr
                slots_total = int(s.get("MemoryDevices") or 0)
        except Exception:
            pass

        speed_mhz = 0
        module_list = []
        if isinstance(modules, list):
            for m in modules:
                sp = int(m.get("Speed") or 0)
                if sp > speed_mhz:
                    speed_mhz = sp
                module_list.append({
                    "slot": m.get("BankLabel", "?"),
                    "größe_gb": m.get("CapacityGB", "?"),
                    "takt_mhz": sp,
                })
        slots_used = len(module_list)
        likely_dual = slots_used in (2, 4)

        if total_gb >= 32:
            vr_rating = "✅ Ausgezeichnet für MSFS VR (32+ GB)"
        elif total_gb >= 24:
            vr_rating = "✅ Gut (24 GB)"
        elif total_gb >= 16:
            vr_rating = "⚠️ Ausreichend — 32 GB empfohlen"
        else:
            vr_rating = f"❌ {total_gb:.0f} GB zu wenig — mind. 16 GB, empfohlen 32 GB"

        return {
            "gesamt_gb": total_gb,
            "verfügbar_gb": round(free_gb, 1),
            "genutzt_gb": round(total_gb - free_gb, 1),
            "takt_mhz": speed_mhz,
            "slots_belegt": slots_used,
            "slots_gesamt": slots_total,
            "dual_channel_wahrscheinlich": likely_dual,
            "module": module_list,
            "vr_bewertung": vr_rating,
            "hinweis": "Dual-Channel (2 oder 4 Module) verdoppelt Speicherbandbreite — wichtig für CPU-Performance in MSFS.",
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_disk_info() -> dict:
    """Laufwerks-Info: Kapazität, freier Speicher, SSD vs. HDD, MSFS-Laufwerk.

    MSFS 2024 sollte auf einer NVMe-SSD installiert sein. HDD führt zu
    langen Ladezeiten und schlechtem Terrain-Streaming in VR.

    Returns:
        laufwerke: Alle Laufwerke mit Typ, Größe, freiem Speicher.
        msfs_laufwerk: Das Laufwerk mit der MSFS-Installation.
        empfehlung: Ob die Konfiguration optimal ist.
    """
    try:
        logical = _ps_json(
            "Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | "
            "Select-Object DeviceID, VolumeName, "
            "@{N='GesamtGB';E={[math]::Round($_.Size/1GB,1)}}, "
            "@{N='FreiGB';E={[math]::Round($_.FreeSpace/1GB,1)}}, "
            "@{N='GenutztProzent';E={[math]::Round(($_.Size-$_.FreeSpace)/$_.Size*100,1)}}"
        )

        physical = _ps_json(
            "Get-PhysicalDisk | Select-Object FriendlyName, MediaType, BusType, "
            "@{N='GroesseGB';E={[math]::Round($_.Size/1GB,1)}}"
        )

        # Build SSD lookup from media type
        ssd_hints: set[str] = set()
        if isinstance(physical, list):
            for pd in physical:
                mt = str(pd.get("MediaType") or "")
                bt = str(pd.get("BusType") or "")
                if "SSD" in mt or "Solid" in mt or "NVMe" in bt or "NVM" in bt:
                    ssd_hints.add(pd.get("FriendlyName", ""))

        # Find MSFS drive
        msfs_drive: str | None = None
        cfg = _find_usercfg()
        if cfg:
            msfs_drive = str(cfg)[:2].upper()

        disk_list = []
        if isinstance(logical, list):
            for d in logical:
                drive_id = str(d.get("DeviceID", "")).upper()
                disk_type = "SSD" if any(h for h in ssd_hints) else "Unbekannt"
                frei = float(d.get("FreiGB") or 0)
                entry = dict(d)
                entry["typ"] = disk_type
                entry["ist_msfs_laufwerk"] = (drive_id == msfs_drive)
                if frei < 10:
                    entry["warnung"] = f"⚠️ Wenig Speicher ({frei:.1f} GB frei)"
                disk_list.append(entry)

        msfs_info = next((d for d in disk_list if d.get("ist_msfs_laufwerk")), None)
        empfehlung = "✅ MSFS-Laufwerk: SSD erkannt." if msfs_info else "MSFS-Laufwerk nicht identifiziert."
        if msfs_info:
            frei = float(msfs_info.get("FreiGB") or 0)
            if frei < 20:
                empfehlung += f" ⚠️ Nur {frei:.1f} GB frei — MSFS benötigt Platz für Rolling Cache."

        return {
            "laufwerke": disk_list,
            "physische_laufwerke": physical if isinstance(physical, list) else [],
            "msfs_laufwerk": msfs_drive,
            "msfs_laufwerk_info": msfs_info,
            "empfehlung": empfehlung,
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_system_temps() -> dict:
    """Liest CPU- und GPU-Temperaturen aus.

    Hohe Temperaturen führen zu Thermal Throttling = VR-Stottern.
    GPU-Temp via NVML (Nvidia), CPU-Temp via WMI (erfordert ggf. Admin).

    Returns:
        cpu_temp_celsius, gpu_temp_celsius, gpu_hotspot_celsius, bewertung.
    """
    try:
        result: dict = {}

        # CPU via WMI thermal zones
        cpu_temp = None
        try:
            temp_data = _ps_json(
                "Get-CimInstance -Namespace root/wmi "
                "-ClassName MSAcpi_ThermalZoneTemperature "
                "-ErrorAction SilentlyContinue "
                "| Select-Object CurrentTemperature"
            )
            if temp_data:
                items = temp_data if isinstance(temp_data, list) else [temp_data]
                valid = []
                for item in items:
                    raw = item.get("CurrentTemperature") or 0
                    if raw > 2000:
                        c = round((raw - 2732) / 10.0, 1)
                        if 0 < c < 120:
                            valid.append(c)
                if valid:
                    cpu_temp = max(valid)
        except Exception:
            pass

        result["cpu_temp_celsius"] = (
            cpu_temp if cpu_temp is not None
            else "nicht verfügbar (WMI-Admin oder HWiNFO erforderlich)"
        )

        # GPU via NVML
        gpu_temp = gpu_hotspot = gpu_mem_temp = None
        try:
            pynvml.nvmlInit()
            h = pynvml.nvmlDeviceGetHandleByIndex(0)
            gpu_temp = pynvml.nvmlDeviceGetTemperature(h, pynvml.NVML_TEMPERATURE_GPU)
            try:
                gpu_hotspot = pynvml.nvmlDeviceGetTemperature(h, 1)
            except Exception:
                pass
            try:
                gpu_mem_temp = pynvml.nvmlDeviceGetTemperature(h, 2)
            except Exception:
                pass
            pynvml.nvmlShutdown()
        except Exception:
            pass

        result["gpu_temp_celsius"] = gpu_temp if gpu_temp is not None else "nicht verfügbar"
        if gpu_hotspot is not None:
            result["gpu_hotspot_celsius"] = gpu_hotspot
        if gpu_mem_temp is not None:
            result["gpu_vram_temp_celsius"] = gpu_mem_temp

        # Assessment
        warnungen = []
        ok_msgs = []
        if isinstance(cpu_temp, float):
            if cpu_temp > 95:
                warnungen.append(f"🔥 CPU-Temperatur kritisch: {cpu_temp}°C!")
            elif cpu_temp > 85:
                warnungen.append(f"⚠️ CPU warm: {cpu_temp}°C")
            else:
                ok_msgs.append(f"✅ CPU: {cpu_temp}°C — ok")
        if isinstance(gpu_temp, int):
            if gpu_temp > 90:
                warnungen.append(f"🔥 GPU-Temperatur kritisch: {gpu_temp}°C!")
            elif gpu_temp > 83:
                warnungen.append(f"⚠️ GPU warm: {gpu_temp}°C")
            else:
                ok_msgs.append(f"✅ GPU: {gpu_temp}°C — ok")

        result["bewertung"] = (warnungen + ok_msgs) if (warnungen or ok_msgs) else [
            "Temperaturmessung nicht verfügbar"
        ]
        result["hinweis"] = "GPU unter VR-Last: 70-83°C normal. Über 90°C = Throttling-Risiko."
        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def kill_background_processes(dry_run: bool = False) -> dict:
    """Beendet bekannte Hintergrundprozesse die VR-Performance beeinträchtigen.

    Ziele: Discord-Overlay, Xbox Game Bar, GeForce Experience Overlay,
    Xbox DVR, Cortana, Windows Search UI, OneDrive Sync.

    Args:
        dry_run: True = zeigt nur was beendet würde, ohne Änderungen.

    Returns:
        beendet: Beendete Prozesse.
        nicht_laufend: Prozesse die nicht aktiv waren.
    """
    # (process_name, display_name, safe_to_auto_kill)
    VR_KILLERS = [
        ("XboxGameBar",       "Xbox Game Bar",              True),
        ("GameBarFTServer",   "Xbox Game Bar FT Server",    True),
        ("XboxApp",           "Xbox App",                   True),
        ("GameBar",           "Windows Game Bar",           True),
        ("Cortana",           "Cortana",                    True),
        ("SearchApp",         "Windows Search App",         True),
        ("SearchUI",          "Windows Search UI",          True),
        ("NvBackend",         "GeForce Experience Backend", True),
        ("nvcontainer",       "Nvidia Container (GFE)",     True),
        ("Discord",           "Discord",                    True),
        ("OneDrive",          "OneDrive Sync",              True),
        ("Teams",             "Microsoft Teams",            True),
        ("slack",             "Slack",                      True),
    ]
    try:
        beendet = []
        nicht_laufend = []
        fehler = []

        for proc_name, display_name, _safe in VR_KILLERS:
            check = _ps_json(
                f"Get-Process -Name '{proc_name}' -ErrorAction SilentlyContinue "
                f"| Select-Object ProcessName, Id"
            )
            if not check:
                nicht_laufend.append(display_name)
                continue

            if dry_run:
                beendet.append(f"[WÜRDE BEENDEN] {display_name}")
            else:
                try:
                    _ps(f"Stop-Process -Name '{proc_name}' -Force -ErrorAction SilentlyContinue")
                    beendet.append(f"✅ {display_name} beendet")
                except Exception as ke:
                    fehler.append(f"⚠️ {display_name}: {ke}")

        # Also disable Xbox Game DVR
        if not dry_run:
            try:
                _reg_add(r"HKCU\System\GameConfigStore", "GameDVR_Enabled", "REG_DWORD", "0")
                _reg_add(
                    r"HKCU\Software\Microsoft\Windows\CurrentVersion\GameDVR",
                    "AppCaptureEnabled", "REG_DWORD", "0"
                )
                beendet.append("✅ Xbox Game DVR deaktiviert")
            except Exception:
                pass

        killed_count = len([x for x in beendet if not x.startswith("[")])
        return {
            "dry_run": dry_run,
            "beendet": beendet,
            "nicht_laufend": nicht_laufend,
            "fehler": fehler,
            "zusammenfassung": (
                f"{'[Simulation] ' if dry_run else ''}"
                f"{killed_count} Prozesse beendet, {len(nicht_laufend)} nicht aktiv."
            ),
            "tipp": "Starte MSFS VR direkt nach diesem Tool für maximale Performance.",
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A6. MSFS Community / Official Folder Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def get_msfs_community_folder(usercfg_path: str = "") -> dict:
    """Findet den MSFS Community-Ordner (Add-on Installationsverzeichnis).

    Liest InstalledPackagesPath aus UserCfg.opt oder sucht in Standard-Pfaden.
    Der Community-Ordner ist wo Add-ons (Flugzeuge, Szenerien, Liveries) landen.

    Returns:
        community_pfad: Absoluter Pfad.
        addon_anzahl: Anzahl installierter Add-ons.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        community_path: Path | None = None

        if cfg_path and cfg_path.exists():
            text = cfg_path.read_text(encoding="utf-8", errors="replace")
            for line in text.splitlines():
                line = line.strip()
                if line.startswith("InstalledPackagesPath"):
                    parts = line.split(None, 1)
                    if len(parts) > 1:
                        ps = parts[1].strip().strip('"')
                        community_path = Path(ps) / "Community"
                        break

        if not community_path:
            detected = _detect_community_folder("")
            if detected:
                community_path = detected

        if not community_path:
            for candidate in [
                Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
                / "Microsoft.Limitless_8wekyb3d8bbwe" / "LocalCache" / "Packages" / "Community",
                Path(os.environ.get("APPDATA", ""))
                / "Microsoft Flight Simulator 2024" / "Packages" / "Community",
                Path(os.environ.get("APPDATA", ""))
                / "Microsoft Flight Simulator" / "Packages" / "Community",
            ]:
                if candidate.exists():
                    community_path = candidate
                    break

        if not community_path:
            return {
                "fehler": "Community-Ordner nicht gefunden.",
                "hinweis": "Starte MSFS einmal um den Community-Ordner zu erstellen.",
            }

        addon_count = sum(1 for p in community_path.iterdir() if p.is_dir()) if community_path.exists() else 0

        return {
            "community_pfad": str(community_path),
            "existiert": community_path.exists(),
            "addon_anzahl": addon_count,
            "hinweis": f"{addon_count} Add-ons gefunden. Nutze list_community_addons() für Details.",
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def list_community_addons(usercfg_path: str = "") -> dict:
    """Listet alle installierten MSFS Add-ons im Community-Ordner.

    Liest manifest.json um Name, Version, Hersteller und Typ zu ermitteln
    (Flugzeug, Szenerie, Livery, etc.). Gibt Größe pro Add-on aus.

    Returns:
        addons: Liste mit Name, Version, Typ, Größe.
        gesamt, größe_gesamt_gb.
    """
    try:
        info = get_msfs_community_folder(usercfg_path)
        if "fehler" in info or "error" in info:
            return info

        community_path = Path(info["community_pfad"])
        if not community_path.exists():
            return {"fehler": "Community-Ordner existiert nicht.", "pfad": str(community_path)}

        addons = []
        total_bytes = 0

        for item in sorted(community_path.iterdir()):
            if not item.is_dir():
                continue

            entry: dict = {
                "ordner": item.name,
                "name": item.name,
                "version": "?",
                "typ": "Unbekannt",
                "hersteller": "?",
                "größe_mb": "?",
            }

            manifest = item / "manifest.json"
            if manifest.exists():
                try:
                    mf = _json.loads(manifest.read_text(encoding="utf-8", errors="replace"))
                    entry["name"] = mf.get("name", item.name)
                    entry["version"] = mf.get("package_version", "?")
                    entry["hersteller"] = mf.get("manufacturer", "?")
                    ct = mf.get("content_type", "").lower()
                    type_map = {
                        "aircraft": "Flugzeug", "airplane": "Flugzeug",
                        "scenery": "Szenerie", "livery": "Livery",
                        "misc": "Sonstiges", "other": "Sonstiges",
                    }
                    entry["typ"] = type_map.get(ct, ct.capitalize() or "Sonstiges")
                except Exception:
                    pass

            # Size (top-level rglob, skip deep trees > 5 s)
            try:
                size = sum(f.stat().st_size for f in item.rglob("*") if f.is_file())
                entry["größe_mb"] = round(size / 1024 / 1024, 1)
                total_bytes += size
            except Exception:
                pass

            addons.append(entry)

        return {
            "addons": addons,
            "gesamt": len(addons),
            "größe_gesamt_gb": round(total_bytes / 1024 ** 3, 2),
            "community_pfad": str(community_path),
            "tipp": "Inaktive Add-ons in 'Community.off'-Ordner verschieben um Ladezeit zu reduzieren.",
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_msfs_official_folder(usercfg_path: str = "") -> dict:
    """Findet den MSFS Official-Ordner (Basis-Spiel und Marketplace-Inhalte).

    Der Official-Ordner enthält von Asobo/Microsoft bereitgestellte Inhalte
    und sollte NIEMALS manuell verändert werden.

    Returns:
        official_pfad: Absoluter Pfad.
        pakete_gesamt: Anzahl offizieller Pakete.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        official_path: Path | None = None

        if cfg_path and cfg_path.exists():
            text = cfg_path.read_text(encoding="utf-8", errors="replace")
            for line in text.splitlines():
                line = line.strip()
                if line.startswith("InstalledPackagesPath"):
                    parts = line.split(None, 1)
                    if len(parts) > 1:
                        ps = parts[1].strip().strip('"')
                        official_path = Path(ps) / "Official"
                        break

        if not official_path:
            for candidate in [
                Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
                / "Microsoft.Limitless_8wekyb3d8bbwe" / "LocalCache" / "Packages" / "Official",
                Path(os.environ.get("APPDATA", ""))
                / "Microsoft Flight Simulator 2024" / "Packages" / "Official",
                Path(os.environ.get("APPDATA", ""))
                / "Microsoft Flight Simulator" / "Packages" / "Official",
            ]:
                if candidate.exists():
                    official_path = candidate
                    break

        if not official_path or not official_path.exists():
            return {
                "fehler": "Official-Ordner nicht gefunden.",
                "hinweis": "MSFS einmal starten um den Official-Ordner zu erstellen.",
            }

        publishers = [p for p in official_path.iterdir() if p.is_dir()]
        total_pkgs = sum(
            1 for pub in publishers for p in pub.iterdir() if p.is_dir()
        )

        return {
            "official_pfad": str(official_path),
            "existiert": True,
            "verleger": [p.name for p in publishers],
            "pakete_gesamt": total_pkgs,
            "warnung": "Diesen Ordner NIEMALS manuell ändern — MSFS repariert sich sonst selbst!",
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A7. Network / Multiplayer
# ---------------------------------------------------------------------------

@mcp.tool()
def get_network_info() -> dict:
    """Netzwerk-Info für MSFS: Ping zu MSFS-Servern, Verbindungstyp, Latenz.

    Nützlich um zu prüfen ob Live Traffic, Live Weather oder Multiplayer
    durch Netzwerkprobleme beeinträchtigt werden.

    Returns:
        ping_ms: Latenz zu MSFS Azure-Servern und DNS-Referenzzielen.
        adapter: Aktiver Netzwerk-Adapter (LAN/WLAN).
        empfehlung: LAN vs. WLAN Hinweis.
    """
    try:
        result: dict = {}

        for name, host in [
            ("azure_westeurope", "westeurope.cloudapp.azure.com"),
            ("cloudflare_dns", "1.1.1.1"),
            ("google_dns", "8.8.8.8"),
        ]:
            try:
                ping_raw = _ps_json(
                    f"Test-Connection -ComputerName '{host}' -Count 3 "
                    f"-ErrorAction SilentlyContinue "
                    f"| Select-Object @{{N='ms';E={{$_.ResponseTime}}}}"
                )
                if ping_raw:
                    times = [float(p.get("ms") or 0) for p in ping_raw if p.get("ms") is not None]
                    result[f"ping_{name}_ms"] = round(sum(times) / len(times), 1) if times else None
                else:
                    result[f"ping_{name}_ms"] = "Timeout"
            except Exception as pe:
                result[f"ping_{name}_ms"] = f"Fehler: {pe}"

        # Active adapters
        try:
            adapters = _ps_json(
                "Get-NetAdapter | Where-Object Status -eq 'Up' "
                "| Select-Object Name, InterfaceDescription, LinkSpeed, Status "
                "| Select-Object -First 3"
            )
            result["adapter"] = adapters if isinstance(adapters, list) else [adapters]
        except Exception:
            result["adapter"] = []

        # LAN vs WLAN detection
        wlan = any(
            "wi-fi" in str(a.get("Name", "")).lower()
            or "wlan" in str(a.get("Name", "")).lower()
            or "wireless" in str(a.get("InterfaceDescription", "")).lower()
            for a in (result.get("adapter") or [])
        )
        result["wlan_aktiv"] = wlan

        goog = result.get("ping_google_dns_ms")
        if isinstance(goog, float):
            if goog < 20:
                result["bewertung"] = "✅ Sehr gute Verbindung"
            elif goog < 50:
                result["bewertung"] = "✅ Gute Verbindung für MSFS Live-Dienste"
            elif goog < 100:
                result["bewertung"] = "⚠️ Mäßige Latenz — Live Weather/Traffic könnte verzögert laden"
            else:
                result["bewertung"] = "❌ Hohe Latenz — Live-Dienste eingeschränkt"
        else:
            result["bewertung"] = "Verbindungstest unvollständig"

        if wlan:
            result["wlan_hinweis"] = (
                "⚠️ WLAN aktiv. Für MSFS VR empfohlen: Ethernet (LAN-Kabel) "
                "für stabile Live Traffic, Live Weather und Multiplayer-Verbindung."
            )

        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_msfs_multiplayer_settings(
    live_traffic: str = "",
    live_weather: bool | None = None,
    multiplayer: str = "",
    usercfg_path: str = "",
) -> dict:
    """Setzt MSFS Multiplayer- und Live-Dienste-Einstellungen in der UserCfg.opt.

    Für VR kann das Reduzieren von Live Traffic und Multiplayer die CPU-Last
    verringern und Stottern eliminieren.

    Args:
        live_traffic: Traffic-Dichte: 'off', 'low' (25%), 'medium' (50%), 'high' (100%).
        live_weather: True = Live Weather an, False = Live Weather aus.
        multiplayer: 'off', 'group_only', 'live_players'.
        usercfg_path: Optional: Pfad zur UserCfg.opt.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        if not cfg_path or not cfg_path.exists():
            return {"error": "UserCfg.opt nicht gefunden. MSFS einmal starten."}

        overrides: dict = {}

        traffic_map = {"off": "0", "low": "25", "medium": "50", "high": "100"}
        if live_traffic:
            key = live_traffic.lower().strip()
            if key not in traffic_map:
                return {"error": f"Ungültiger live_traffic: '{live_traffic}'. Gültig: {list(traffic_map)}"}
            overrides["Traffic.AirTrafficDensity"] = traffic_map[key]
            overrides["Traffic.AirlineTrafficEnabled"] = "0" if key == "off" else "1"
            overrides["Traffic.GeneralAviationTrafficEnabled"] = "0" if key == "off" else "1"

        if live_weather is not None:
            overrides["Weather.WeatherDataSource"] = "1" if live_weather else "0"

        multi_map = {
            "off":          ("0", "0"),
            "group_only":   ("1", "0"),
            "live_players": ("1", "1"),
        }
        if multiplayer:
            key = multiplayer.lower().strip()
            if key not in multi_map:
                return {"error": f"Ungültiger multiplayer: '{multiplayer}'. Gültig: {list(multi_map)}"}
            mp_en, live_mp = multi_map[key]
            overrides["Multiplayer.MultiplayerEnabled"] = mp_en
            overrides["Multiplayer.LivePlayersEnabled"] = live_mp

        if not overrides:
            return {"error": "Keine Einstellungen angegeben. Nutze mind. einen Parameter."}

        snapshot_id = _snapshot(cfg_path, "Before set_msfs_multiplayer_settings")
        text = cfg_path.read_text(encoding="utf-8")
        entries = _parse_usercfg(text)
        new_entries, not_applied = _apply_overrides(entries, overrides)
        cfg_path.write_text(_entries_to_text(new_entries), encoding="utf-8")

        return {
            "status": "ok",
            "angewendet": {k: v for k, v in overrides.items() if k not in not_applied},
            "nicht_gefunden": list(not_applied),
            "snapshot": snapshot_id,
            "hinweis": (
                "Änderungen wirksam beim nächsten MSFS-Start. "
                "Für VR empfohlen: live_traffic='low'/'off', multiplayer='off' oder 'group_only'."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A8. VR Comfort Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def get_vr_comfort_settings(usercfg_path: str = "") -> dict:
    """Liest VR-Komfort-Einstellungen aus der MSFS UserCfg.opt.

    Zeigt Motion Blur, Vignette, DLSS, Schatten- und Wolken-Qualität
    mit konkreten Empfehlungen für Motion-Sickness-Reduktion.

    Returns:
        komfort_einstellungen: Aktuelle Werte + Empfehlung + Begründung.
        alle_vr_einstellungen: Alle GraphicsVR.* Werte als Referenz.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        if not cfg_path or not cfg_path.exists():
            return {"error": "UserCfg.opt nicht gefunden. MSFS einmal starten."}

        current = _read_current_settings(cfg_path)

        comfort_defs = {
            "GraphicsVR.MotionBlur": {
                "label": "Motion Blur (VR)",
                "empfehlung": "0 (aus)",
                "grund": "Motion Blur in VR verursacht häufig Motion Sickness",
                "werte": {"0": "Aus", "1": "Normal", "2": "High"},
            },
            "GraphicsVR.Vignette": {
                "label": "Vignette/Tunnelblick (VR)",
                "empfehlung": "0 (aus) — individuell",
                "grund": "Tunnelblick beim Drehen reduziert Motion Sickness, kann aber störend wirken",
                "werte": {"0": "Aus", "1": "Niedrig", "2": "Mittel", "3": "Hoch"},
            },
            "Video.DLSS": {
                "label": "DLSS",
                "empfehlung": "2 (Balanced) oder 3 (Performance)",
                "grund": "DLSS reduziert Shimmer/Aliasing in VR und erhöht FPS erheblich",
                "werte": {"0": "Aus", "1": "Quality", "2": "Balanced", "3": "Performance", "4": "Ultra Performance"},
            },
            "GraphicsVR.ShadowQuality": {
                "label": "Schatten-Qualität (VR)",
                "empfehlung": "0-1 (Niedrig/Mittel) für bessere FPS",
                "grund": "Schatten sind sehr GPU-intensiv in VR",
                "werte": {"0": "Niedrig", "1": "Mittel", "2": "Hoch", "3": "Ultra"},
            },
            "GraphicsVR.CloudsQuality": {
                "label": "Wolken-Qualität (VR)",
                "empfehlung": "1-2 (Mittel/Hoch)",
                "grund": "Ultra+ Wolken kosten in VR überproportional viel GPU",
                "werte": {"0": "Niedrig", "1": "Mittel", "2": "Hoch", "3": "Ultra", "4": "Ultra+"},
            },
            "GraphicsVR.TextureResolution": {
                "label": "Textur-Auflösung (VR)",
                "empfehlung": "2-3 (Hoch/Ultra) — VRAM abhängig",
                "grund": "Hochauflösende Texturen verbessern VR-Klarheit deutlich",
                "werte": {"0": "Niedrig", "1": "Mittel", "2": "Hoch", "3": "Ultra"},
            },
            "GraphicsVR.TerrainLoD": {
                "label": "Terrain LoD (VR)",
                "empfehlung": "80-150 je nach GPU-Tier",
                "grund": "Sehr hohe Terrain-LoD erhöht CPU-Last und kann Stottern verursachen",
                "werte": {},
            },
        }

        komfort = {}
        for key, info in comfort_defs.items():
            raw = current.get(key, "nicht gesetzt")
            werte = info.get("werte", {})
            display = werte.get(str(raw), raw) if werte else raw
            komfort[key] = {
                "label": info["label"],
                "aktuell_roh": raw,
                "aktuell_anzeige": display,
                "empfehlung": info.get("empfehlung", ""),
                "grund": info.get("grund", ""),
            }

        vr_all = {k: v for k, v in current.items() if k.startswith("GraphicsVR.")}

        return {
            "komfort_einstellungen": komfort,
            "alle_vr_einstellungen": vr_all,
            "tipp": (
                "Anti-Motion-Sickness Priorität: Motion Blur AUS, DLSS aktivieren, "
                "stabiles FPS > Frametime-Spikes vermeiden (Terrain-LoD reduzieren)."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# A9. Logging & Diagnostics
# ---------------------------------------------------------------------------

@mcp.tool()
def get_msfs_log(lines: int = 50) -> dict:
    """Liest die letzten Zeilen des MSFS Flight Simulator Logs.

    Nützlich um Abstürze, Fehler und Warnungen zu diagnostizieren.
    MSFS schreibt Logs in AppData — sucht automatisch in allen bekannten Pfaden.

    Args:
        lines: Anzahl der letzten Zeilen (Standard: 50, Max: 500).

    Returns:
        log_inhalt: Letzte N Zeilen des Logs.
        fehler_gefunden: Anzahl Fehler/Kritischer Einträge.
    """
    try:
        lines = min(max(1, lines), 500)

        log_candidates: list[Path] = []
        for base in [
            Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator 2024",
            Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
            / "Microsoft.Limitless_8wekyb3d8bbwe" / "LocalCache",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
            / "Microsoft.FlightSimulator_8wekyb3d8bbwe" / "LocalCache",
        ]:
            log_candidates.extend([
                base / "FlightSimulator.CFG",
                base / "logfile.log",
                base / "FlightSimulator.log",
            ])

        log_path: Path | None = None
        for candidate in log_candidates:
            if candidate.exists() and candidate.is_file():
                log_path = candidate
                break

        if not log_path:
            # Try Windows Event Log
            try:
                ev = _ps_json(
                    "Get-EventLog -LogName Application "
                    "-Source '*FlightSimulator*','*MSFS*' "
                    f"-Newest {min(lines, 50)} -ErrorAction SilentlyContinue "
                    "| Select-Object TimeGenerated, EntryType, Message"
                )
                if ev:
                    return {
                        "log_typ": "windows_event_log",
                        "einträge": ev,
                        "hinweis": "MSFS Log-Datei nicht gefunden. Windows Event Log gezeigt.",
                    }
            except Exception:
                pass
            return {
                "fehler": "MSFS Log-Datei nicht gefunden.",
                "hinweis": "MSFS einmal starten um Logs zu generieren.",
            }

        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            all_lines = f.readlines()
        last = all_lines[-lines:]
        log_text = "".join(last)

        error_lines = [
            l.strip() for l in last
            if any(w in l.lower() for w in ["error", "critical", "crash", "exception", "fatal"])
        ]
        warn_lines = [l.strip() for l in last if "warning" in l.lower()]

        return {
            "log_pfad": str(log_path),
            "zeilen_gelesen": len(last),
            "log_inhalt": log_text,
            "fehler_gefunden": len(error_lines),
            "fehler_zeilen": error_lines[:10],
            "warnungen_gefunden": len(warn_lines),
            "zusammenfassung": (
                f"{len(error_lines)} Fehler, {len(warn_lines)} Warnungen "
                f"in den letzten {len(last)} Log-Zeilen."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_event_log_errors(hours: int = 1, source_filter: str = "") -> dict:
    """Liest Windows Event Log Fehler der letzten N Stunden (GPU/MSFS/System).

    Findet GPU-Treiberabstürze (nvlddmkm TDR), MSFS-Fehler, Direct3D-Probleme.
    TDR-Fehler deuten auf instabile GPU-Übertaktung oder Treiberfehler hin.

    Args:
        hours: Zeitraum in Stunden (Standard: 1, Max: 24).
        source_filter: Optionaler Quell-Filter (z.B. 'nvlddmkm').

    Returns:
        gpu_abstürze_tdr: Nvidia Timeout Detection Recovery Fehler.
        msfs_fehler: MSFS-spezifische Fehler.
        bewertung: Zusammenfassung.
    """
    try:
        hours = min(max(1, hours), 24)

        extra = f"-or $_.Source -like '*{source_filter}*'" if source_filter else ""

        gpu_crashes = _ps_json(
            f"""
$since = (Get-Date).AddHours(-{hours})
Get-EventLog -LogName System -EntryType Error -Newest 100 -ErrorAction SilentlyContinue |
Where-Object {{ $_.TimeGenerated -gt $since -and
    ($_.Source -like '*nvlddmkm*' -or $_.Source -like '*dxgkrnl*' -or
     $_.Source -like '*dxgi*' -or $_.Source -like '*Display*' {extra}) }} |
Select-Object TimeGenerated, Source, EventID,
    @{{N='Message';E={{$_.Message.Substring(0,[Math]::Min(200,$_.Message.Length))}}}}
"""
        )

        app_errors = _ps_json(
            f"""
$since = (Get-Date).AddHours(-{hours})
Get-EventLog -LogName Application -EntryType Error -Newest 100 -ErrorAction SilentlyContinue |
Where-Object {{ $_.TimeGenerated -gt $since -and
    ($_.Source -like '*FlightSimulator*' -or $_.Source -like '*MSFS*' {extra}) }} |
Select-Object TimeGenerated, Source, EventID,
    @{{N='Message';E={{$_.Message.Substring(0,[Math]::Min(200,$_.Message.Length))}}}}
"""
        )

        gpu_list = gpu_crashes if isinstance(gpu_crashes, list) else []
        app_list = app_errors if isinstance(app_errors, list) else []

        if gpu_list:
            bewertung = f"❌ {len(gpu_list)} GPU-Treiberabstürze (TDR) in {hours}h — Treiber/OC prüfen!"
        elif app_list:
            bewertung = f"⚠️ {len(app_list)} MSFS-Fehler in {hours}h"
        else:
            bewertung = f"✅ Keine GPU- oder MSFS-Fehler in den letzten {hours}h"

        return {
            "zeitraum_stunden": hours,
            "gpu_abstürze_tdr": gpu_list,
            "msfs_fehler": app_list,
            "bewertung": bewertung,
            "tdr_hinweis": (
                "nvlddmkm TDR = Nvidia Timeout Detection Recovery. "
                "Ursachen: Übertaktung zu aggressiv, GPU-Unterspannung, Überhitzung, veralteter Treiber."
            ) if gpu_list else "",
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def export_vr_diagnostic_report(output_path: str = "") -> dict:
    """Exportiert vollständigen VR-Diagnose-Bericht als HTML auf den Desktop.

    Sammelt Hardware, MSFS-Einstellungen, Pimax, OpenXR, SteamVR, Netzwerk
    und Event-Log-Fehler und schreibt alles in eine übersichtliche HTML-Datei.

    Args:
        output_path: Optionaler Ausgabepfad. Standard: Desktop/VR_Diagnose_<datum>.html.

    Returns:
        bericht_pfad: Pfad zur erzeugten HTML-Datei.
        abschnitte: Gesammelte Datenbereiche.
    """
    try:
        now = datetime.datetime.now()
        date_str = now.strftime("%Y-%m-%d_%H-%M")
        if not output_path:
            output_path = str(Path.home() / "Desktop" / f"VR_Diagnose_{date_str}.html")

        out = Path(output_path)
        sections: dict = {}
        errs: list[str] = []

        def _c(name: str, fn):
            try:
                sections[name] = fn()
            except Exception as exc:
                sections[name] = {"error": str(exc)}
                errs.append(f"{name}: {exc}")

        _c("hardware", get_detailed_hardware_profile)
        _c("cpu", get_cpu_info)
        _c("ram", get_ram_info)
        _c("temperaturen", get_system_temps)
        _c("openxr", get_openxr_runtime)
        _c("steamvr", get_steamvr_settings)
        _c("pimax", get_pimax_headset_info)
        _c("render_auflösung", get_vr_render_resolution)
        _c("netzwerk", get_network_info)
        _c("msfs_prozess", get_msfs_process_info)
        _c("community", get_msfs_community_folder)
        _c("event_errors_24h", lambda: get_event_log_errors(hours=24))
        _c("vr_diagnose", diagnose_vr_complete)

        def _sec(title: str, data: dict) -> str:
            rows = ""
            for k, v in data.items():
                if isinstance(v, (dict, list)):
                    v_str = _json.dumps(v, ensure_ascii=False, indent=2)
                    rows += f"<tr><td><b>{k}</b></td><td><pre>{v_str}</pre></td></tr>\n"
                else:
                    rows += f"<tr><td><b>{k}</b></td><td>{v}</td></tr>\n"
            return (
                f'<div class="section"><h2>{title}</h2>'
                f"<table><tbody>{rows}</tbody></table></div>"
            )

        body = "\n".join(_sec(n, d) for n, d in sections.items())

        html = f"""<!DOCTYPE html>
<html lang="de"><head>
<meta charset="UTF-8">
<title>Game Copilot VR-Diagnose {date_str}</title>
<style>
  body{{font-family:'Segoe UI',Arial,sans-serif;background:#1a1a2e;color:#e0e0e0;padding:20px}}
  h1{{color:#4ade80;border-bottom:2px solid #4ade80;padding-bottom:8px}}
  h2{{color:#69daff;margin-top:24px}}
  .section{{background:#16213e;border-radius:8px;padding:16px;margin-bottom:16px;border:1px solid #0f3460}}
  table{{width:100%;border-collapse:collapse}}
  td{{padding:5px 10px;border-bottom:1px solid #0f3460;vertical-align:top}}
  td:first-child{{width:220px;color:#a0aec0;white-space:nowrap}}
  pre{{font-size:11px;white-space:pre-wrap;margin:0;color:#cbd5e0}}
  .footer{{text-align:center;color:#4a5568;margin-top:32px;font-size:12px}}
</style></head><body>
<h1>🎮 Game Copilot — VR-Diagnose-Bericht</h1>
<p>Erstellt: <b>{now.strftime("%d.%m.%Y %H:%M:%S")}</b> &nbsp;|&nbsp; Version: <b>3.6.0</b></p>
{body}
<div class="footer">Erstellt von Game Copilot MCP Server v3.6.0</div>
</body></html>"""

        out.write_text(html, encoding="utf-8")

        return {
            "status": "ok",
            "bericht_pfad": str(out),
            "erstellt_am": now.strftime("%d.%m.%Y %H:%M"),
            "abschnitte": list(sections.keys()),
            "fehler_beim_sammeln": errs if errs else "Keine",
            "hinweis": f"HTML-Bericht gespeichert: {out}",
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# B. get_msfs_current_settings — Read ALL current MSFS settings
# ---------------------------------------------------------------------------

@mcp.tool()
def get_msfs_current_settings(usercfg_path: str = "", section_filter: str = "") -> dict:
    """Liest ALLE aktuellen MSFS Grafik- und VR-Einstellungen aus der UserCfg.opt.

    Essentiell für Before/After-Vergleiche nach Optimierungen.
    Gibt alle Schlüssel mit Rohwert, lesbarem Wert und Label zurück.

    Args:
        usercfg_path: Optional: Pfad zur UserCfg.opt.
        section_filter: Filter nach Kategorie-Präfix (z.B. 'GraphicsVR', 'Video').
                        Leer = alle Einstellungen.

    Returns:
        einstellungen: Alle Einstellungen gruppiert nach Kategorie.
        cfg_pfad, letztes_änderungsdatum, gesamt_einstellungen.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        if not cfg_path or not cfg_path.exists():
            return {
                "error": "UserCfg.opt nicht gefunden.",
                "hinweis": "Starte MSFS einmal um die Konfigurationsdatei zu erstellen.",
            }

        current = _read_current_settings(cfg_path)
        grouped: dict[str, dict] = {}

        for key, value in sorted(current.items()):
            if section_filter and not key.lower().startswith(section_filter.lower()):
                continue
            parts = key.split(".", 1)
            group = parts[0] if len(parts) > 1 else "Sonstige"
            if group not in grouped:
                grouped[group] = {}
            defn = SETTING_DEFS.get(key, {})
            grouped[group][key] = {
                "roh": value,
                "anzeige": defn.get("values", {}).get(str(value), str(value)),
                "label": defn.get("label", key),
            }

        stat = cfg_path.stat()
        mod_time = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M:%S")

        return {
            "cfg_pfad": str(cfg_path),
            "letztes_änderungsdatum": mod_time,
            "gesamt_einstellungen": len(current),
            "kategorien": list(grouped.keys()),
            "einstellungen": grouped,
            "hinweis": (
                f"Alle {len(current)} MSFS-Einstellungen gelesen. "
                "section_filter='GraphicsVR' für nur VR-Einstellungen."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# NEW TOOLS — v3.6.1 additions
# ---------------------------------------------------------------------------


@mcp.tool()
def get_msfs_weather_settings(usercfg_path: str = "") -> dict:
    """Liest aktuelle MSFS Wetter-Einstellungen aus der UserCfg.opt.

    Wann aufrufen: Wenn der User fragt ob Live Weather aktiv ist, ob Wolken
    aktiviert sind, oder welche Wetter-Konfiguration MSFS gerade hat.
    Typische Trigger: 'Wetter', 'Live Weather', 'Wolken', 'Weather'.

    Args:
        usercfg_path: Optional: Pfad zur UserCfg.opt. Leer = automatische Suche.

    Returns:
        live_weather_aktiv: bool — True wenn Live Weather aus MSFS-Servern.
        wetter_datenquelle: '0' = statisch, '1' = Live, '2' = manuell.
        wolken_qualitaet: Aktueller Wert der Wolken-Qualitätsstufe.
        wolken_menge: Dichte/Menge der Wolken.
        sichtweite_km: Konfigurierte Sichtweite in km.
        hinweis: Empfehlung für VR-Performance.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        if not cfg_path or not cfg_path.exists():
            return {
                "error": "UserCfg.opt nicht gefunden.",
                "hinweis": "MSFS einmal starten um die Konfiguration zu erstellen.",
            }

        current = _read_current_settings(cfg_path)

        data_source = current.get("Weather.WeatherDataSource", "1")
        live_weather = data_source == "1"
        cloud_quality = current.get("GraphicsVR.CloudQuality",
                                    current.get("Graphics.CloudQuality", "N/A"))
        cloud_draw_distance = current.get("Weather.CloudDrawDistance", "N/A")
        visibility_km = current.get("Weather.VisibilityDistance", "N/A")
        wind_effects = current.get("Weather.WindEffectsEnabled", "N/A")

        source_labels = {"0": "Statisch (kein Live)", "1": "Live Weather (Server)", "2": "Manuell/Custom"}
        quality_labels = {"0": "Aus", "1": "Niedrig", "2": "Mittel", "3": "Hoch", "4": "Ultra"}

        vr_hinweis = (
            "Für VR-Performance: Live Weather (1) erhöht CPU-Last. "
            "Wolken-Qualität 1-2 empfohlen für VR. "
            "Cloud Draw Distance reduzieren verbessert Framerate."
        )

        return {
            "live_weather_aktiv": live_weather,
            "wetter_datenquelle_roh": data_source,
            "wetter_datenquelle": source_labels.get(data_source, data_source),
            "wolken_qualitaet_roh": cloud_quality,
            "wolken_qualitaet": quality_labels.get(str(cloud_quality), str(cloud_quality)),
            "wolken_sichtweite": cloud_draw_distance,
            "sichtweite_km": visibility_km,
            "wind_effekte": wind_effects,
            "cfg_pfad": str(cfg_path),
            "hinweis": vr_hinweis,
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_msfs_weather_settings(
    live_weather: bool | None = None,
    cloud_quality: int | None = None,
    usercfg_path: str = "",
) -> dict:
    """Setzt MSFS Wetter-Einstellungen in der UserCfg.opt.

    Wann aufrufen: Wenn der User Live Weather ein-/ausschalten will,
    die Wolkenqualität für VR-Performance optimieren möchte, oder explizit
    nach Wetter-Optimierung fragt.
    Trigger: 'Live Weather aus', 'Wolken reduzieren', 'Wetter optimieren'.

    NUR nach expliziter User-Bestätigung aufrufen (schreibt Konfiguration).

    Args:
        live_weather: True = Live Weather aktivieren, False = deaktivieren.
        cloud_quality: Wolken-Qualität 0-4 (0=Aus, 1=Niedrig, 2=Mittel, 3=Hoch, 4=Ultra).
                       Für VR empfohlen: 1 oder 2.
        usercfg_path: Optional: Pfad zur UserCfg.opt.

    Returns:
        status, angewendet, snapshot (für Rollback), hinweis.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        if not cfg_path or not cfg_path.exists():
            return {"error": "UserCfg.opt nicht gefunden. MSFS einmal starten."}

        overrides: dict = {}

        if live_weather is not None:
            overrides["Weather.WeatherDataSource"] = "1" if live_weather else "0"

        if cloud_quality is not None:
            if not 0 <= cloud_quality <= 4:
                return {"error": "cloud_quality muss zwischen 0 und 4 liegen."}
            overrides["GraphicsVR.CloudQuality"] = str(cloud_quality)
            overrides["Graphics.CloudQuality"] = str(cloud_quality)

        if not overrides:
            return {"error": "Keine Parameter angegeben. Nutze live_weather oder cloud_quality."}

        snapshot_id = _snapshot(cfg_path, "Before set_msfs_weather_settings")
        text = cfg_path.read_text(encoding="utf-8")
        entries = _parse_usercfg(text)
        new_entries, not_applied = _apply_overrides(entries, overrides)
        cfg_path.write_text(_entries_to_text(new_entries), encoding="utf-8")

        quality_labels = {"0": "Aus", "1": "Niedrig", "2": "Mittel", "3": "Hoch", "4": "Ultra"}
        return {
            "status": "ok",
            "angewendet": {k: v for k, v in overrides.items() if k not in not_applied},
            "nicht_gefunden": list(not_applied),
            "snapshot_id": snapshot_id,
            "live_weather": ("Aktiv" if live_weather else "Deaktiviert") if live_weather is not None else "Unverändert",
            "wolken_qualitaet": quality_labels.get(str(cloud_quality), str(cloud_quality)) if cloud_quality is not None else "Unverändert",
            "hinweis": "Änderungen wirksam beim nächsten MSFS-Start. Rollback: restore_msfs_graphics.",
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_msfs_traffic_settings(usercfg_path: str = "") -> dict:
    """Liest aktuelle MSFS Traffic-Einstellungen (KI-Flugzeuge, Online-Traffic).

    Wann aufrufen: Wenn der User nach Traffic-Dichte, KI-Flugzeugen,
    Online-Multiplayer-Flugzeugen fragt, oder wenn VR-Performance durch
    Traffic limitiert sein könnte.
    Trigger: 'Traffic', 'KI Flugzeuge', 'AI traffic', 'Multiplayer Dichte'.

    Args:
        usercfg_path: Optional: Pfad zur UserCfg.opt.

    Returns:
        traffic_dichte: Prozentsatz KI-Flugverkehr (0-100).
        airline_traffic: bool — Airline-KI aktiv.
        ga_traffic: bool — Kleinflugzeug-KI aktiv.
        multiplayer_aktiv: bool.
        live_players: bool — Online-Spieler sichtbar.
        vr_empfehlung: Empfohlene Werte für VR-Performance.
    """
    try:
        cfg_path = _find_usercfg(usercfg_path)
        if not cfg_path or not cfg_path.exists():
            return {
                "error": "UserCfg.opt nicht gefunden.",
                "hinweis": "MSFS einmal starten.",
            }

        current = _read_current_settings(cfg_path)

        traffic_density = current.get("Traffic.AirTrafficDensity", "50")
        airline_enabled = current.get("Traffic.AirlineTrafficEnabled", "1") == "1"
        ga_enabled = current.get("Traffic.GeneralAviationTrafficEnabled", "1") == "1"
        mp_enabled = current.get("Multiplayer.MultiplayerEnabled", "1") == "1"
        live_players = current.get("Multiplayer.LivePlayersEnabled", "1") == "1"
        ground_density = current.get("Traffic.GroundAircraftDensity", "N/A")
        boat_density = current.get("Traffic.BoatTrafficDensity", "N/A")

        try:
            density_int = int(traffic_density)
            if density_int == 0:
                traffic_level = "Aus (VR-optimal)"
            elif density_int <= 25:
                traffic_level = "Niedrig (gut für VR)"
            elif density_int <= 50:
                traffic_level = "Mittel (akzeptabel)"
            else:
                traffic_level = "Hoch (CPU-intensiv, nicht VR-optimal)"
        except ValueError:
            traffic_level = traffic_density

        return {
            "traffic_dichte_prozent": traffic_density,
            "traffic_level": traffic_level,
            "airline_traffic_aktiv": airline_enabled,
            "ga_traffic_aktiv": ga_enabled,
            "boden_traffic": ground_density,
            "boot_traffic": boat_density,
            "multiplayer_aktiv": mp_enabled,
            "live_players_sichtbar": live_players,
            "cfg_pfad": str(cfg_path),
            "vr_empfehlung": (
                "Für VR: traffic_dichte 0-25%, airline/GA traffic aus oder niedrig, "
                "Multiplayer auf 'group_only'. Reduziert CPU-Last signifikant."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_vr_headset_connected() -> dict:
    """Erkennt ob und welches VR-Headset gerade verbunden/aktiv ist.

    Wann aufrufen: Wenn der User fragt ob sein VR-Headset erkannt wird,
    bevor VR gestartet wird, oder bei VR-Verbindungsproblemen.
    Trigger: 'Headset erkannt', 'VR Headset verbunden', 'VR nicht erkannt',
    'Pimax gefunden', 'SteamVR status'.

    Returns:
        headset_gefunden: bool.
        headset_typ: 'Pimax', 'SteamVR Generic', 'WMR', 'Oculus/Meta', 'Unbekannt'.
        aktive_prozesse: Liste erkannter VR-Prozesse.
        openxr_runtime: Aktuell registrierte OpenXR-Runtime (aus Registry).
        empfehlung: Nächster Schritt wenn kein Headset gefunden.
    """
    import subprocess

    vr_processes = {
        # Pimax
        "pimax_client.exe":    "Pimax Client",
        "pimax_runtime.exe":   "Pimax Runtime",
        "pimaxclient.exe":     "Pimax Client",
        "pimaxruntime.exe":    "Pimax Runtime",
        "pvrservice.exe":      "Pimax VR Service",
        # SteamVR
        "vrserver.exe":        "SteamVR Server",
        "vrstartup.exe":       "SteamVR Startup",
        "vrcompositor.exe":    "SteamVR Compositor",
        "vrmonitor.exe":       "SteamVR Monitor",
        # WMR / OpenXR Tools
        "mixedreality.exe":    "WMR Portal",
        "holographicshell.exe":"WMR Shell",
        # Meta / Oculus
        "oculusclient.exe":    "Meta/Oculus Client",
        "ovrserver.exe":       "Oculus VR Server",
        "ovrservice.exe":      "Oculus Service",
        # OpenXR
        "openxrexplorer.exe":  "OpenXR Explorer",
    }

    found: list[dict] = []

    try:
        result = subprocess.run(
            ["tasklist", "/fo", "csv", "/nh"],
            capture_output=True, text=True, timeout=5
        )
        running = {line.split(",")[0].strip('"').lower() for line in result.stdout.splitlines() if line}
        for proc, label in vr_processes.items():
            if proc in running:
                found.append({"prozess": proc, "bezeichnung": label})
    except Exception:
        pass

    # Detect headset family
    names = {f["bezeichnung"] for f in found}
    if any("Pimax" in n for n in names):
        headset_typ = "Pimax"
    elif any("SteamVR" in n for n in names):
        headset_typ = "SteamVR Generic"
    elif any("WMR" in n for n in names):
        headset_typ = "Windows Mixed Reality"
    elif any("Oculus" in n or "Meta" in n for n in names):
        headset_typ = "Meta/Oculus"
    else:
        headset_typ = "Nicht erkannt"

    # Read active OpenXR runtime from registry
    openxr_runtime = "Unbekannt"
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Khronos\OpenXR\1",
        )
        val, _ = winreg.QueryValueEx(key, "ActiveRuntime")
        openxr_runtime = str(val)
        winreg.CloseKey(key)
    except Exception:
        pass

    headset_gefunden = len(found) > 0

    empfehlung = ""
    if not headset_gefunden:
        empfehlung = (
            "Kein VR-Headset erkannt. Prüfe: "
            "1) Pimax Client starten (Startmenü → Pimax), "
            "2) SteamVR starten (Steam → SteamVR), "
            "3) USB-Verbindung und Treiber prüfen."
        )

    return {
        "headset_gefunden": headset_gefunden,
        "headset_typ": headset_typ,
        "aktive_vr_prozesse": found,
        "anzahl_prozesse": len(found),
        "openxr_runtime": openxr_runtime,
        "empfehlung": empfehlung,
    }


@mcp.tool()
def restart_steamvr() -> dict:
    """Beendet und startet SteamVR neu.

    Wann aufrufen: Wenn SteamVR abgestürzt ist, eingefroren ist,
    der User 'SteamVR neustarten' sagt, oder nach VR-Einstellungsänderungen
    die einen Neustart erfordern.
    Trigger: 'SteamVR neustart', 'SteamVR hängt', 'SteamVR neu starten'.

    NUR nach expliziter User-Bestätigung aufrufen — beendet laufende VR-Session.

    Returns:
        status, beendete_prozesse, gestartet, steamvr_pfad, hinweis.
    """
    import subprocess
    import time

    steamvr_processes = [
        "vrserver.exe",
        "vrcompositor.exe",
        "vrdashboard.exe",
        "vrmonitor.exe",
        "vrstartup.exe",
        "vrwebhelper.exe",
        "vrservice.exe",
    ]

    # Step 1: Kill all SteamVR processes
    killed = []
    for proc in steamvr_processes:
        try:
            result = subprocess.run(
                ["taskkill", "/f", "/im", proc],
                capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                killed.append(proc)
        except Exception:
            pass

    if killed:
        time.sleep(2)  # Give OS time to clean up

    # Step 2: Find SteamVR startup executable
    steamvr_paths = [
        r"C:\Program Files (x86)\Steam\steamapps\common\SteamVR\bin\win64\vrstartup.exe",
        r"C:\Program Files\Steam\steamapps\common\SteamVR\bin\win64\vrstartup.exe",
    ]

    # Check registry for Steam install path
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Valve\Steam")
        steam_path, _ = winreg.QueryValueEx(key, "InstallPath")
        winreg.CloseKey(key)
        registry_path = os.path.join(
            steam_path, "steamapps", "common", "SteamVR", "bin", "win64", "vrstartup.exe"
        )
        steamvr_paths.insert(0, registry_path)
    except Exception:
        pass

    vrstartup = next((p for p in steamvr_paths if os.path.exists(p)), None)

    # Step 3: Start SteamVR
    started = False
    if vrstartup:
        try:
            subprocess.Popen(
                [vrstartup],
                creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            )
            started = True
        except Exception as start_err:
            return {
                "status": "error",
                "beendete_prozesse": killed,
                "fehler": f"Konnte SteamVR nicht starten: {start_err}",
                "steamvr_pfad": vrstartup,
            }
    else:
        return {
            "status": "warning",
            "beendete_prozesse": killed,
            "gestartet": False,
            "hinweis": (
                "SteamVR-Prozesse beendet, aber vrstartup.exe nicht gefunden. "
                "SteamVR manuell über Steam starten."
            ),
        }

    return {
        "status": "ok",
        "beendete_prozesse": killed,
        "anzahl_beendet": len(killed),
        "gestartet": started,
        "steamvr_pfad": vrstartup,
        "hinweis": "SteamVR wird neu gestartet. Bitte 10-15 Sekunden warten.",
    }


@mcp.tool()
def get_gpu_overclock_status() -> dict:
    """Prüft ob die GPU aktuell übertaktet ist (via MSI Afterburner Registry oder NVML).

    Wann aufrufen: Wenn der User fragt ob seine GPU übertaktet ist,
    bei GPU-Stabilitätsproblemen, oder vor Treiberinstallation.
    Trigger: 'Übertaktung', 'OC Status', 'GPU Boost', 'Afterburner', 'overclock'.

    Returns:
        übertaktet: bool.
        kern_offset_mhz: Aktueller Kern-Takt-Offset (MSI Afterburner).
        speicher_offset_mhz: Speicher-Takt-Offset.
        afterburner_gefunden: bool — MSI Afterburner installiert.
        basis_takt_mhz: GPU Basistakt (NVML).
        aktueller_takt_mhz: Aktuell gemessener Takt (NVML).
        hinweis: Empfehlung für Stabilität.
    """
    afterburner_found = False
    core_offset = 0
    mem_offset = 0
    ab_profile = "N/A"

    # Check MSI Afterburner via Registry
    try:
        import winreg
        ab_key_path = r"SOFTWARE\WOW6432Node\MSI\Afterburner"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, ab_key_path)
        afterburner_found = True

        # Try to read current overclocking profile
        try:
            core_offset_val, _ = winreg.QueryValueEx(key, "CoreClockOffset")
            core_offset = int(core_offset_val)
        except Exception:
            pass
        try:
            mem_offset_val, _ = winreg.QueryValueEx(key, "MemoryClockOffset")
            mem_offset = int(mem_offset_val)
        except Exception:
            pass
        try:
            profile_val, _ = winreg.QueryValueEx(key, "LastUsedProfile")
            ab_profile = str(profile_val)
        except Exception:
            pass

        winreg.CloseKey(key)
    except Exception:
        pass

    # Get base and actual clocks via NVML
    base_clock = 0
    current_clock = 0
    vram_clock = 0
    nvml_error = ""

    try:
        import pynvml  # type: ignore
        pynvml.nvmlInit()
        handle = pynvml.nvmlDeviceGetHandleByIndex(0)

        try:
            base_clock = pynvml.nvmlDeviceGetClockInfo(handle, pynvml.NVML_CLOCK_GRAPHICS)
        except Exception:
            pass
        try:
            current_clock = pynvml.nvmlDeviceGetClockInfo(handle, pynvml.NVML_CLOCK_SM)
        except Exception:
            pass
        try:
            vram_clock = pynvml.nvmlDeviceGetClockInfo(handle, pynvml.NVML_CLOCK_MEM)
        except Exception:
            pass

        pynvml.nvmlShutdown()
    except Exception as e:
        nvml_error = str(e)

    # Determine OC status
    is_overclocked = (
        (afterburner_found and (abs(core_offset) > 0 or abs(mem_offset) > 0))
    )

    if is_overclocked:
        hinweis = (
            f"GPU ist übertaktet: Core +{core_offset} MHz, Mem +{mem_offset} MHz. "
            "Bei Instabilität (Abstürze, TDR) OC in MSI Afterburner zurücksetzen."
        )
    elif afterburner_found:
        hinweis = "MSI Afterburner gefunden, aber kein aktiver Offset erkannt. GPU läuft auf Stock-Takten."
    else:
        hinweis = (
            "MSI Afterburner nicht gefunden. "
            "GPU-Übertaktungsstatus kann nicht ohne Afterburner geprüft werden. "
            "NVML-Taktdaten zeigen Echtzeit-Boost-Frequenz."
        )

    return {
        "übertaktet": is_overclocked,
        "afterburner_gefunden": afterburner_found,
        "kern_offset_mhz": core_offset,
        "speicher_offset_mhz": mem_offset,
        "afterburner_profil": ab_profile,
        "aktueller_kern_takt_mhz": current_clock,
        "aktueller_speicher_takt_mhz": vram_clock,
        "nvml_fehler": nvml_error if nvml_error else None,
        "hinweis": hinweis,
    }


@mcp.tool()
def set_windows_theme(dark: bool = True) -> dict:
    """Schaltet Windows Dark Mode / Light Mode um.

    Wann aufrufen: Wenn der User Windows auf Dark-Mode oder Light-Mode umschalten möchte.
    Trigger: 'Dark Mode', 'Light Mode', 'Windows Farbe', 'Thema ändern'.

    Args:
        dark: True = Dark Mode aktivieren, False = Light Mode aktivieren.

    Returns:
        status: "ok" | "error"
        modus: "dark" | "light"
        hinweis: Erklärung.
    """
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        value = 0 if dark else 1  # 0 = Dark, 1 = Light
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, key_path,
            0, winreg.KEY_SET_VALUE
        ) as key:
            winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, value)
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, value)
        return {
            "status": "ok",
            "modus": "dark" if dark else "light",
            "hinweis": (
                "Dark Mode aktiviert. Einige Apps benötigen einen Neustart."
                if dark else
                "Light Mode aktiviert. Einige Apps benötigen einen Neustart."
            ),
        }
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def get_msfs_installed_version() -> dict:
    """Liest die installierte MSFS 2024-Version aus dem System (Registry oder EXE).

    Wann aufrufen: Wenn der User fragt welche MSFS-Version installiert ist,
    oder vor einem Update-Check.
    Trigger: 'MSFS Version', 'welche Version', 'FS2024 installiert'.

    Returns:
        version: Versionsnummer als String, z.B. "1.5.12.0".
        quelle: Woher die Version stammt ("registry" | "exe" | "unbekannt").
        pfad: Installationspfad (falls gefunden).
    """
    version = "unbekannt"
    quelle = "unbekannt"
    pfad = ""

    # 1. Try Steam registry
    try:
        import winreg
        steam_apps_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, steam_apps_key) as base:
            for i in range(winreg.QueryInfoKey(base)[0]):
                try:
                    sub_name = winreg.EnumKey(base, i)
                    with winreg.OpenKey(base, sub_name) as sub:
                        try:
                            display_name, _ = winreg.QueryValueEx(sub, "DisplayName")
                            if "Microsoft Flight Simulator" in str(display_name) and "2024" in str(display_name):
                                try:
                                    v, _ = winreg.QueryValueEx(sub, "DisplayVersion")
                                    version = str(v)
                                    quelle = "registry"
                                except Exception:
                                    pass
                                try:
                                    loc, _ = winreg.QueryValueEx(sub, "InstallLocation")
                                    pfad = str(loc)
                                except Exception:
                                    pass
                                break
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass

    # 2. Try reading version from the EXE
    if version == "unbekannt":
        import os
        candidates = [
            r"C:\XboxGames\Microsoft Flight Simulator 2024\Content\FlightSimulator.exe",
            r"C:\Program Files\WindowsApps\Microsoft.Limitless_1.0.0.0_x64__8wekyb3d8bbwe\FlightSimulator.exe",
        ]
        # Check Steam path from registry
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\WOW6432Node\Valve\Steam"
            ) as k:
                steam_path, _ = winreg.QueryValueEx(k, "InstallPath")
                candidates.append(os.path.join(
                    str(steam_path),
                    "steamapps", "common",
                    "Microsoft Flight Simulator 2024",
                    "FlightSimulator.exe"
                ))
        except Exception:
            pass

        for exe in candidates:
            if os.path.isfile(exe):
                pfad = exe
                try:
                    import win32api  # type: ignore
                    info = win32api.GetFileVersionInfo(exe, "\\")
                    ms = info["FileVersionMS"]
                    ls = info["FileVersionLS"]
                    version = f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
                    quelle = "exe"
                except Exception:
                    version = "gefunden (Version nicht lesbar)"
                    quelle = "exe"
                break

    return {
        "version": version,
        "quelle": quelle,
        "pfad": pfad,
        "hinweis": "MSFS 2024 Version erfolgreich gelesen." if version != "unbekannt"
                   else "Version konnte nicht ermittelt werden. MSFS möglicherweise nicht installiert.",
    }


@mcp.tool()
def set_cpu_priority_msfs(priority: str = "high") -> dict:
    """Setzt die CPU-Prozesspriorität von MSFS auf High oder Realtime.

    Wann aufrufen: Wenn der User mehr FPS will oder MSFS auf niedrige
    CPU-Priorität eingestellt ist. Nur wenn MSFS läuft.
    Trigger: 'MSFS Priorität', 'CPU Priorität erhöhen', 'High Priority'.

    Args:
        priority: "high" (empfohlen) | "realtime" (vorsichtig verwenden) | "normal".

    Returns:
        status: "ok" | "error"
        prioritaet: Gesetzte Priorität.
        pid: Prozess-ID von MSFS.
    """
    import subprocess
    priority_map = {
        "realtime": "Realtime",
        "high":     "High",
        "normal":   "Normal",
        "abovenormal": "AboveNormal",
    }
    wmi_priority = priority_map.get(priority.lower(), "High")

    # Find MSFS process
    ps_find = (
        "Get-Process | Where-Object { $_.Name -match 'FlightSimulator' } "
        "| Select-Object -First 1 -ExpandProperty Id"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_find],
            capture_output=True, text=True, timeout=10
        )
        pid_str = result.stdout.strip()
        if not pid_str:
            return {
                "status": "error",
                "error": "MSFS läuft nicht. Bitte MSFS zuerst starten.",
            }
        pid = int(pid_str)
    except Exception as e:
        return {"status": "error", "error": f"Prozess-Suche fehlgeschlagen: {e}"}

    # Set priority via WMI
    ps_set = (
        f"$proc = Get-Process -Id {pid}; "
        f"$proc.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::{wmi_priority}; "
        f"Write-Output 'OK'"
    )
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_set],
            capture_output=True, text=True, timeout=10
        )
        if "OK" in res.stdout:
            return {
                "status": "ok",
                "prioritaet": wmi_priority,
                "pid": pid,
                "hinweis": (
                    f"MSFS CPU-Priorität auf {wmi_priority} gesetzt (PID {pid}). "
                    "Gilt bis zum nächsten MSFS-Neustart."
                ),
            }
        return {"status": "error", "error": res.stderr.strip() or "Unbekannter Fehler"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def get_msfs_addon_count(community_path: str = "") -> dict:
    """Zählt die Anzahl der Add-ons im MSFS Community-Ordner (schnelle Version).

    Wann aufrufen: Wenn der User wissen will wie viele Add-ons installiert sind,
    ohne die komplette Liste zu laden.
    Trigger: 'Addon Anzahl', 'wie viele Mods', 'Community Ordner Größe'.

    Args:
        community_path: Optionaler Pfad zum Community-Ordner. Wird automatisch erkannt falls leer.

    Returns:
        anzahl: Anzahl der Add-on-Ordner.
        pfad: Genutzter Community-Ordner-Pfad.
        gesamt_groesse_mb: Gesamtgröße in MB.
    """
    import os

    if not community_path:
        # Auto-detect community folder
        candidates = [
            os.path.join(os.environ.get("APPDATA", ""), "Microsoft Flight Simulator 2024", "Packages", "Community"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft Flight Simulator 2024", "Packages", "Community"),
        ]
        community_path = next((p for p in candidates if os.path.isdir(p)), "")

    if not community_path or not os.path.isdir(community_path):
        return {
            "status": "error",
            "error": "Community-Ordner nicht gefunden.",
            "hinweis": "Bitte community_path angeben.",
        }

    try:
        entries = [e for e in os.scandir(community_path) if e.is_dir()]
        anzahl = len(entries)

        # Quick size estimate (only top-level)
        total_bytes = sum(e.stat().st_size for e in os.scandir(community_path))
        groesse_mb = round(total_bytes / (1024 * 1024), 1)
    except Exception as e:
        return {"status": "error", "error": str(e)}

    return {
        "status": "ok",
        "anzahl": anzahl,
        "pfad": community_path,
        "gesamt_groesse_mb": groesse_mb,
        "hinweis": f"{anzahl} Add-ons im Community-Ordner gefunden.",
    }


@mcp.tool()
def set_steamvr_supersampling(value: float = 1.0) -> dict:
    """Setzt den SteamVR Supersampling-Wert (renderResolution / supersampleScale).

    Wann aufrufen: Wenn der User direkt die Renderauflösung in SteamVR ändern will.
    Trigger: 'SteamVR Supersampling', 'SS Wert', 'Renderauflösung SteamVR'.

    Args:
        value: Supersampling-Faktor, z.B. 1.0 = 100%, 1.5 = 150%, 2.0 = 200%.
               Empfohlen: 1.0–1.5 für Pimax MSFS VR.

    Returns:
        status: "ok" | "error"
        neuer_wert: Gesetzter SS-Wert.
    """
    return set_steamvr_setting(setting="supersampling", value=str(value))


@mcp.tool()
def get_pimax_play_version() -> dict:
    """Liest die installierte Pimax Play / PiTool Version aus Registry oder EXE.

    Wann aufrufen: Wenn der User nach der Pimax-Software-Version fragt,
    oder vor Firmware-Updates.
    Trigger: 'Pimax Play Version', 'PiTool Version', 'Pimax Software'.

    Returns:
        version: Versionsnummer als String.
        software: "Pimax Play" | "PiTool" | "unbekannt".
        pfad: Installationspfad.
    """
    import os

    version = "unbekannt"
    software = "unbekannt"
    pfad = ""

    try:
        import winreg

        # 1. Try Pimax Play
        for key_path in [
            r"SOFTWARE\Pimax\PimaxPlay",
            r"SOFTWARE\WOW6432Node\Pimax\PimaxPlay",
        ]:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as k:
                    try:
                        v, _ = winreg.QueryValueEx(k, "Version")
                        version = str(v)
                    except Exception:
                        pass
                    try:
                        loc, _ = winreg.QueryValueEx(k, "InstallDir")
                        pfad = str(loc)
                    except Exception:
                        pass
                    software = "Pimax Play"
                    break
            except Exception:
                pass

        # 2. Try PiTool
        if software == "unbekannt":
            for key_path in [
                r"SOFTWARE\Pimax\PiTool",
                r"SOFTWARE\WOW6432Node\Pimax\PiTool",
            ]:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as k:
                        try:
                            v, _ = winreg.QueryValueEx(k, "Version")
                            version = str(v)
                        except Exception:
                            pass
                        try:
                            loc, _ = winreg.QueryValueEx(k, "InstallDir")
                            pfad = str(loc)
                        except Exception:
                            pass
                        software = "PiTool"
                        break
                except Exception:
                    pass
    except Exception:
        pass

    # 3. Fallback: check common EXE paths
    if software == "unbekannt":
        exe_candidates = [
            r"C:\Program Files\Pimax\PimaxPlay\PimaxPlay.exe",
            r"C:\Program Files\Pimax\PimaxClient\pimaxui\PimaxClient.exe",
            r"C:\Program Files (x86)\PiTool\PiServiceLauncher.exe",
        ]
        for exe in exe_candidates:
            if os.path.isfile(exe):
                pfad = exe
                software = "Pimax Play" if "PimaxPlay" in exe else "PiTool"
                version = "gefunden (Version nicht lesbar)"
                break

    return {
        "version": version,
        "software": software,
        "pfad": pfad,
        "hinweis": f"{software} {version} gefunden." if software != "unbekannt"
                   else "Keine Pimax-Software gefunden. Pimax Play oder PiTool installieren.",
    }


@mcp.tool()
def take_vr_screenshot() -> dict:
    """Löst einen MSFS-Screenshot per Tastendruck aus (F12 / PrintScreen).

    Wann aufrufen: Wenn der User einen Screenshot machen möchte während MSFS läuft.
    Trigger: 'Screenshot', 'Foto machen', 'Bild aufnehmen'.

    Returns:
        status: "ok" | "error"
        taste: Genutzte Taste.
        hinweis: Wo der Screenshot gespeichert wird.
    """
    import subprocess

    # Check if MSFS is running first
    ps_check = "Get-Process | Where-Object { $_.Name -match 'FlightSimulator' } | Measure-Object | Select-Object -ExpandProperty Count"
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_check],
            capture_output=True, text=True, timeout=8
        )
        count = int(r.stdout.strip() or "0")
        if count == 0:
            return {
                "status": "error",
                "error": "MSFS läuft nicht. Screenshot nicht möglich.",
            }
    except Exception as e:
        return {"status": "error", "error": f"Prozess-Check fehlgeschlagen: {e}"}

    # Send PrintScreen key to the focused window
    ps_key = (
        "Add-Type -AssemblyName System.Windows.Forms; "
        "[System.Windows.Forms.SendKeys]::SendWait('{PRTSC}'); "
        "Write-Output 'OK'"
    )
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_key],
            capture_output=True, text=True, timeout=10
        )
        if "OK" in res.stdout:
            return {
                "status": "ok",
                "taste": "PrintScreen",
                "hinweis": (
                    "Screenshot-Taste gesendet. Das Bild wird in MSFS unter "
                    "Dokumente\\Microsoft Flight Simulator 2024\\Screenshots gespeichert."
                ),
            }
        return {"status": "error", "error": res.stderr.strip()}
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def open_msfs_devmode(enable: bool = True, usercfg_path: str = "") -> dict:
    """Aktiviert oder deaktiviert den MSFS Developer Mode in UserCfg.opt.

    Wann aufrufen: Wenn der User den Developer Mode ein- oder ausschalten will
    (z.B. für FPS-Anzeige in MSFS oder Add-on-Entwicklung).
    Trigger: 'DevMode', 'Developer Mode', 'FPS Zähler MSFS', 'Entwicklermodus'.

    Args:
        enable: True = DevMode aktivieren, False = deaktivieren.
        usercfg_path: Optionaler Pfad zu UserCfg.opt.

    Returns:
        status: "ok" | "error"
        dev_mode: Neuer Wert ("1" oder "0").
    """
    cfg_path = _find_usercfg(usercfg_path)
    if not cfg_path or not cfg_path.exists():
        return {
            "status": "error",
            "error": "UserCfg.opt nicht gefunden. MSFS muss mindestens einmal gestartet worden sein.",
        }

    try:
        content = cfg_path.read_text(encoding="utf-8", errors="replace")
        new_value = "1" if enable else "0"

        import re as _re
        if _re.search(r'DevMode\s+\d', content):
            new_content = _re.sub(r'(DevMode\s+)\d', rf'\g<1>{new_value}', content)
        else:
            # Append DevMode setting
            new_content = content.rstrip() + f"\nDevMode {new_value}\n"

        cfg_path.write_text(new_content, encoding="utf-8")
        _invalidate_usercfg_cache()

        return {
            "status": "ok",
            "dev_mode": new_value,
            "pfad": str(cfg_path),
            "hinweis": (
                "Developer Mode aktiviert. In MSFS: Optionen → Allgemein → Entwicklermodus. "
                "Ermöglicht FPS-Overlay (Strg+F in MSFS)."
                if enable else
                "Developer Mode deaktiviert."
            ),
        }
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def get_frame_generation_status(usercfg_path: str = "") -> dict:
    """Prüft ob DLSS Frame Generation oder FSR Frame Generation in MSFS aktiviert ist.

    Wann aufrufen: Wenn der User nach Frame Generation fragt oder FG-Probleme
    diagnostiziert werden sollen.
    Trigger: 'Frame Generation', 'DLSS FG', 'FSR FG', 'Framegen Status'.

    Returns:
        frame_generation_aktiv: bool.
        dlss_modus: Aktueller DLSS-Modus.
        upscaling_typ: "DLSS" | "FSR" | "XeSS" | "deaktiviert".
    """
    cfg_path = _find_usercfg(usercfg_path)
    if not cfg_path or not cfg_path.exists():
        return {
            "status": "error",
            "error": "UserCfg.opt nicht gefunden.",
        }

    try:
        import re as _re
        content = cfg_path.read_text(encoding="utf-8", errors="replace")

        # Frame generation flag
        fg_match = _re.search(r'FrameGeneration\s+(\d+)', content)
        fg_aktiv = fg_match and fg_match.group(1) == "1"

        # Upscaling type (DLSS=1, FSR=2, XeSS=3, TAA=0)
        up_match = _re.search(r'UpscalingMode\s+(\d+)', content)
        up_code = int(up_match.group(1)) if up_match else -1
        upscaling_map = {0: "TAA (kein Upscaling)", 1: "DLSS", 2: "FSR", 3: "XeSS", -1: "unbekannt"}
        upscaling_typ = upscaling_map.get(up_code, "unbekannt")

        # DLSS mode (Quality=0, Balanced=1, Performance=2, Ultra Performance=3)
        dlss_match = _re.search(r'DlssQualityMode\s+(\d+)', content)
        dlss_code = int(dlss_match.group(1)) if dlss_match else -1
        dlss_map = {0: "Quality", 1: "Balanced", 2: "Performance", 3: "Ultra Performance", -1: "unbekannt"}
        dlss_modus = dlss_map.get(dlss_code, "unbekannt")

        return {
            "status": "ok",
            "frame_generation_aktiv": bool(fg_aktiv),
            "upscaling_typ": upscaling_typ,
            "dlss_modus": dlss_modus if upscaling_typ == "DLSS" else "—",
            "hinweis": (
                f"Frame Generation ist {'aktiviert' if fg_aktiv else 'deaktiviert'}. "
                f"Upscaling: {upscaling_typ}."
            ),
        }
    except Exception as e:
        return {"status": "error", "error": str(e)}


@mcp.tool()
def set_frame_generation(enable: bool = True, usercfg_path: str = "") -> dict:
    """Aktiviert oder deaktiviert MSFS Frame Generation (DLSS FG / FSR FG).

    Wann aufrufen: Wenn der User Frame Generation ein- oder ausschalten will.
    MSFS muss neu gestartet werden damit die Änderung wirkt.
    Trigger: 'Frame Generation aktivieren', 'FG einschalten', 'DLSS FG aus'.

    Args:
        enable: True = Frame Generation aktivieren, False = deaktivieren.
        usercfg_path: Optionaler Pfad zu UserCfg.opt.

    Returns:
        status: "ok" | "error"
        frame_generation: Neuer Wert ("1" oder "0").
    """
    cfg_path = _find_usercfg(usercfg_path)
    if not cfg_path or not cfg_path.exists():
        return {
            "status": "error",
            "error": "UserCfg.opt nicht gefunden.",
        }

    try:
        import re as _re
        content = cfg_path.read_text(encoding="utf-8", errors="replace")
        new_value = "1" if enable else "0"

        if _re.search(r'FrameGeneration\s+\d', content):
            new_content = _re.sub(
                r'(FrameGeneration\s+)\d',
                rf'\g<1>{new_value}',
                content
            )
        else:
            new_content = content.rstrip() + f"\nFrameGeneration {new_value}\n"

        cfg_path.write_text(new_content, encoding="utf-8")
        _invalidate_usercfg_cache()

        return {
            "status": "ok",
            "frame_generation": new_value,
            "pfad": str(cfg_path),
            "hinweis": (
                "Frame Generation aktiviert. MSFS neu starten damit die Änderung wirkt. "
                "Benötigt DLSS oder FSR als Upscaling-Modus."
                if enable else
                "Frame Generation deaktiviert. MSFS neu starten."
            ),
        }
    except Exception as e:
        return {"status": "error", "error": str(e)}


# ---------------------------------------------------------------------------
# ReShade — complete support optimised for MSFS VR
# ---------------------------------------------------------------------------

import configparser as _rs_configparser
import glob as _glob
import shutil as _sh
import struct as _struct

_RESHADE_MSFS_PATHS: list[Path] = [
    # MSFS 2024 MS Store
    Path(os.environ.get("LOCALAPPDATA", ""))
    / "Packages/Microsoft.Limitless_8wekyb3d8bbwe/LocalCache",
    # MSFS 2020 MS Store
    Path(os.environ.get("LOCALAPPDATA", ""))
    / "Packages/Microsoft.FlightSimulator_8wekyb3d8bbwe/LocalCache",
    # MSFS 2024 Steam
    Path("C:/Program Files (x86)/Steam/steamapps/common/MicrosoftFlightSimulator2024"),
    # MSFS 2020 Steam
    Path("C:/Program Files (x86)/Steam/steamapps/common/MicrosoftFlightSimulator"),
    # Boxed / retail
    Path("C:/Program Files/Microsoft Flight Simulator 2024"),
    Path("C:/Program Files/Microsoft Flight Simulator"),
]

# Wildcard paths resolved via glob
_RESHADE_GLOB_PATTERNS: list[str] = [
    "C:/Program Files/WindowsApps/Microsoft.Limitless_*",
    "C:/Program Files/WindowsApps/Microsoft.FlightSimulator_*",
]

# Effects whose names hint at high GPU cost
_RESHADE_HEAVY_EFFECTS = {"rtgi", "dof", "ssr", "mxao", "adof", "godrays", "bloom"}
_RESHADE_LIGHT_EFFECTS = {"cas", "lut", "smaa", "fxaa", "levels", "colorcorrection",
                          "clarity", "vibrance", "curves"}


def _rs_locate_ini() -> "Path | None":
    """Search all known MSFS install locations for ReShade.ini."""
    candidates: list[Path] = list(_RESHADE_MSFS_PATHS)
    for pattern in _RESHADE_GLOB_PATTERNS:
        try:
            candidates.extend(Path(p) for p in _glob.glob(pattern))
        except Exception:
            pass

    for base in candidates:
        ini = base / "ReShade.ini"
        if ini.exists():
            return ini

    # Fallback: scan APPDATA / LOCALAPPDATA trees one level deep
    for env_var in ("LOCALAPPDATA", "APPDATA"):
        root = Path(os.environ.get(env_var, ""))
        if not root.exists():
            continue
        for child in root.iterdir():
            ini = child / "ReShade.ini"
            if ini.exists():
                return ini
    return None


def _rs_load_ini() -> "tuple[_rs_configparser.ConfigParser, Path | None]":
    """Return (parsed ConfigParser, path).  Path is None if not found."""
    path = _rs_locate_ini()
    cp = _rs_configparser.ConfigParser(strict=False, allow_no_value=True)
    cp.optionxform = str  # preserve case
    if path and path.exists():
        try:
            cp.read(str(path), encoding="utf-8")
        except Exception:
            cp.read(str(path), encoding="latin-1")
    return cp, path


def _rs_save_ini(cp: "_rs_configparser.ConfigParser", path: Path) -> None:
    """Backup then overwrite ReShade.ini from configparser."""
    try:
        _sh.copy2(str(path), str(path.with_suffix(".bak")))
    except Exception:
        pass
    with open(str(path), "w", encoding="utf-8") as fh:
        cp.write(fh)


def _rs_list_presets(ini_path: "Path | None" = None) -> list[Path]:
    """Return list of .ini preset files near the ReShade install."""
    if ini_path is None:
        ini_path = _rs_locate_ini()
    if ini_path is None:
        return []
    base = ini_path.parent
    presets: list[Path] = []
    for candidate in base.iterdir():
        if candidate.suffix.lower() == ".ini" and candidate.name != "ReShade.ini":
            presets.append(candidate)
    # Also check a "reshade-presets" sub-folder
    preset_dir = base / "reshade-presets"
    if preset_dir.is_dir():
        presets.extend(preset_dir.glob("*.ini"))
    return presets


def _get_file_version(dll_path: Path) -> str:
    """Read PE file version string from a DLL (Windows only)."""
    try:
        import ctypes
        from ctypes import wintypes
        ver_size = ctypes.windll.version.GetFileVersionInfoSizeW(str(dll_path), None)
        if not ver_size:
            return "unknown"
        buf = ctypes.create_string_buffer(ver_size)
        ctypes.windll.version.GetFileVersionInfoW(str(dll_path), None, ver_size, buf)
        sub_block = "\\".encode("utf-16-le")
        lp_buffer = ctypes.c_void_p()
        pu_len = ctypes.c_uint()
        ctypes.windll.version.VerQueryValueW(
            buf, "\\", ctypes.byref(lp_buffer), ctypes.byref(pu_len)
        )
        if pu_len.value >= 20:
            ms = _struct.unpack_from("<I", (ctypes.c_char * 4).from_address(lp_buffer.value + 8))[0]
            ls = _struct.unpack_from("<I", (ctypes.c_char * 4).from_address(lp_buffer.value + 12))[0]
            return f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
    except Exception:
        pass
    return "unknown"


def _effect_cost(name: str) -> str:
    lo = name.lower()
    if any(h in lo for h in _RESHADE_HEAVY_EFFECTS):
        return "high"
    if any(l in lo for l in _RESHADE_LIGHT_EFFECTS):
        return "low"
    return "medium"


def _parse_techniques(techniques_str: str) -> list[str]:
    """Split a ReShade Techniques= value into individual effect names."""
    if not techniques_str:
        return []
    parts = [t.strip() for t in techniques_str.replace(",", "@").split("@") if t.strip()]
    # Each part can be "EffectName" or "EffectName@file.fx" — keep the name part
    names = []
    for p in parts:
        names.append(p.split("@")[0] if "@" in p else p)
    return names


@mcp.tool()
def get_reshade_status() -> dict:
    """Checks whether ReShade is installed for MSFS and returns a quick status overview.

    Wann aufrufen: Wenn der User fragt ob ReShade installiert ist, wie ReShade läuft,
    oder ob ReShade für VR konfiguriert ist.
    Trigger: 'ReShade Status', 'ist ReShade installiert', 'ReShade check'.

    Returns:
        installed: bool
        ini_path: str
        version: ReShade DLL version string
        current_preset: active preset file name
        performance_mode: bool
        depth_reversed_configured: bool (RESHADE_DEPTH_INPUT_IS_REVERSED=1)
        active_effects_count: int
        effects_folder_count: int — number of .fx files found
    """
    try:
        cp, ini_path = _rs_load_ini()

        if ini_path is None or not ini_path.exists():
            return {
                "installed": False,
                "ini_path": None,
                "message": "ReShade.ini not found in any known MSFS install path.",
            }

        base = ini_path.parent
        # Detect version from DLL
        version = "unknown"
        for dll_name in ("dxgi.dll", "d3d11.dll", "openxr_api_layer.dll"):
            dll = base / dll_name
            if dll.exists():
                version = _get_file_version(dll)
                break

        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        depth = cp["DEPTH"] if cp.has_section("DEPTH") else {}

        preset_path_raw = general.get("PresetPath", "")
        preset_name = Path(preset_path_raw).name if preset_path_raw else "none"

        perf_mode = general.get("PerformanceMode", "0") == "1"

        # Check preprocessor for depth reversed flag
        preprocessor = general.get("PreprocessorDefinitions", "")
        depth_reversed = "RESHADE_DEPTH_INPUT_IS_REVERSED=1" in preprocessor

        # Count .fx shaders
        fx_count = len(list(base.rglob("*.fx")))

        # Count active effects from current preset
        active_count = 0
        if preset_path_raw:
            preset_p = Path(preset_path_raw)
            if not preset_p.is_absolute():
                preset_p = base / preset_path_raw
            if preset_p.exists():
                pcp = _rs_configparser.ConfigParser(strict=False)
                pcp.optionxform = str
                try:
                    pcp.read(str(preset_p), encoding="utf-8")
                except Exception:
                    pass
                for sec in pcp.sections():
                    techs = pcp.get(sec, "Techniques", fallback="")
                    active_count += len(_parse_techniques(techs))

        return {
            "installed": True,
            "ini_path": str(ini_path),
            "version": version,
            "current_preset": preset_name,
            "performance_mode": perf_mode,
            "depth_reversed_configured": depth_reversed,
            "active_effects_count": active_count,
            "effects_folder_count": fx_count,
            "vr_ready": perf_mode and depth_reversed,
            "recommendation": (
                "ReShade is well configured for VR."
                if perf_mode and depth_reversed
                else "Run optimize_reshade_for_vr() for optimal MSFS VR performance."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_reshade_settings() -> dict:
    """Read all current ReShade.ini settings and the active preset's effect list.

    Wann aufrufen: Wenn der User die aktuellen ReShade-Einstellungen sehen will
    oder wissen will welche Effekte aktiv sind.
    Trigger: 'ReShade Einstellungen', 'ReShade settings', 'welche ReShade Effekte'.

    Returns full [GENERAL], [DEPTH], any [VR]/[OPENXR] section, plus active preset effects.
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found."}

        result: dict = {"ini_path": str(ini_path), "sections": {}}
        for section in cp.sections():
            result["sections"][section] = dict(cp[section])

        # Parse active preset
        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        preset_raw = general.get("PresetPath", "")
        preset_info: dict = {"path": preset_raw, "effects": []}
        if preset_raw:
            preset_p = Path(preset_raw)
            if not preset_p.is_absolute():
                preset_p = ini_path.parent / preset_raw
            if preset_p.exists():
                pcp = _rs_configparser.ConfigParser(strict=False)
                pcp.optionxform = str
                try:
                    pcp.read(str(preset_p), encoding="utf-8")
                except Exception:
                    pass
                for sec in pcp.sections():
                    techs_str = pcp.get(sec, "Techniques", fallback="")
                    for eff in _parse_techniques(techs_str):
                        preset_info["effects"].append({
                            "name": eff,
                            "cost": _effect_cost(eff),
                        })
                preset_info["effect_count"] = len(preset_info["effects"])

        result["active_preset"] = preset_info
        return result
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_reshade_performance_mode(enabled: bool) -> dict:
    """Enable or disable ReShade performance mode (strongly recommended ON for VR).

    Performance mode compiles shaders once and hides the overlay — critical for VR FPS.
    Disabling allows live shader editing but costs significant GPU time.

    Wann aufrufen: Wenn der User ReShade Performance-Modus ein- oder ausschalten will,
    oder wenn VR-Performance durch ReShade beeinträchtigt ist.
    Trigger: 'ReShade Performance Modus', 'ReShade schneller machen', 'ReShade VR Performance'.

    Args:
        enabled: True = performance mode ON (recommended for VR), False = OFF.

    Returns:
        status, previous value, new value, recommendation.
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found."}

        if not cp.has_section("GENERAL"):
            cp.add_section("GENERAL")

        prev = cp.get("GENERAL", "PerformanceMode", fallback="0")
        cp.set("GENERAL", "PerformanceMode", "1" if enabled else "0")
        _rs_save_ini(cp, ini_path)

        return {
            "status": "ok",
            "previous_value": prev,
            "new_value": "1" if enabled else "0",
            "performance_mode_enabled": enabled,
            "recommendation": (
                "Performance mode ON — ReShade will not impact VR frame-time during flight."
                if enabled
                else "Performance mode OFF — shader editing enabled but VR FPS may drop."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_reshade_presets() -> dict:
    """List all available ReShade preset files and their enabled effects.

    Wann aufrufen: Wenn der User wissen will welche ReShade Presets vorhanden sind
    oder zwischen Presets wechseln will.
    Trigger: 'ReShade Presets', 'ReShade Profile', 'welche ReShade Presets gibt es'.

    Returns list of presets with: name, path, active, effect_count, file_size_kb, effects.
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found."}

        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        active_path = general.get("PresetPath", "")

        presets = _rs_list_presets(ini_path)
        result_list = []

        for p in presets:
            pcp = _rs_configparser.ConfigParser(strict=False)
            pcp.optionxform = str
            try:
                pcp.read(str(p), encoding="utf-8")
            except Exception:
                continue

            effects = []
            for sec in pcp.sections():
                techs_str = pcp.get(sec, "Techniques", fallback="")
                for eff in _parse_techniques(techs_str):
                    effects.append({"name": eff, "cost": _effect_cost(eff)})

            is_active = (
                str(p.resolve()) == str(Path(active_path).resolve())
                if active_path
                else False
            )

            result_list.append({
                "name": p.stem,
                "path": str(p),
                "active": is_active,
                "effect_count": len(effects),
                "file_size_kb": round(p.stat().st_size / 1024, 1),
                "effects": effects,
            })

        return {
            "preset_count": len(result_list),
            "active_preset": Path(active_path).name if active_path else "none",
            "presets": result_list,
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def set_reshade_preset(preset_name: str) -> dict:
    """Switch ReShade to a different preset by name (without .ini extension).

    Wann aufrufen: Wenn der User das ReShade Preset wechseln will.
    Trigger: 'ReShade Preset wechseln', 'ReShade Profil laden', 'ReShade auf X umschalten'.

    Args:
        preset_name: Name of the preset (stem, e.g. 'VR_Light' or 'GameCopilot_VR').

    Returns:
        status, previous preset, new preset path.
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found."}

        presets = _rs_list_presets(ini_path)
        match = None
        for p in presets:
            if p.stem.lower() == preset_name.lower() or p.name.lower() == preset_name.lower():
                match = p
                break

        if match is None:
            available = [p.stem for p in presets]
            return {
                "error": f"Preset '{preset_name}' not found.",
                "available_presets": available,
            }

        if not cp.has_section("GENERAL"):
            cp.add_section("GENERAL")

        prev = cp.get("GENERAL", "PresetPath", fallback="none")
        cp.set("GENERAL", "PresetPath", str(match))
        _rs_save_ini(cp, ini_path)

        return {
            "status": "ok",
            "previous_preset": Path(prev).name if prev != "none" else "none",
            "new_preset": match.name,
            "new_preset_path": str(match),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def optimize_reshade_for_vr() -> dict:
    """Optimize all ReShade settings for maximum VR performance in MSFS.

    Sets performance mode ON, fixes depth buffer for MSFS, disables the overlay,
    and provides GPU-tier-specific effect recommendations.

    Wann aufrufen: Wenn der User ReShade für VR optimieren will, VR-Performance
    durch ReShade leidet, oder ReShade in VR einstellen will.
    Trigger: 'ReShade für VR optimieren', 'ReShade VR einstellen',
             'ReShade performance VR', 'ReShade optimieren MSFS'.

    Returns:
        status, changes applied, GPU tier, recommended effects to disable.
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found. Is ReShade installed for MSFS?"}

        changes: list[str] = []

        # Ensure sections exist
        for sec in ("GENERAL", "DEPTH"):
            if not cp.has_section(sec):
                cp.add_section(sec)
                changes.append(f"Created [{sec}] section")

        # 1. Performance mode
        if cp.get("GENERAL", "PerformanceMode", fallback="0") != "1":
            cp.set("GENERAL", "PerformanceMode", "1")
            changes.append("PerformanceMode = 1 (was 0)")

        # 2. Depth buffer settings for MSFS
        # Both keys exist across different ReShade versions — set both for compatibility
        if cp.get("DEPTH", "DepthCopyBeforeClears", fallback="0") != "1":
            cp.set("DEPTH", "DepthCopyBeforeClears", "1")
            changes.append("DepthCopyBeforeClears = 1")

        if cp.get("DEPTH", "CopyDepthBufferBeforeClears", fallback="0") != "1":
            cp.set("DEPTH", "CopyDepthBufferBeforeClears", "1")
            changes.append("CopyDepthBufferBeforeClears = 1")

        if cp.get("DEPTH", "UseAspectRatioHeuristics", fallback="1") != "0":
            cp.set("DEPTH", "UseAspectRatioHeuristics", "0")
            changes.append("UseAspectRatioHeuristics = 0")

        # 3. Preprocessor: RESHADE_DEPTH_INPUT_IS_REVERSED=1
        existing_pre = cp.get("GENERAL", "PreprocessorDefinitions", fallback="")
        pre_parts = [p.strip() for p in existing_pre.split(",") if p.strip()] if existing_pre else []
        # Remove any existing RESHADE_DEPTH_INPUT_IS_REVERSED entry
        pre_parts = [p for p in pre_parts if not p.startswith("RESHADE_DEPTH_INPUT_IS_REVERSED")]
        pre_parts.append("RESHADE_DEPTH_INPUT_IS_REVERSED=1")
        new_pre = ",".join(pre_parts)
        if new_pre != existing_pre:
            cp.set("GENERAL", "PreprocessorDefinitions", new_pre)
            changes.append("PreprocessorDefinitions: RESHADE_DEPTH_INPUT_IS_REVERSED=1 added")

        # 4. Suppress tutorial overlay
        if cp.get("GENERAL", "TutorialProgress", fallback="0") != "4":
            cp.set("GENERAL", "TutorialProgress", "4")
            changes.append("TutorialProgress = 4 (overlay suppressed)")

        # 5. Write
        _rs_save_ini(cp, ini_path)

        # 6. GPU-tier recommendations
        tier = "unknown"
        effect_advice: list[str] = []
        try:
            pynvml.nvmlInit()
            try:
                h = pynvml.nvmlDeviceGetHandleByIndex(0)
                name = pynvml.nvmlDeviceGetName(h)
                if isinstance(name, bytes):
                    name = name.decode()
                mem = pynvml.nvmlDeviceGetMemoryInfo(h)
                vram_gb = mem.total / 1024 ** 3
                tier = _classify_vr_tier(vram_gb, name, 8, 16)
            finally:
                pynvml.nvmlShutdown()
        except Exception:
            tier = "unknown"

        if tier in ("low", "mid"):
            effect_advice = [
                "Disable RTGI, MXAO, DOF, SSR — too expensive for VR on your GPU tier.",
                "Keep only CAS (sharpening) and LUT (colour grading) for best results.",
                "Consider using the GameCopilot_VR preset (create_reshade_vr_preset).",
            ]
        elif tier == "mid_high":
            effect_advice = [
                "Disable RTGI and MXAO — too expensive.",
                "SMAA, CAS, LUT are safe to use.",
            ]
        else:  # high / ultra / unknown
            effect_advice = [
                "Most effects are fine; avoid RTGI/MXAO in heavy VR scenarios.",
                "Monitor GPU frame time — keep it below 11 ms for 90 Hz VR.",
            ]

        # 7. Create / update visual quality preset for beautiful VR image
        visual_result: dict = {}
        try:
            _vr_tier = tier if tier in _RS_VISUAL_PROFILES else None
            visual_result = create_reshade_vr_visual_preset(
                style="auto", gpu_override=_vr_tier
            )
        except Exception as _ve:
            visual_result = {"note": f"Visual preset creation skipped: {_ve}"}

        return {
            "status": "ok",
            "ini_path": str(ini_path),
            "changes_applied": changes,
            "gpu_tier": tier,
            "effect_recommendations": effect_advice,
            "visual_preset": visual_result,
            "summary": (
                f"{len(changes)} change(s) applied. "
                "ReShade is now optimised for MSFS VR with visual quality preset. "
                "Restart MSFS for changes to take effect."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_reshade_effects() -> dict:
    """List all installed ReShade shader effects with enabled status and estimated GPU cost.

    Wann aufrufen: Wenn der User wissen will welche ReShade Effekte installiert sind,
    welche Effekte aktiv sind, oder welche Effekte Performance kosten.
    Trigger: 'ReShade Effekte', 'ReShade Shader Liste', 'welche ReShade Effekte aktiv'.

    Returns list of effects: name, file, enabled_in_current_preset, cost (low/medium/high).
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found."}

        base = ini_path.parent
        # Collect .fx files
        fx_files: list[Path] = list(base.rglob("*.fx"))

        # Get enabled effects from active preset
        enabled_names: set[str] = set()
        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        preset_raw = general.get("PresetPath", "")
        if preset_raw:
            preset_p = Path(preset_raw)
            if not preset_p.is_absolute():
                preset_p = base / preset_raw
            if preset_p.exists():
                pcp = _rs_configparser.ConfigParser(strict=False)
                pcp.optionxform = str
                try:
                    pcp.read(str(preset_p), encoding="utf-8")
                except Exception:
                    pass
                for sec in pcp.sections():
                    for eff in _parse_techniques(pcp.get(sec, "Techniques", fallback="")):
                        enabled_names.add(eff.lower())

        effects = []
        for fx in fx_files:
            name = fx.stem
            effects.append({
                "name": name,
                "file": fx.name,
                "relative_path": str(fx.relative_to(base)),
                "enabled_in_preset": name.lower() in enabled_names,
                "cost": _effect_cost(name),
            })

        effects.sort(key=lambda e: (0 if e["enabled_in_preset"] else 1, e["name"]))

        return {
            "total_effects": len(effects),
            "enabled_count": sum(1 for e in effects if e["enabled_in_preset"]),
            "effects": effects,
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def enable_reshade_technique(technique_name: str, enabled: bool) -> dict:
    """Enable or disable a specific ReShade effect in the currently active preset.

    Wann aufrufen: Wenn der User einen bestimmten ReShade-Effekt aktivieren oder
    deaktivieren will.
    Trigger: 'ReShade Effekt aktivieren', 'ReShade RTGI deaktivieren',
             'CAS einschalten ReShade', 'Effekt ausschalten ReShade'.

    Args:
        effect_name: Effect name, e.g. 'CAS', 'RTGI', 'DOF', 'SMAA'.
        enabled: True = enable, False = disable.

    Returns:
        status, effect_name, new state, preset updated.
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found."}

        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        preset_raw = general.get("PresetPath", "")
        if not preset_raw:
            return {"error": "No active preset configured in ReShade.ini."}

        preset_p = Path(preset_raw)
        if not preset_p.is_absolute():
            preset_p = ini_path.parent / preset_raw
        if not preset_p.exists():
            return {"error": f"Preset file not found: {preset_p}"}

        pcp = _rs_configparser.ConfigParser(strict=False)
        pcp.optionxform = str
        try:
            pcp.read(str(preset_p), encoding="utf-8")
        except Exception:
            pcp.read(str(preset_p), encoding="latin-1")

        # Find the section that holds Techniques (usually first section)
        target_section = pcp.sections()[0] if pcp.sections() else None
        if target_section is None:
            pcp.add_section("default")
            target_section = "default"

        raw_techs = pcp.get(target_section, "Techniques", fallback="")
        techs = [t.strip() for t in raw_techs.split(",") if t.strip()] if raw_techs else []

        # Remove existing entry for this effect (case-insensitive)
        remaining = [t for t in techs if not t.lower().startswith(technique_name.lower())]
        was_enabled = len(remaining) < len(techs)

        if enabled:
            # Add the effect (look for its .fx file for the full token)
            base = ini_path.parent
            fx_matches = list(base.rglob(f"{technique_name}.fx"))
            if fx_matches:
                token = f"{technique_name}@{fx_matches[0].name}"
            else:
                token = technique_name
            remaining.append(token)

        pcp.set(target_section, "Techniques", ",".join(remaining))

        try:
            _sh.copy2(str(preset_p), str(preset_p.with_suffix(".bak")))
        except Exception:
            pass
        with open(str(preset_p), "w", encoding="utf-8") as fh:
            pcp.write(fh)

        return {
            "status": "ok",
            "effect": technique_name,
            "was_enabled": was_enabled,
            "now_enabled": enabled,
            "preset": preset_p.name,
            "techniques_count": len(remaining),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def create_reshade_vr_preset() -> dict:
    """Create an optimised VR preset 'GameCopilot_VR' and activate it in ReShade.

    Enables only VR-friendly effects (CAS sharpening, LUT colour grading, SMAA AA).
    Disables all performance-heavy effects (RTGI, DOF, SSR, MXAO, heavy Bloom).

    Wann aufrufen: Wenn der User ein optimiertes VR-Preset erstellen will oder
    fragt wie man ReShade für VR am besten einstellt.
    Trigger: 'ReShade VR Preset erstellen', 'optimiertes ReShade Preset',
             'GameCopilot VR Preset', 'ReShade VR Profil'.

    Returns:
        status, preset_path, effects_enabled, previous_preset.
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini not found."}

        base = ini_path.parent
        preset_path = base / "GameCopilot_VR.ini"

        # Find available VR-friendly fx files
        vr_effects: list[str] = []
        for fx_stem in ("CAS", "LUT", "SMAA"):
            matches = list(base.rglob(f"{fx_stem}.fx"))
            if matches:
                vr_effects.append(f"{fx_stem}@{matches[0].name}")
            else:
                vr_effects.append(fx_stem)

        preset_content = (
            "# GameCopilot VR Preset — optimised for MSFS VR\n"
            "# Created by Game Copilot\n\n"
            "[TECHNIQUE_SORT]\n"
            f"Techniques={','.join(vr_effects)}\n\n"
        )

        try:
            if preset_path.exists():
                _sh.copy2(str(preset_path), str(preset_path.with_suffix(".bak")))
        except Exception:
            pass

        preset_path.write_text(preset_content, encoding="utf-8")

        # Activate it
        if not cp.has_section("GENERAL"):
            cp.add_section("GENERAL")
        prev_preset = cp.get("GENERAL", "PresetPath", fallback="none")
        cp.set("GENERAL", "PresetPath", str(preset_path))
        _rs_save_ini(cp, ini_path)

        return {
            "status": "ok",
            "preset_path": str(preset_path),
            "effects_enabled": vr_effects,
            "previous_preset": Path(prev_preset).name if prev_preset != "none" else "none",
            "note": (
                "GameCopilot_VR preset created and activated. "
                "Restart MSFS to apply. Only lightweight VR-safe effects are enabled."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def diagnose_reshade_vr() -> dict:
    """Full ReShade VR diagnostic for MSFS — checks installation, config, and active effects.

    Returns a prioritised list of warnings and actionable recommendations.

    Wann aufrufen: Wenn der User VR-Probleme mit ReShade hat, ReShade diagnostizieren will,
    oder fragt ob ReShade korrekt für VR konfiguriert ist.
    Trigger: 'ReShade VR Diagnose', 'ReShade überprüfen', 'ReShade Probleme VR',
             'warum ist ReShade langsam VR', 'ReShade debug'.

    Returns:
        overall_health: 'good' | 'warning' | 'critical'
        checks: list of individual check results
        recommendations: prioritised list of actions
    """
    try:
        cp, ini_path = _rs_load_ini()

        checks: list[dict] = []
        recommendations: list[str] = []

        # 1. Installation check
        if ini_path is None or not ini_path.exists():
            return {
                "overall_health": "critical",
                "checks": [{"name": "ReShade installed", "passed": False,
                             "detail": "ReShade.ini not found in any MSFS path."}],
                "recommendations": [
                    "Install ReShade from https://reshade.me into your MSFS exe folder.",
                    "Run the ReShade installer and point it to FlightSimulator.exe.",
                ],
            }
        checks.append({"name": "ReShade installed", "passed": True,
                        "detail": str(ini_path)})

        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        depth = cp["DEPTH"] if cp.has_section("DEPTH") else {}

        # 2. DLL injection method check — MSFS 2024 must use dxgi.dll (DX12)
        base = ini_path.parent
        _dll_found = None
        for _dll_candidate in ("dxgi.dll", "d3d11.dll", "d3d12.dll", "opengl32.dll"):
            if (base / _dll_candidate).exists():
                _dll_found = _dll_candidate
                break
        _openxr_layer = (base / "openxr_api_layer.json").exists()
        _dxgi_ok = _dll_found == "dxgi.dll"
        checks.append({
            "name": "Correct DLL injection (dxgi.dll)",
            "passed": _dxgi_ok,
            "detail": (
                f"dxgi.dll present — correct for MSFS 2024 (DX12)." if _dxgi_ok
                else f"Found: {_dll_found or 'none'} — MSFS 2024 requires dxgi.dll. "
                     "Re-install ReShade and select DirectX 10/11/12."
            ),
        })
        if not _dxgi_ok:
            recommendations.append(
                "[HIGH] Wrong injection DLL — reinstall ReShade, choose DirectX 10/11/12 "
                "(creates dxgi.dll). Current: " + (_dll_found or "none found") + "."
            )
        if _openxr_layer:
            checks.append({
                "name": "OpenXR layer present",
                "passed": True,
                "detail": "openxr_api_layer.json found — OpenXR add-on layer active.",
            })

        # 4. Performance mode
        perf = general.get("PerformanceMode", "0") == "1"
        checks.append({"name": "Performance mode ON", "passed": perf,
                        "detail": "PerformanceMode=" + ("1" if perf else "0 (OFF — hurts VR FPS!)")})
        if not perf:
            recommendations.append(
                "[HIGH] Enable performance mode: set_reshade_performance_mode(True) "
                "— saves 3-8 ms per frame in VR."
            )

        # 5. Depth buffer reversed for MSFS
        preprocessor = general.get("PreprocessorDefinitions", "")
        depth_ok = "RESHADE_DEPTH_INPUT_IS_REVERSED=1" in preprocessor
        checks.append({"name": "Depth buffer reversed (MSFS)", "passed": depth_ok,
                        "detail": "RESHADE_DEPTH_INPUT_IS_REVERSED=1" + (" set" if depth_ok else " MISSING")})
        if not depth_ok:
            recommendations.append(
                "[HIGH] Depth buffer not reversed — depth-based effects (DOF, MXAO) "
                "will look wrong. Run optimize_reshade_for_vr()."
            )

        # 4. DepthCopyBeforeClears
        dcbc = depth.get("DepthCopyBeforeClears", "0") == "1"
        checks.append({"name": "DepthCopyBeforeClears", "passed": dcbc,
                        "detail": "Needed for correct depth in MSFS"})
        if not dcbc:
            recommendations.append(
                "[MEDIUM] Set DepthCopyBeforeClears=1 in [DEPTH] for correct depth buffer access."
            )

        # 5. Overlay / tutorial suppressed
        tutorial_done = general.get("TutorialProgress", "0") == "4"
        checks.append({"name": "Overlay suppressed", "passed": tutorial_done,
                        "detail": "TutorialProgress=" + general.get("TutorialProgress", "0")})
        if not tutorial_done:
            recommendations.append(
                "[LOW] ReShade tutorial overlay may appear in VR — run optimize_reshade_for_vr()."
            )

        # 6. Heavy effects in active preset
        preset_raw = general.get("PresetPath", "")
        heavy_active: list[str] = []
        if preset_raw:
            preset_p = Path(preset_raw)
            if not preset_p.is_absolute():
                preset_p = ini_path.parent / preset_raw
            if preset_p.exists():
                pcp = _rs_configparser.ConfigParser(strict=False)
                pcp.optionxform = str
                try:
                    pcp.read(str(preset_p), encoding="utf-8")
                except Exception:
                    pass
                for sec in pcp.sections():
                    for eff in _parse_techniques(pcp.get(sec, "Techniques", fallback="")):
                        if _effect_cost(eff) == "high":
                            heavy_active.append(eff)

        heavy_ok = len(heavy_active) == 0
        checks.append({
            "name": "No heavy effects active",
            "passed": heavy_ok,
            "detail": (
                "No high-cost effects in preset." if heavy_ok
                else f"Heavy effects active: {', '.join(heavy_active)}"
            ),
        })
        if not heavy_ok:
            recommendations.append(
                f"[HIGH] Disable expensive effects: {', '.join(heavy_active)}. "
                "These can add 5-20 ms per eye in VR. Use set_reshade_effect() or create_reshade_vr_preset()."
            )

        # 7. UseAspectRatioHeuristics
        arh = depth.get("UseAspectRatioHeuristics", "1")
        arh_ok = arh == "0"
        checks.append({"name": "AspectRatioHeuristics disabled", "passed": arh_ok,
                        "detail": "Should be 0 for VR; currently " + arh})
        if not arh_ok:
            recommendations.append(
                "[MEDIUM] Set UseAspectRatioHeuristics=0 — heuristics interfere with VR depth detection."
            )

        # 8. Visual quality analysis: sharpness + color depth + Bildqualitäts-Score
        sharpness_val: "float | None" = None
        active_color_efx: list[str] = []
        if preset_raw:
            _vq_path = Path(preset_raw)
            if not _vq_path.is_absolute():
                _vq_path = ini_path.parent / preset_raw
            if _vq_path.exists():
                try:
                    _vq_cp = _rs_configparser.ConfigParser(strict=False)
                    _vq_cp.optionxform = str
                    _vq_cp.read(str(_vq_path), encoding="utf-8")
                    # CAS sharpness value
                    try:
                        sharpness_val = float(
                            _vq_cp.get("CAS.fx", "CAS_SHARPENING_AMOUNT", fallback="")
                        )
                    except (ValueError, TypeError):
                        pass
                    # All active techniques → filter for color effects
                    _all_techs: set[str] = set()
                    for _vsec in _vq_cp.sections():
                        for _vt in _parse_techniques(
                            _vq_cp.get(_vsec, "Techniques", fallback="")
                        ):
                            _all_techs.add(_vt)
                    _color_names = {"LiftGammaGain", "ColorMatrix", "Curves", "LUT",
                                    "Vibrance", "Tonemap"}
                    active_color_efx = sorted(_all_techs & _color_names)
                except Exception:
                    pass

        if sharpness_val is None:
            sharpness_rating = "unbekannt (CAS nicht aktiv)"
        elif sharpness_val < 0.30:
            sharpness_rating = "zu niedrig — Bild wirkt unscharf (<0.30)"
        elif sharpness_val > 0.75:
            sharpness_rating = "zu hoch — Halos/Ringing möglich (>0.75)"
        else:
            sharpness_rating = "gut"

        if not active_color_efx:
            color_rating = "keine Farbkorrektur aktiv"
        elif len(active_color_efx) >= 3:
            color_rating = "vollständige Farbkorrektur"
        elif len(active_color_efx) >= 2:
            color_rating = "gute Farbkorrektur"
        else:
            color_rating = "minimale Farbkorrektur"

        # Sharpness/color recommendations
        if sharpness_val is not None and sharpness_val < 0.30:
            recommendations.append(
                "[LOW] CAS-Schärfe zu niedrig (<0.30) — Bild wirkt unscharf. "
                "Verwende tune_reshade_sharpness('medium') oder 'high'."
            )
        if sharpness_val is not None and sharpness_val > 0.75:
            recommendations.append(
                "[LOW] CAS-Schärfe sehr hoch (>0.75) — Halos/Ringing möglich. "
                "Verwende tune_reshade_sharpness('high')."
            )
        if not active_color_efx:
            recommendations.append(
                "[LOW] Keine Farbkorrektur aktiv — bessere Bildqualität mit "
                "create_reshade_vr_visual_preset() oder tune_reshade_colors()."
            )

        # Bildqualitäts-Score 0–100
        _qs = 0
        if perf:          _qs += 15
        if depth_ok:      _qs += 10
        if heavy_ok:      _qs += 15
        if dcbc:          _qs += 5
        if arh_ok:        _qs += 5
        if tutorial_done: _qs += 3
        if preset_raw:    _qs += 2
        if sharpness_val is not None and 0.30 <= sharpness_val <= 0.75:
            _qs += 20
        elif sharpness_val is not None:
            _qs += 5
        _qs += min(20, len(active_color_efx) * 5)
        image_quality_score = min(100, _qs)

        passed = sum(1 for c in checks if c["passed"])
        total = len(checks)
        if passed == total:
            health = "good"
        elif passed >= total - 2:
            health = "warning"
        else:
            health = "critical"

        return {
            "overall_health": health,
            "passed_checks": passed,
            "total_checks": total,
            "checks": checks,
            "recommendations": recommendations,
            "sharpness_value": sharpness_val,
            "sharpness_rating": sharpness_rating,
            "color_rating": color_rating,
            "active_color_effects": active_color_efx,
            "image_quality_score": image_quality_score,
            "quick_fix": (
                "Run optimize_reshade_for_vr() to fix all critical issues automatically."
                if health != "good"
                else "ReShade is well configured for MSFS VR."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# ReShade Visual Quality System — GPU-tier profiles + visual tuning tools
# ---------------------------------------------------------------------------

# ── Per-tier visual profiles ──────────────────────────────────────────────

_RS_VISUAL_PROFILES: dict = {
    "ultra": {  # RTX 4090, 7900 XTX — can afford everything beautiful
        "description": "Ultra-Qualität: Filmische Farbgebung, starkes Sharpening, volle Farbkorrektur + Vibrance/Clarity/Deband",
        "techniques": ["CAS", "LUT", "SMAA", "Levels", "Curves", "ColorMatrix", "LiftGammaGain",
                       "Vibrance", "Clarity", "FilmGrain", "Deband"],
        "disabled_techniques": ["RTGI", "DOF", "MotionBlur"],
        "parameters": {
            "CAS": {"CAS_SHARPENING_AMOUNT": 0.65},
            "Levels": {"BlackPoint": 5, "WhitePoint": 245},
            "Curves": {"Mode": 0, "Formula": 4, "Contrast": 0.25},
            "LiftGammaGain": {
                # Cool/blue shadows (atmospheric scattering depth cue)
                # Warm R/G midtones, MSFS tends cool/desaturated
                # Near-white highlights — don't blow out HDR-like sky
                "RGB_Lift":  [0.997, 0.998, 1.018],
                "RGB_Gamma": [1.010, 1.006, 0.990],
                "RGB_Gain":  [1.006, 1.002, 0.997],
            },
            "ColorMatrix": {
                # Enhanced greens for terrain; slightly boosted blues for sky
                "ColorMatrix_Red":   [0.860, 0.140, 0.000],
                "ColorMatrix_Green": [0.040, 0.938, 0.022],
                "ColorMatrix_Blue":  [0.000, 0.042, 0.958],
                "Strength": 0.35,
            },
            # Vibrance: hue-safe saturation boost (won't shift sky/skin tones)
            "Vibrance": {"Vibrance_Intensity": 0.15},
            # Clarity: midtone microcontrast — scenery "pops" without oversharpening
            "Clarity": {"Clarity_BlendMode": 2, "Clarity_Strength": 0.20},
            # FilmGrain: tiny amount reduces VR screen-door effect perception
            "FilmGrain": {"FilmGrain_Intensity": 0.020},
            # Deband: removes colour banding in VR sky gradients
            "Deband": {"Deband_Threshold": 64, "Deband_Range": 16},
        },
    },
    "high": {  # RTX 4080, 4070 Ti — good quality, some restraint
        "description": "High-Qualität: Klare Farben, gutes Sharpening, leichte Atmosphäre + Vibrance/Deband",
        "techniques": ["CAS", "LUT", "SMAA", "Levels", "LiftGammaGain", "Vibrance", "Deband"],
        "disabled_techniques": ["RTGI", "DOF", "ColorMatrix", "Curves", "Clarity", "FilmGrain"],
        "parameters": {
            "CAS": {"CAS_SHARPENING_AMOUNT": 0.55},
            "Levels": {"BlackPoint": 8, "WhitePoint": 248},
            "LiftGammaGain": {
                # Cool shadows, warm midtones, clean highlights
                "RGB_Lift":  [0.998, 0.999, 1.012],
                "RGB_Gamma": [1.007, 1.004, 0.992],
                "RGB_Gain":  [1.004, 1.001, 0.998],
            },
            "Vibrance": {"Vibrance_Intensity": 0.15},
            "Deband": {"Deband_Threshold": 64, "Deband_Range": 16},
        },
    },
    "mid_high": {  # RTX 3080, 4070 — balanced, still beautiful
        "description": "Ausgewogen: Sharpening + Tonemap + Deband, VR-optimiert",
        "techniques": ["CAS", "LUT", "SMAA", "Levels", "Tonemap", "Deband"],
        "disabled_techniques": ["RTGI", "DOF", "ColorMatrix", "LiftGammaGain", "Curves",
                                 "Vibrance", "Clarity", "FilmGrain"],
        "parameters": {
            "CAS": {"CAS_SHARPENING_AMOUNT": 0.50},
            "Levels": {"BlackPoint": 10, "WhitePoint": 245},
            # Tonemap: gentle exposure/saturation lift — better than Levels alone
            "Tonemap": {"Gamma": 1.0, "Exposure": 0.0, "Saturation": 0.10,
                        "Bleach": 0.0, "Defog": 0.0},
            "Deband": {"Deband_Threshold": 64, "Deband_Range": 16},
        },
    },
    "mid": {  # RTX 3070, RX 6700 XT — minimal but impactful
        "description": "Effizient: CAS Sharpening + Tonemap + Deband, kaum FPS-Kosten",
        "techniques": ["CAS", "Tonemap", "Deband"],
        "disabled_techniques": ["RTGI", "DOF", "ColorMatrix", "LiftGammaGain", "SMAA", "Curves",
                                 "LUT", "Levels", "Vibrance", "Clarity", "FilmGrain"],
        "parameters": {
            "CAS": {"CAS_SHARPENING_AMOUNT": 0.40},
            "Tonemap": {"Gamma": 1.0, "Exposure": 0.0, "Saturation": 0.10,
                        "Bleach": 0.0, "Defog": 0.0},
            "Deband": {"Deband_Threshold": 64, "Deband_Range": 16},
        },
    },
    "low": {  # RTX 3060 and below
        "description": "Minimal: CAS Sharpening + Deband, maximale Performance",
        "techniques": ["CAS", "Deband"],
        "disabled_techniques": ["RTGI", "DOF", "ColorMatrix", "LiftGammaGain", "SMAA", "Curves",
                                 "LUT", "Levels", "Vibrance", "Clarity", "FilmGrain", "Tonemap"],
        "parameters": {
            "CAS": {"CAS_SHARPENING_AMOUNT": 0.35},
            "Deband": {"Deband_Threshold": 64, "Deband_Range": 16},
        },
    },
}


# ── Visual helpers ─────────────────────────────────────────────────────────

def _rs_detect_hdr() -> bool:
    """Return True if MSFS / Windows HDR is likely active.

    Checks (in order):
    1. Windows HDR registry flag (AdvancedColorEnabled) for the primary display.
    2. MSFS UserCfg.opt GraphicsVR/HDR key.
    Falls back to False on any error.
    """
    try:
        import winreg
        # Windows 11/10 HDR registry path
        _hdr_paths = [
            (winreg.HKEY_LOCAL_MACHINE,
             r"SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration"),
        ]
        # A lighter check: DisplayAdvancedColorInfo via PowerShell would be ideal,
        # but we can query the simpler EnableHDR DWM key on Win11
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\VideoSettings",
            ) as k:
                val, _ = winreg.QueryValueEx(k, "EnableHDROutput")
                if val == 1:
                    return True
        except OSError:
            pass

        # Also check MSFS UserCfg.opt for HDR flag
        _cfg_paths = [
            Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator 2024" / "UserCfg.opt",
            Path(os.environ.get("LOCALAPPDATA", ""))
            / "Packages/Microsoft.Limitless_8wekyb3d8bbwe/LocalCache/UserCfg.opt",
            Path(os.environ.get("APPDATA", "")) / "Microsoft Flight Simulator" / "UserCfg.opt",
        ]
        for _p in _cfg_paths:
            if _p.exists():
                try:
                    txt = _p.read_text(encoding="utf-8", errors="ignore")
                    # MSFS uses "Hdr 1" in UserCfg.opt
                    if "\nHdr 1" in txt or "\r\nHdr 1" in txt:
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    return False


def _rs_detect_gpu_tier() -> str:
    """Detect GPU tier using pynvml, falling back to 'mid'."""
    try:
        pynvml.nvmlInit()
        try:
            h = pynvml.nvmlDeviceGetHandleByIndex(0)
            name = pynvml.nvmlDeviceGetName(h)
            if isinstance(name, bytes):
                name = name.decode()
            mem = pynvml.nvmlDeviceGetMemoryInfo(h)
            vram_gb = mem.total / 1024 ** 3
        finally:
            pynvml.nvmlShutdown()
        return _classify_vr_tier(vram_gb, name, 8, 16)
    except Exception:
        return "mid"


def _rs_write_visual_preset(profile: dict, preset_path: "Path") -> None:
    """Create/overwrite a ReShade preset .ini with technique list and exact parameter values.

    ReShade preset format:
        [TECHNIQUE_SORT]
        Techniques=CAS,LUT,...

        [CAS.fx]
        CAS_SHARPENING_AMOUNT=0.650000

        [LiftGammaGain.fx]
        RGB_Lift=1.000000 1.000000 1.010000
    """
    import configparser as _rcp

    cp = _rcp.RawConfigParser()
    cp.optionxform = str  # preserve case

    # Top section: enabled techniques only
    cp.add_section("TECHNIQUE_SORT")
    cp.set("TECHNIQUE_SORT", "Techniques", ",".join(profile["techniques"]))

    # Per-technique parameter sections
    for tech, params in profile.get("parameters", {}).items():
        sec_name = f"{tech}.fx"
        cp.add_section(sec_name)
        for key, val in params.items():
            if isinstance(val, list):
                formatted = " ".join(f"{v:.6f}" for v in val)
            elif isinstance(val, float):
                formatted = f"{val:.6f}"
            else:
                formatted = str(val)
            cp.set(sec_name, key, formatted)

    # Backup existing file
    try:
        if preset_path.exists():
            _sh.copy2(str(preset_path), str(preset_path.with_suffix(".bak")))
    except Exception:
        pass

    preset_path.parent.mkdir(parents=True, exist_ok=True)
    with open(str(preset_path), "w", encoding="utf-8") as fh:
        fh.write("# GameCopilot Visual VR Preset\n")
        fh.write("# Auto-generated by Game Copilot — do not edit manually\n\n")
        cp.write(fh)


def _rs_apply_style_overrides(profile: dict, style: str) -> dict:
    """Return a deep copy of *profile* with style-specific parameter overrides applied."""
    import copy
    p = copy.deepcopy(profile)
    params = p.setdefault("parameters", {})

    if style == "cinematic":
        # Warm shadows, stronger contrast, cinematic S-curve
        lgg = params.setdefault("LiftGammaGain", {})
        lgg["RGB_Lift"]  = [1.000, 0.998, 1.015]
        lgg["RGB_Gamma"] = [0.975, 0.980, 1.000]
        lgg["RGB_Gain"]  = [1.015, 1.008, 1.000]
        curves = params.setdefault("Curves", {})
        curves["Mode"] = 0
        curves["Formula"] = 4
        curves["Contrast"] = 0.35
        for eff in ("LiftGammaGain", "Curves"):
            if eff not in p["techniques"]:
                p["techniques"].append(eff)

    elif style == "sharp":
        cas = params.setdefault("CAS", {})
        cas["CAS_SHARPENING_AMOUNT"] = min(
            0.80, float(cas.get("CAS_SHARPENING_AMOUNT", 0.50)) + 0.10
        )
        if "SMAA" not in p["techniques"]:
            p["techniques"].append("SMAA")

    elif style == "natural":
        # Gentle Levels, very subtle LiftGammaGain — almost invisible touch
        lv = params.setdefault("Levels", {})
        lv["BlackPoint"] = 3
        lv["WhitePoint"] = 252
        lgg = params.setdefault("LiftGammaGain", {})
        lgg["RGB_Lift"]  = [1.000, 1.000, 1.002]
        lgg["RGB_Gamma"] = [0.995, 0.997, 1.000]
        lgg["RGB_Gain"]  = [1.002, 1.001, 1.000]
        for eff in ("Levels", "LiftGammaGain"):
            if eff not in p["techniques"]:
                p["techniques"].append(eff)

    elif style == "vivid":
        # Stronger ColorMatrix saturation boost
        cm = params.setdefault("ColorMatrix", {})
        cm["ColorMatrix_Red"]   = [0.840, 0.160, 0.000]
        cm["ColorMatrix_Green"] = [0.040, 0.930, 0.030]
        cm["ColorMatrix_Blue"]  = [0.000, 0.040, 0.960]
        cm["Strength"] = 0.50
        if "ColorMatrix" not in p["techniques"]:
            p["techniques"].append("ColorMatrix")

    # Deduplicate technique list, preserving order
    seen: set = set()
    unique: list = []
    for t in p["techniques"]:
        if t not in seen:
            seen.add(t)
            unique.append(t)
    p["techniques"] = unique
    return p


def _rs_read_preset_params(preset_path: "Path") -> dict:
    """Read a ReShade preset .ini and return {section: {key: value}} dict."""
    import configparser as _rcp
    cp = _rcp.RawConfigParser()
    cp.optionxform = str
    try:
        cp.read(str(preset_path), encoding="utf-8")
    except Exception:
        try:
            cp.read(str(preset_path), encoding="latin-1")
        except Exception:
            return {}
    return {sec: dict(cp[sec]) for sec in cp.sections()}


def _rs_get_active_preset_path(ini_path: "Path") -> "Path | None":
    """Return the currently active ReShade preset Path, or None."""
    import configparser as _rcp
    cp = _rcp.RawConfigParser()
    cp.optionxform = str
    try:
        cp.read(str(ini_path), encoding="utf-8")
    except Exception:
        return None
    general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
    preset_raw = general.get("PresetPath", "")
    if not preset_raw:
        return None
    p = Path(preset_raw)
    if not p.is_absolute():
        p = ini_path.parent / preset_raw
    return p if p.exists() else None


def _rs_patch_preset_param(preset_path: "Path", section: str, key: str, value: str) -> bool:
    """Patch a single key=value in a preset .ini section. Returns True on success."""
    import configparser as _rcp
    cp = _rcp.RawConfigParser()
    cp.optionxform = str
    try:
        cp.read(str(preset_path), encoding="utf-8")
    except Exception:
        return False
    if not cp.has_section(section):
        cp.add_section(section)
    cp.set(section, key, value)
    try:
        _sh.copy2(str(preset_path), str(preset_path.with_suffix(".bak")))
    except Exception:
        pass
    with open(str(preset_path), "w", encoding="utf-8") as fh:
        cp.write(fh)
    return True


# ── New visual quality MCP tools ───────────────────────────────────────────

@mcp.tool()
def create_reshade_vr_visual_preset(
    style: str = "auto",
    gpu_override: "str | None" = None,
) -> dict:
    """Erstellt ein perfekt abgestimmtes ReShade VR-Preset für maximale Bildqualität.

    Analysiert die GPU automatisch und wählt das passende Qualitätsprofil.
    Schreibt alle Shader-Parameter (CAS, LUT, Levels, LiftGammaGain, ColorMatrix, Curves)
    präzise in die Preset-Datei — kein manuelles Einstellen nötig.

    Wann aufrufen: 'ReShade VR Preset erstellen', 'perfektes Bild einstellen',
    'Farben optimieren', 'schärfer machen', 'besseres Bild in VR',
    'GameCopilot VR Preset', 'ReShade VR Profil', 'optimiertes ReShade Preset'.

    Args:
        style: Visueller Stil — 'auto'|'cinematic'|'sharp'|'natural'|'vivid'
               auto      = GPU-Tier-Standardwerte
               cinematic = warme Schatten, starke Kontraste, filmischer Look
               sharp     = maximale Schärfe (+0.10 CAS), SMAA hinzu
               natural   = sanfte Farbkorrektur, minimaler Eingriff
               vivid     = verstärkte Sättigung via ColorMatrix
        gpu_override: GPU-Tier erzwingen — 'ultra'|'high'|'mid_high'|'mid'|'low'

    Returns:
        status, preset_path, gpu_tier, profile_description, techniques_enabled,
        techniques_disabled, parameters_written, previous_preset
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden. Ist ReShade installiert?"}

        # Detect or override tier
        tier = (
            gpu_override
            if gpu_override in _RS_VISUAL_PROFILES
            else _rs_detect_gpu_tier()
        )
        profile = _RS_VISUAL_PROFILES[tier]

        # Apply style overrides (deep copy so profile dict is unchanged)
        if style and style not in ("auto", ""):
            profile = _rs_apply_style_overrides(profile, style)

        # HDR awareness: if HDR is active, CAS adds visible ringing and Levels
        # clamps the extended range — dial both back
        import copy as _copy_mod
        hdr_active = _rs_detect_hdr()
        hdr_note = ""
        if hdr_active:
            profile = _copy_mod.deepcopy(profile)
            params = profile.setdefault("parameters", {})
            cas_params = params.setdefault("CAS", {})
            original_cas = float(cas_params.get("CAS_SHARPENING_AMOUNT", 0.50))
            cas_params["CAS_SHARPENING_AMOUNT"] = round(max(0.20, original_cas - 0.15), 2)
            # Remove Levels — HDR pipeline handles black/white point natively
            if "Levels" in profile.get("techniques", []):
                profile["techniques"] = [t for t in profile["techniques"] if t != "Levels"]
                profile.setdefault("disabled_techniques", [])
                if "Levels" not in profile["disabled_techniques"]:
                    profile["disabled_techniques"].append("Levels")
            hdr_note = (
                f" [HDR erkannt: CAS {original_cas:.2f}→{cas_params['CAS_SHARPENING_AMOUNT']:.2f},"
                " Levels deaktiviert (HDR übernimmt Tonwerte nativ)]"
            )

        # Build preset path inside reshade-presets sub-folder
        base = ini_path.parent
        preset_dir = base / "reshade-presets"
        preset_dir.mkdir(exist_ok=True)
        preset_path = preset_dir / "GameCopilot_VR_Visual.ini"

        _rs_write_visual_preset(profile, preset_path)

        # Point ReShade.ini to this preset
        if not cp.has_section("GENERAL"):
            cp.add_section("GENERAL")
        prev_preset = cp.get("GENERAL", "PresetPath", fallback="none")
        cp.set("GENERAL", "PresetPath", str(preset_path))
        _rs_save_ini(cp, ini_path)

        # Summarise parameters for the return value
        params_written: dict = {}
        for tech, vals in profile.get("parameters", {}).items():
            params_written[tech] = {
                k: (
                    " ".join(f"{v:.6f}" for v in val)
                    if isinstance(val, list)
                    else f"{val:.6f}"
                    if isinstance(val, float)
                    else str(val)
                )
                for k, val in vals.items()
            }

        return {
            "status": "ok",
            "preset_path": str(preset_path),
            "gpu_tier": tier,
            "style": style or "auto",
            "hdr_active": hdr_active,
            "profile_description": profile["description"],
            "techniques_enabled": profile["techniques"],
            "techniques_disabled": profile.get("disabled_techniques", []),
            "parameters_written": params_written,
            "previous_preset": Path(prev_preset).name if prev_preset != "none" else "none",
            "note": (
                f"GameCopilot_VR_Visual ({tier}, {style or 'auto'}) erstellt und aktiviert."
                + hdr_note + " "
                "MSFS neu starten um Änderungen zu sehen. "
                "Falls sofort sichtbar: Home-Taste → ReShade-Overlay → Preset neu laden."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def tune_reshade_sharpness(level: str) -> dict:
    """Passe die ReShade CAS-Schärfe im aktiven Preset an.

    Wann aufrufen: 'schärfer machen', 'weniger scharf', 'Schärfe anpassen',
    'CAS Wert ändern', 'Bild zu unscharf', 'Bild zu scharf'.

    Args:
        level: 'low'|'medium'|'high'|'ultra'
               low=0.25 (weich), medium=0.45 (ausgewogen), high=0.65 (scharf), ultra=0.80 (max)

    Returns:
        status, level, previous_value, new_value, preset_updated
    """
    try:
        level_map = {"low": 0.25, "medium": 0.45, "high": 0.65, "ultra": 0.80}
        if level not in level_map:
            return {
                "error": f"Ungültiger Level '{level}'. Gültig: {', '.join(level_map)}",
                "hint": "low=0.25 (weich), medium=0.45, high=0.65, ultra=0.80",
            }

        new_val = level_map[level]
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        preset_path = _rs_get_active_preset_path(ini_path)
        if preset_path is None:
            return {
                "error": "Kein aktives Preset gefunden.",
                "hint": "Erstelle zuerst ein Preset mit create_reshade_vr_visual_preset().",
            }

        # Read current value for reporting
        existing = _rs_read_preset_params(preset_path)
        prev_raw = existing.get("CAS.fx", {}).get("CAS_SHARPENING_AMOUNT", "unbekannt")

        success = _rs_patch_preset_param(
            preset_path, "CAS.fx", "CAS_SHARPENING_AMOUNT", f"{new_val:.6f}"
        )
        if not success:
            return {"error": "Konnte Preset-Datei nicht schreiben."}

        return {
            "status": "ok",
            "level": level,
            "previous_value": prev_raw,
            "new_value": f"{new_val:.6f}",
            "preset": preset_path.name,
            "note": (
                f"CAS Schärfe auf '{level}' ({new_val:.2f}) gesetzt. "
                "ReShade übernimmt Änderungen automatisch; "
                "falls nicht: Home-Taste → ReShade-Overlay → Preset neu laden."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def tune_reshade_colors(
    saturation: str = "normal",
    warmth: str = "neutral",
    contrast: str = "normal",
) -> dict:
    """Passe ReShade Farben individuell an — Sättigung, Farbtemperatur, Kontrast.

    Schreibt direkt die LiftGammaGain, ColorMatrix und Curves Parameter im aktiven Preset.

    Wann aufrufen: 'Farben wärmer machen', 'mehr Sättigung', 'mehr Kontrast',
    'Farben anpassen', 'kältere Farben', 'wärmere Farben', 'Bild zu blass',
    'Bild zu flach', 'Farben lebendiger'.

    Args:
        saturation: 'low'|'normal'|'high'|'vivid'   — Farbsättigung via ColorMatrix
        warmth:     'cool'|'neutral'|'warm'          — Farbtemperatur via LiftGammaGain
        contrast:   'flat'|'normal'|'punchy'         — Kontrast via Curves

    Returns:
        status, saturation, warmth, contrast, parameters_written, preset_updated
    """
    try:
        valid = {
            "saturation": ("low", "normal", "high", "vivid"),
            "warmth": ("cool", "neutral", "warm"),
            "contrast": ("flat", "normal", "punchy"),
        }
        errors = [
            f"{k} muss eines von {v} sein — got '{eval(k)}'"
            for k, v in valid.items()
            if eval(k) not in v
        ]
        if errors:
            return {"error": "; ".join(errors)}

        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        preset_path = _rs_get_active_preset_path(ini_path)
        if preset_path is None:
            return {
                "error": "Kein aktives Preset gefunden.",
                "hint": "Erstelle zuerst ein Preset mit create_reshade_vr_visual_preset().",
            }

        params_written: dict = {}

        # ── LiftGammaGain — warmth ─────────────────────────────────────────
        warmth_lgg: dict = {
            "cool": {
                "RGB_Lift":  [1.000, 1.000, 1.020],
                "RGB_Gamma": [0.990, 0.992, 1.010],
                "RGB_Gain":  [0.995, 0.998, 1.015],
            },
            "neutral": {
                "RGB_Lift":  [1.000, 1.000, 1.000],
                "RGB_Gamma": [1.000, 1.000, 1.000],
                "RGB_Gain":  [1.000, 1.000, 1.000],
            },
            "warm": {
                "RGB_Lift":  [1.010, 1.005, 1.000],
                "RGB_Gamma": [1.008, 1.002, 0.990],
                "RGB_Gain":  [1.015, 1.008, 0.995],
            },
        }
        for key, vals in warmth_lgg[warmth].items():
            fmt = " ".join(f"{v:.6f}" for v in vals)
            _rs_patch_preset_param(preset_path, "LiftGammaGain.fx", key, fmt)
            params_written[f"LiftGammaGain.fx/{key}"] = fmt

        # ── ColorMatrix — saturation ───────────────────────────────────────
        sat_cm: dict = {
            "low": {
                "ColorMatrix_Red":   [0.700, 0.200, 0.100],
                "ColorMatrix_Green": [0.100, 0.800, 0.100],
                "ColorMatrix_Blue":  [0.050, 0.150, 0.800],
                "Strength": 0.20,
            },
            "normal": {
                "ColorMatrix_Red":   [0.900, 0.100, 0.000],
                "ColorMatrix_Green": [0.050, 0.950, 0.000],
                "ColorMatrix_Blue":  [0.000, 0.050, 0.950],
                "Strength": 0.25,
            },
            "high": {
                "ColorMatrix_Red":   [0.860, 0.140, 0.000],
                "ColorMatrix_Green": [0.050, 0.930, 0.020],
                "ColorMatrix_Blue":  [0.000, 0.060, 0.940],
                "Strength": 0.40,
            },
            "vivid": {
                "ColorMatrix_Red":   [0.840, 0.160, 0.000],
                "ColorMatrix_Green": [0.040, 0.940, 0.020],
                "ColorMatrix_Blue":  [0.000, 0.040, 0.960],
                "Strength": 0.55,
            },
        }
        for key, val in sat_cm[saturation].items():
            fmt = " ".join(f"{v:.6f}" for v in val) if isinstance(val, list) else f"{val:.6f}"
            _rs_patch_preset_param(preset_path, "ColorMatrix.fx", key, fmt)
            params_written[f"ColorMatrix.fx/{key}"] = fmt

        # ── Curves — contrast ──────────────────────────────────────────────
        contrast_curves: dict = {
            "flat":   {"Mode": "0", "Formula": "4", "Contrast": "0.050000"},
            "normal": {"Mode": "0", "Formula": "4", "Contrast": "0.200000"},
            "punchy": {"Mode": "0", "Formula": "4", "Contrast": "0.400000"},
        }
        for key, val in contrast_curves[contrast].items():
            _rs_patch_preset_param(preset_path, "Curves.fx", key, val)
            params_written[f"Curves.fx/{key}"] = val

        return {
            "status": "ok",
            "saturation": saturation,
            "warmth": warmth,
            "contrast": contrast,
            "preset": preset_path.name,
            "parameters_written": params_written,
            "note": (
                f"Farben angepasst: Sättigung={saturation}, Wärme={warmth}, Kontrast={contrast}. "
                "ReShade übernimmt Änderungen automatisch; "
                "falls nicht: Home-Taste → ReShade-Overlay → Preset neu laden."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_reshade_visual_analysis() -> dict:
    """Analysiert das aktuelle ReShade-Setup und gibt konkrete Verbesserungsvorschläge.

    Berechnet einen Bildqualitäts-Score (0–100) basierend auf GPU-Tier, aktiven Effekten,
    Schärfewert und Farbkorrektur-Konfiguration.

    Wann aufrufen: 'ReShade Bildqualität analysieren', 'ReShade überprüfen',
    'wie sieht mein ReShade aus', 'ReShade Verbesserungen', 'Bildqualität prüfen'.

    Returns:
        current_tier, tier_description, current_preset, active_effects,
        expected_effects_for_tier, missing_effects, heavy_effects_active,
        sharpness_value, sharpness_rating, color_rating, active_color_effects,
        image_quality_score (0–100), performance_mode, recommendations
    """
    try:
        cp, ini_path = _rs_load_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        tier = _rs_detect_gpu_tier()
        profile = _RS_VISUAL_PROFILES[tier]
        expected_techniques = set(profile["techniques"])

        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        perf_mode = general.get("PerformanceMode", "0") == "1"
        preset_raw = general.get("PresetPath", "")
        preset_name = Path(preset_raw).name if preset_raw else "none"

        active_techniques: set[str] = set()
        sharpness_val: "float | None" = None

        if preset_raw:
            preset_p = Path(preset_raw)
            if not preset_p.is_absolute():
                preset_p = ini_path.parent / preset_raw
            if preset_p.exists():
                pdata = _rs_read_preset_params(preset_p)
                for sec_vals in pdata.values():
                    if isinstance(sec_vals, dict):
                        for t in _parse_techniques(sec_vals.get("Techniques", "")):
                            active_techniques.add(t)
                # CAS sharpness
                try:
                    sharpness_val = float(
                        pdata.get("CAS.fx", {}).get("CAS_SHARPENING_AMOUNT", "")
                    )
                except (ValueError, TypeError):
                    pass

        missing_effects = sorted(expected_techniques - active_techniques)
        heavy_active = [t for t in active_techniques if _effect_cost(t) == "high"]

        # Sharpness rating
        if sharpness_val is None:
            sharpness_rating = "unbekannt (CAS nicht aktiv oder Wert nicht lesbar)"
        elif sharpness_val < 0.30:
            sharpness_rating = "zu niedrig — Bild wirkt unscharf (<0.30)"
        elif sharpness_val > 0.75:
            sharpness_rating = "zu hoch — Halos/Ringing möglich (>0.75)"
        else:
            sharpness_rating = "gut"

        # Color depth — includes Vibrance (hue-safe saturation) in scoring
        _color_names = {"LiftGammaGain", "ColorMatrix", "Curves", "LUT", "Vibrance", "Tonemap"}
        active_color = sorted(active_techniques & _color_names)
        if not active_color:
            color_rating = "keine Farbkorrektur aktiv"
        elif len(active_color) >= 3:
            color_rating = "vollständige Farbkorrektur"
        elif len(active_color) >= 2:
            color_rating = "gute Farbkorrektur"
        else:
            color_rating = "minimale Farbkorrektur"

        # Image quality score 0–100
        score = 0
        preprocessor = general.get("PreprocessorDefinitions", "")
        depth_ok = "RESHADE_DEPTH_INPUT_IS_REVERSED=1" in preprocessor
        if perf_mode:   score += 15
        if depth_ok:    score += 10
        if not heavy_active: score += 15
        coverage = len(expected_techniques & active_techniques) / max(len(expected_techniques), 1)
        score += int(coverage * 35)
        if sharpness_val is not None and 0.30 <= sharpness_val <= 0.75:
            score += 15
        elif sharpness_val is not None:
            score += 5
        score += min(10, len(active_color) * 3)
        score = min(100, score)

        # Recommendations
        recommendations: list[str] = []
        if not perf_mode:
            recommendations.append(
                "[KRITISCH] Performance Mode deaktiviert — ReShade kostet VR-FPS. "
                "Aktiviere mit set_reshade_performance_mode(True)."
            )
        if heavy_active:
            recommendations.append(
                f"[HOCH] Schwere Effekte aktiv ({', '.join(heavy_active)}) — in VR vermeiden!"
            )
        if missing_effects:
            recommendations.append(
                f"[MITTEL] Für {tier}-Tier empfohlene Effekte fehlen: "
                f"{', '.join(missing_effects)}. "
                "Erstelle Preset mit create_reshade_vr_visual_preset()."
            )
        if sharpness_val is not None and sharpness_val < 0.30:
            recommendations.append(
                "[NIEDRIG] CAS-Wert zu niedrig — verwende tune_reshade_sharpness('medium') "
                "oder 'high'."
            )
        if sharpness_val is not None and sharpness_val > 0.75:
            recommendations.append(
                "[NIEDRIG] CAS-Wert sehr hoch — Halos möglich. "
                "Verwende tune_reshade_sharpness('high')."
            )
        if not active_color:
            recommendations.append(
                "[NIEDRIG] Keine Farbkorrektur aktiv — bessere Farben mit "
                "tune_reshade_colors() oder create_reshade_vr_visual_preset()."
            )
        if not recommendations:
            recommendations.append("ReShade ist optimal für VR konfiguriert!")

        return {
            "current_tier": tier,
            "tier_description": profile["description"],
            "current_preset": preset_name,
            "active_effects": sorted(active_techniques),
            "expected_effects_for_tier": sorted(expected_techniques),
            "missing_effects": missing_effects,
            "heavy_effects_active": heavy_active,
            "sharpness_value": sharpness_val,
            "sharpness_rating": sharpness_rating,
            "color_rating": color_rating,
            "active_color_effects": active_color,
            "image_quality_score": score,
            "performance_mode": perf_mode,
            "recommendations": recommendations,
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def apply_reshade_style(style: str) -> dict:
    """Wende einen visuellen Stil auf das aktuell aktive ReShade Preset an.

    Schreibt direkt die passenden Parameter-Werte in die Preset-Datei.
    ReShade übernimmt die Änderungen automatisch (kein Neustart nötig).

    Wann aufrufen: 'filmischer Look', 'Nachtmodus ReShade', 'Sonnenuntergang-Farben',
    'schärferes Bild', 'lebendige Farben', 'natürliche Farben', 'wärmeres Bild',
    'kälteres Bild', 'Tageslicht-Preset'.

    Args:
        style: 'cinematic'|'sharp'|'natural'|'vivid'|'night'|'day'|'sunset'|'dawn'|'dusk'
               cinematic = filmische Kontraste, warme Schatten
               sharp     = maximale Schärfe (CAS 0.75), klare Kanten
               natural   = natürliche Farben, minimaler Eingriff
               vivid     = lebendige Farben, hohe Sättigung
               night     = dunkleres Bild, kältere Töne (Nachtflüge)
               day       = helles Bild, leicht warm (Tagflüge)
               sunset    = orange/goldene Töne (Abendflüge)
               dawn      = warmes Morgenlicht, aufhellend (Morgenflüge)
               dusk      = kühles Blau-Violett (Dämmerungsflüge)

    Returns:
        status, style, parameters_changed, preset_updated
    """
    try:
        valid_styles = ("cinematic", "sharp", "natural", "vivid", "night", "day", "sunset",
                        "dawn", "dusk")
        if style not in valid_styles:
            return {
                "error": f"Ungültiger Stil '{style}'.",
                "valid_styles": list(valid_styles),
            }

        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        preset_path = _rs_get_active_preset_path(ini_path)
        if preset_path is None:
            return {
                "error": "Kein aktives Preset gefunden.",
                "hint": "Erstelle zuerst ein Preset mit create_reshade_vr_visual_preset().",
            }

        # Style parameter tables: {(section, key): value_string}
        style_params: dict = {
            "cinematic": {
                ("LiftGammaGain.fx", "RGB_Lift"):  "1.000000 0.998000 1.015000",
                ("LiftGammaGain.fx", "RGB_Gamma"): "0.975000 0.980000 1.000000",
                ("LiftGammaGain.fx", "RGB_Gain"):  "1.015000 1.008000 1.000000",
                ("Curves.fx", "Mode"):              "0",
                ("Curves.fx", "Formula"):           "4",
                ("Curves.fx", "Contrast"):          "0.350000",
                ("CAS.fx", "CAS_SHARPENING_AMOUNT"): "0.600000",
            },
            "sharp": {
                ("CAS.fx", "CAS_SHARPENING_AMOUNT"): "0.750000",
            },
            "natural": {
                ("Levels.fx", "BlackPoint"):         "3",
                ("Levels.fx", "WhitePoint"):         "252",
                ("LiftGammaGain.fx", "RGB_Lift"):    "1.000000 1.000000 1.002000",
                ("LiftGammaGain.fx", "RGB_Gamma"):   "0.995000 0.997000 1.000000",
                ("LiftGammaGain.fx", "RGB_Gain"):    "1.002000 1.001000 1.000000",
                ("Curves.fx", "Contrast"):           "0.100000",
            },
            "vivid": {
                ("ColorMatrix.fx", "ColorMatrix_Red"):   "0.840000 0.160000 0.000000",
                ("ColorMatrix.fx", "ColorMatrix_Green"): "0.040000 0.930000 0.030000",
                ("ColorMatrix.fx", "ColorMatrix_Blue"):  "0.000000 0.040000 0.960000",
                ("ColorMatrix.fx", "Strength"):          "0.500000",
            },
            "night": {
                ("Levels.fx", "BlackPoint"):             "15",
                ("Levels.fx", "WhitePoint"):             "230",
                ("LiftGammaGain.fx", "RGB_Lift"):        "0.995000 0.998000 1.020000",
                ("LiftGammaGain.fx", "RGB_Gamma"):       "0.988000 0.992000 1.008000",
                ("LiftGammaGain.fx", "RGB_Gain"):        "0.992000 0.995000 1.010000",
                ("Curves.fx", "Contrast"):               "0.300000",
                ("CAS.fx", "CAS_SHARPENING_AMOUNT"):     "0.500000",
            },
            "day": {
                ("Levels.fx", "BlackPoint"):             "5",
                ("Levels.fx", "WhitePoint"):             "248",
                ("LiftGammaGain.fx", "RGB_Lift"):        "1.005000 1.002000 1.000000",
                ("LiftGammaGain.fx", "RGB_Gamma"):       "1.005000 1.003000 0.995000",
                ("LiftGammaGain.fx", "RGB_Gain"):        "1.010000 1.006000 0.998000",
                ("Curves.fx", "Contrast"):               "0.180000",
                ("CAS.fx", "CAS_SHARPENING_AMOUNT"):     "0.550000",
            },
            "sunset": {
                ("LiftGammaGain.fx", "RGB_Lift"):        "1.020000 1.010000 0.990000",
                ("LiftGammaGain.fx", "RGB_Gamma"):       "1.025000 1.008000 0.975000",
                ("LiftGammaGain.fx", "RGB_Gain"):        "1.030000 1.015000 0.980000",
                ("ColorMatrix.fx", "ColorMatrix_Red"):   "0.870000 0.130000 0.000000",
                ("ColorMatrix.fx", "ColorMatrix_Green"): "0.060000 0.910000 0.030000",
                ("ColorMatrix.fx", "ColorMatrix_Blue"):  "0.000000 0.060000 0.940000",
                ("ColorMatrix.fx", "Strength"):          "0.450000",
                ("Curves.fx", "Contrast"):               "0.280000",
                ("CAS.fx", "CAS_SHARPENING_AMOUNT"):     "0.580000",
            },
            # Dawn: warm, brightening — early morning golden light, soft contrast
            "dawn": {
                ("Levels.fx", "BlackPoint"):             "3",
                ("Levels.fx", "WhitePoint"):             "250",
                ("LiftGammaGain.fx", "RGB_Lift"):        "1.012000 1.006000 0.995000",
                ("LiftGammaGain.fx", "RGB_Gamma"):       "1.018000 1.010000 0.985000",
                ("LiftGammaGain.fx", "RGB_Gain"):        "1.025000 1.012000 0.988000",
                ("Curves.fx", "Contrast"):               "0.150000",
                ("CAS.fx", "CAS_SHARPENING_AMOUNT"):     "0.520000",
            },
            # Dusk: cool twilight blue with slightly compressed tones
            "dusk": {
                ("Levels.fx", "BlackPoint"):             "10",
                ("Levels.fx", "WhitePoint"):             "235",
                ("LiftGammaGain.fx", "RGB_Lift"):        "0.996000 0.998000 1.018000",
                ("LiftGammaGain.fx", "RGB_Gamma"):       "0.990000 0.992000 1.012000",
                ("LiftGammaGain.fx", "RGB_Gain"):        "0.993000 0.996000 1.012000",
                ("Curves.fx", "Contrast"):               "0.250000",
                ("CAS.fx", "CAS_SHARPENING_AMOUNT"):     "0.520000",
            },
        }

        changes: dict = {}
        for (section, key), val in style_params[style].items():
            _rs_patch_preset_param(preset_path, section, key, val)
            changes[f"{section}/{key}"] = val

        return {
            "status": "ok",
            "style": style,
            "preset": preset_path.name,
            "parameters_changed": changes,
            "note": (
                f"Stil '{style}' auf '{preset_path.name}' angewendet. "
                "ReShade übernimmt Änderungen automatisch; "
                "falls nicht: Home-Taste → ReShade-Overlay → Preset neu laden."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# ReShade — new tools: time-of-day, addons, backup/restore, custom profiles
# ---------------------------------------------------------------------------

@mcp.tool()
def auto_apply_reshade_for_time(hour: "int | None" = None) -> dict:
    """Passt ReShade automatisch an die aktuelle Tageszeit an.

    Wählt automatisch den passenden visuellen Stil basierend auf der Uhrzeit
    und wendet ihn auf das aktive Preset an.

    Wann aufrufen: 'ReShade für jetzt anpassen', 'Tageszeit ReShade',
    'automatischer Modus', 'ReShade Nacht/Tag/Sonnenuntergang automatisch'.

    Args:
        hour: Stunde (0–23). Wenn None, wird die aktuelle Systemzeit verwendet.

    Returns:
        status, hour_used, time_period, style_applied, parameters_changed
    """
    try:
        import datetime
        if hour is None:
            hour = datetime.datetime.now().hour
        hour = int(hour) % 24

        # Map hour to time period and style
        if 0 <= hour <= 5:
            period, style = "night", "night"
        elif 6 <= hour <= 8:
            period, style = "dawn", "dawn"
        elif 9 <= hour <= 16:
            period, style = "day", "day"
        elif 17 <= hour <= 19:
            period, style = "sunset", "sunset"
        elif 20 <= hour <= 21:
            period, style = "dusk", "dusk"
        else:  # 22–23
            period, style = "night", "night"

        result = apply_reshade_style(style)
        if "error" in result:
            return result

        return {
            "status": "ok",
            "hour_used": hour,
            "time_period": period,
            "style_applied": style,
            "parameters_changed": result.get("parameters_changed", {}),
            "preset": result.get("preset", ""),
            "note": (
                f"Tageszeit {hour:02d}:xx → Periode '{period}' → Stil '{style}' angewendet. "
                + result.get("note", "")
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def get_reshade_addons() -> dict:
    """Listet alle installierten ReShade .addon-Dateien im MSFS-Ordner auf.

    ReShade 5.x+ unterstützt .addon-Dateien für erweiterte Funktionen (z.B. VR-Addons,
    OpenXR-Integration). Zeigt Name und Vorhandensein jedes Addons.

    Wann aufrufen: 'ReShade Addons', 'welche Addons sind installiert',
    'ReShade VR Addon', 'reshade-vr Addon'.

    Returns:
        addon_files: list of addon names found
        addon_dir: directory searched
        vr_addon_present: bool (reshade-vr or similar detected)
    """
    try:
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        base = ini_path.parent
        addon_files: list[dict] = []
        vr_addon_present = False

        for f in sorted(base.rglob("*.addon")):
            name = f.name
            addon_files.append({
                "name": name,
                "path": str(f),
                "size_kb": round(f.stat().st_size / 1024, 1),
            })
            if any(kw in name.lower() for kw in ("vr", "openxr", "openvr", "fsr")):
                vr_addon_present = True

        return {
            "status": "ok",
            "addon_dir": str(base),
            "addon_count": len(addon_files),
            "addon_files": addon_files,
            "vr_addon_present": vr_addon_present,
            "note": (
                "Keine .addon-Dateien gefunden. ReShade 5+ benötigt .addon-Dateien für "
                "erweiterte VR-Unterstützung." if not addon_files
                else f"{len(addon_files)} Addon(s) gefunden."
                     + (" VR-Addon erkannt." if vr_addon_present else "")
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def check_reshade_vr_compatibility() -> dict:
    """Prüft ob die installierte ReShade-Version VR (OpenXR) vollständig unterstützt.

    ReShade 5.0+ wird für volle OpenXR/VR-Unterstützung benötigt.
    Liest die DLL-Version von dxgi.dll und vergleicht mit der Mindestanforderung.

    Wann aufrufen: 'ReShade VR kompatibel', 'welche ReShade Version',
    'ReShade OpenXR Support', 'ReShade Version prüfen'.

    Returns:
        installed_version, version_ok (>=5.0), dll_found, recommendations
    """
    try:
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        base = ini_path.parent
        dll_path: "Path | None" = None
        dll_name = "none"
        for candidate in ("dxgi.dll", "d3d11.dll", "d3d12.dll"):
            p = base / candidate
            if p.exists():
                dll_path = p
                dll_name = candidate
                break

        if dll_path is None:
            return {
                "status": "warning",
                "installed_version": "unknown",
                "dll_found": "none",
                "version_ok": False,
                "recommendations": [
                    "Keine ReShade DLL gefunden. Reinstalliere ReShade von https://reshade.me.",
                ],
            }

        version_str = _get_file_version(dll_path)

        # Parse major version
        version_ok = False
        major = 0
        try:
            parts = version_str.split(".")
            major = int(parts[0])
            version_ok = major >= 5
        except Exception:
            pass

        recommendations: list[str] = []
        if not version_ok:
            recommendations.append(
                f"ReShade {version_str} ist veraltet — für volle VR/OpenXR-Unterstützung "
                "wird ReShade 5.0+ benötigt. Aktualisiere auf https://reshade.me."
            )
        else:
            recommendations.append(
                f"ReShade {version_str} unterstützt VR/OpenXR vollständig."
            )

        # Check for addon support (5.x feature)
        addon_support = major >= 5
        if addon_support:
            addon_count = len(list(base.rglob("*.addon")))
            if addon_count == 0:
                recommendations.append(
                    "ReShade 5+ Addon-Support verfügbar, aber keine .addon-Dateien gefunden. "
                    "Erwäge reshade-vr Addon für bessere OpenXR-Integration."
                )

        return {
            "status": "ok",
            "dll_found": dll_name,
            "dll_path": str(dll_path),
            "installed_version": version_str,
            "major_version": major,
            "version_ok": version_ok,
            "addon_support": addon_support,
            "vr_recommendation": (
                "VR-kompatibel: ReShade 5+ unterstützt OpenXR und addon-basierte VR-Features."
                if version_ok
                else "Update empfohlen: ReShade <5.0 hat eingeschränkte VR/OpenXR-Unterstützung."
            ),
            "recommendations": recommendations,
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def backup_reshade_preset(preset_name: str = "") -> dict:
    """Erstellt ein Backup des aktuell aktiven ReShade-Presets mit Datum im Dateinamen.

    Kopiert das aktive Preset als 'PresetName_backup_YYYYMMDD.ini' in denselben Ordner.

    Wann aufrufen: 'ReShade Preset sichern', 'Preset Backup erstellen',
    'aktuelles Preset speichern', 'ReShade Sicherung'.

    Args:
        preset_name: Optionaler Name für das Backup (leer = Name des aktiven Presets).

    Returns:
        status, source_preset, backup_path, backup_name
    """
    try:
        import datetime
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        preset_path = _rs_get_active_preset_path(ini_path)
        if preset_path is None:
            return {
                "error": "Kein aktives Preset gefunden.",
                "hint": "Erstelle zuerst ein Preset mit create_reshade_vr_visual_preset().",
            }

        ts = datetime.datetime.now().strftime("%Y%m%d")
        base_name = preset_name.strip() if preset_name.strip() else preset_path.stem
        # Sanitise
        base_name = "".join(c for c in base_name if c.isalnum() or c in ("_", "-"))
        backup_name = f"{base_name}_backup_{ts}.ini"
        backup_path = preset_path.parent / backup_name

        # Avoid overwriting existing backup in same day — add suffix
        _suffix = 0
        while backup_path.exists():
            _suffix += 1
            backup_path = preset_path.parent / f"{base_name}_backup_{ts}_{_suffix}.ini"

        _sh.copy2(str(preset_path), str(backup_path))

        return {
            "status": "ok",
            "source_preset": preset_path.name,
            "backup_path": str(backup_path),
            "backup_name": backup_path.name,
            "note": f"Backup erstellt: {backup_path.name}",
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def list_reshade_preset_backups() -> dict:
    """Listet alle verfügbaren ReShade-Preset-Backups mit Datum auf.

    Sucht nach *_backup_*.ini-Dateien im reshade-presets-Ordner.

    Wann aufrufen: 'ReShade Backups anzeigen', 'welche Preset-Backups gibt es',
    'ReShade Sicherungen Liste'.

    Returns:
        backups: list of backup files with name and date
    """
    try:
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        base = ini_path.parent
        search_dirs = [base, base / "reshade-presets"]
        backups: list[dict] = []
        import re as _re
        _pattern = _re.compile(r"_backup_(\d{8})", _re.IGNORECASE)

        for d in search_dirs:
            if not d.is_dir():
                continue
            for f in sorted(d.glob("*_backup_*.ini")):
                m = _pattern.search(f.stem)
                date_str = m.group(1) if m else "unknown"
                backups.append({
                    "name": f.name,
                    "path": str(f),
                    "date": date_str,
                    "size_kb": round(f.stat().st_size / 1024, 1),
                })

        return {
            "status": "ok",
            "backup_count": len(backups),
            "backups": backups,
            "note": (
                "Keine Backups gefunden. Erstelle eines mit backup_reshade_preset()."
                if not backups
                else f"{len(backups)} Backup(s) gefunden."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def restore_reshade_preset(backup_name: str) -> dict:
    """Stellt ein ReShade-Preset aus einem Backup wieder her und aktiviert es.

    Kopiert die Backup-Datei als aktives Preset und setzt ReShade.ini auf dieses Preset.

    Wann aufrufen: 'ReShade Preset wiederherstellen', 'Backup laden',
    'ReShade auf alten Stand zurücksetzen'.

    Args:
        backup_name: Name der Backup-Datei (mit oder ohne .ini).

    Returns:
        status, restored_from, preset_path, activated
    """
    try:
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        base = ini_path.parent
        # Find the backup file
        if not backup_name.lower().endswith(".ini"):
            backup_name += ".ini"

        backup_path: "Path | None" = None
        for d in (base, base / "reshade-presets"):
            candidate = d / backup_name
            if candidate.exists():
                backup_path = candidate
                break

        if backup_path is None:
            # Try case-insensitive search
            for d in (base, base / "reshade-presets"):
                if not d.is_dir():
                    continue
                for f in d.iterdir():
                    if f.name.lower() == backup_name.lower():
                        backup_path = f
                        break
                if backup_path:
                    break

        if backup_path is None:
            backups = list_reshade_preset_backups()
            return {
                "error": f"Backup '{backup_name}' nicht gefunden.",
                "available_backups": [b["name"] for b in backups.get("backups", [])],
            }

        # Derive restored preset name (strip _backup_YYYYMMDD suffix)
        import re as _re
        restored_stem = _re.sub(r"_backup_\d{8}(_\d+)?$", "", backup_path.stem,
                                 flags=_re.IGNORECASE)
        restored_name = f"{restored_stem}_restored.ini"
        restored_path = backup_path.parent / restored_name
        _sh.copy2(str(backup_path), str(restored_path))

        # Activate in ReShade.ini
        cp, _ = _rs_load_ini()
        if not cp.has_section("GENERAL"):
            cp.add_section("GENERAL")
        cp.set("GENERAL", "PresetPath", str(restored_path))
        _rs_save_ini(cp, ini_path)

        return {
            "status": "ok",
            "restored_from": backup_path.name,
            "preset_path": str(restored_path),
            "activated": True,
            "note": (
                f"Preset '{restored_path.name}' aus Backup '{backup_path.name}' "
                "wiederhergestellt und aktiviert. Home-Taste → ReShade-Overlay → Preset neu laden."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def save_reshade_profile_as(profile_name: str) -> dict:
    """Speichert das aktuelle ReShade-Preset unter einem benutzerdefinierten Namen.

    Kopiert das aktive Preset in den reshade-presets-Ordner mit dem gewünschten Namen
    und aktiviert es als neues Preset.

    Wann aufrufen: 'ReShade Preset als X speichern', 'Preset umbenennen',
    'eigenes ReShade Profil', 'Preset für dieses Flugzeug speichern'.

    Args:
        profile_name: Name für das neue Preset (ohne .ini).

    Returns:
        status, saved_path, activated
    """
    try:
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        preset_path = _rs_get_active_preset_path(ini_path)
        if preset_path is None:
            return {
                "error": "Kein aktives Preset gefunden.",
                "hint": "Erstelle zuerst ein Preset mit create_reshade_vr_visual_preset().",
            }

        # Sanitise name
        safe_name = "".join(c for c in profile_name.strip()
                            if c.isalnum() or c in (" ", "_", "-")).strip()
        safe_name = safe_name.replace(" ", "_")
        if not safe_name:
            return {"error": "Ungültiger Profilname. Nutze Buchstaben, Zahlen, _ oder -."}

        dest_dir = ini_path.parent / "reshade-presets"
        dest_dir.mkdir(exist_ok=True)
        dest_path = dest_dir / f"{safe_name}.ini"

        _sh.copy2(str(preset_path), str(dest_path))

        # Activate the new named preset
        cp, _ = _rs_load_ini()
        if not cp.has_section("GENERAL"):
            cp.add_section("GENERAL")
        cp.set("GENERAL", "PresetPath", str(dest_path))
        _rs_save_ini(cp, ini_path)

        return {
            "status": "ok",
            "original_preset": preset_path.name,
            "saved_as": dest_path.name,
            "saved_path": str(dest_path),
            "activated": True,
            "note": (
                f"Preset als '{safe_name}.ini' gespeichert und aktiviert. "
                "Home-Taste → ReShade-Overlay → Preset neu laden."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


@mcp.tool()
def list_reshade_custom_profiles() -> dict:
    """Listet alle benutzerdefinierten ReShade-Presets im GameCopilot/reshade-presets-Ordner.

    Zeigt alle .ini-Presets mit Größe und Datum — nützlich um eigene Profile zu verwalten.

    Wann aufrufen: 'ReShade Profile anzeigen', 'welche eigenen Presets gibt es',
    'Preset-Liste', 'ReShade Profil-Übersicht'.

    Returns:
        presets: list of preset names with metadata
        active_preset: currently active preset name
    """
    try:
        ini_path = _rs_locate_ini()
        if ini_path is None:
            return {"error": "ReShade.ini nicht gefunden."}

        base = ini_path.parent
        cp, _ = _rs_load_ini()
        general = cp["GENERAL"] if cp.has_section("GENERAL") else {}
        active_raw = general.get("PresetPath", "")
        active_name = Path(active_raw).name if active_raw else ""

        preset_dirs = [base / "reshade-presets", base / "Presets", base / "presets"]
        presets: list[dict] = []
        seen: set = set()

        for d in preset_dirs:
            if not d.is_dir():
                continue
            for f in sorted(d.glob("*.ini")):
                if f.name in seen:
                    continue
                seen.add(f.name)
                import datetime as _dt
                mtime = _dt.datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                presets.append({
                    "name": f.name,
                    "path": str(f),
                    "size_kb": round(f.stat().st_size / 1024, 1),
                    "modified": mtime,
                    "active": f.name == active_name,
                })

        return {
            "status": "ok",
            "preset_count": len(presets),
            "active_preset": active_name or "none",
            "presets": presets,
            "note": (
                "Keine Presets gefunden. Erstelle eines mit create_reshade_vr_visual_preset()."
                if not presets
                else f"{len(presets)} Preset(s) gefunden. Aktiv: {active_name or 'none'}."
            ),
        }
    except Exception as e:
        return {"error": str(e)}


if __name__ == "__main__":
    mcp.run()
