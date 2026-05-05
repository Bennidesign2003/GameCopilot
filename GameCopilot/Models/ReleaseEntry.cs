using System;
using System.Collections.Generic;

namespace GameCopilot.Models;

public class ReleaseEntry
{
    public string TagName { get; set; } = "";
    public string Title { get; set; } = "";
    public string Label { get; set; } = ""; // e.g. "CURRENT BUILD", "STABLE RELEASE"
    public string Date { get; set; } = "";
    public string Body { get; set; } = "";
    public List<string> Changes { get; set; } = new();
    public bool IsCurrent { get; set; }
}
