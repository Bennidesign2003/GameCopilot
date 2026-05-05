using System;
using System.Collections.Generic;
using System.IO;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using SkiaSharp;

namespace GameCopilot.Controls;

/// <summary>
/// Image control that plays animated GIFs.
/// Uses SkiaSharp SKCodec to decode frames, shared cache across instances.
/// </summary>
public class GifImage : Image
{
    // Shared frame cache so multiple instances of the same GIF don't duplicate memory
    private static readonly Dictionary<string, GifFrameData> FrameCache = new();

    private DispatcherTimer? _timer;
    private int _currentFrame;
    private GifFrameData? _data;

    public static readonly StyledProperty<string?> SourceUriProperty =
        AvaloniaProperty.Register<GifImage, string?>(nameof(SourceUri));

    public string? SourceUri
    {
        get => GetValue(SourceUriProperty);
        set => SetValue(SourceUriProperty, value);
    }

    static GifImage()
    {
        SourceUriProperty.Changed.AddClassHandler<GifImage>((x, _) => x.LoadGif());
    }

    private void LoadGif()
    {
        Stop();
        _data = null;

        var uriStr = SourceUri;
        if (string.IsNullOrEmpty(uriStr)) return;

        // Check cache first
        if (FrameCache.TryGetValue(uriStr, out var cached))
        {
            _data = cached;
            StartAnimation();
            return;
        }

        try
        {
            var uri = new Uri(uriStr);
            using var stream = AssetLoader.Open(uri);
            using var ms = new MemoryStream();
            stream.CopyTo(ms);

            using var skData = SKData.CreateCopy(ms.ToArray());
            using var codec = SKCodec.Create(skData);
            if (codec == null) return;

            var frameCount = codec.FrameCount;
            if (frameCount <= 1)
            {
                ms.Position = 0;
                Source = new Bitmap(ms);
                return;
            }

            var info = new SKImageInfo(codec.Info.Width, codec.Info.Height,
                SKColorType.Bgra8888, SKAlphaType.Premul);

            var frames = new List<Bitmap>(frameCount);
            var durations = new List<int>(frameCount);

            for (int i = 0; i < frameCount; i++)
            {
                using var bmp = new SKBitmap(info);
                codec.GetPixels(info, bmp.GetPixels(), new SKCodecOptions(i));

                using var img = SKImage.FromBitmap(bmp);
                using var encoded = img.Encode(SKEncodedImageFormat.Png, 90);
                using var pngStream = new MemoryStream(encoded.ToArray());
                frames.Add(new Bitmap(pngStream));

                var fi = codec.FrameInfo[i];
                durations.Add(fi.Duration > 0 ? fi.Duration : 100);
            }

            _data = new GifFrameData(frames, durations);
            FrameCache[uriStr] = _data;
            StartAnimation();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"GifImage: {ex.Message}");
        }
    }

    private void StartAnimation()
    {
        if (_data == null || _data.Frames.Count <= 1) return;

        _currentFrame = 0;
        Source = _data.Frames[0];

        _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(_data.Durations[0])
        };
        _timer.Tick += OnTick;
        _timer.Start();
    }

    private void OnTick(object? sender, EventArgs e)
    {
        if (_data == null) return;
        _currentFrame = (_currentFrame + 1) % _data.Frames.Count;
        Source = _data.Frames[_currentFrame];

        if (_timer != null)
            _timer.Interval = TimeSpan.FromMilliseconds(_data.Durations[_currentFrame]);
    }

    private void Stop()
    {
        _timer?.Stop();
        _timer = null;
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        Stop();
        base.OnDetachedFromVisualTree(e);
    }

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        if (_data != null && _timer == null)
            StartAnimation();
    }

    private class GifFrameData
    {
        public List<Bitmap> Frames { get; }
        public List<int> Durations { get; }

        public GifFrameData(List<Bitmap> frames, List<int> durations)
        {
            Frames = frames;
            Durations = durations;
        }
    }
}
