using System;
using System.Collections.Generic;
using System.Globalization;
using Avalonia;
using Avalonia.Data.Converters;
using Avalonia.Media;
using Avalonia.Media.Imaging;

namespace GameCopilot.Converters;

public class BoolToStatusTextConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool b)
            return b ? "Gefunden" : "Nicht gefunden";
        return "Unbekannt";
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public class BoolToColorConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool b)
            return b ? SolidColorBrush.Parse("#00FF88") : SolidColorBrush.Parse("#FF4444");
        return SolidColorBrush.Parse("#B0B0B0");
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public class PercentConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is double d)
            return $"{(int)(d * 100)}%";
        return "0%";
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>
/// Converts CurrentPage string to IsVisible bool.
/// ConverterParameter = page name to match.
/// </summary>
public class PageVisibleConverter : IValueConverter
{
    public static readonly PageVisibleConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is string current && parameter is string page)
            return current == page;
        return false;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>
/// Returns Background brush for nav button active state.
/// WPF used dynamic Style switching; Avalonia uses property bindings.
/// </summary>
public class NavBgConverter : IValueConverter
{
    public static readonly NavBgConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool active && active)
            return SolidColorBrush.Parse("#1d2023");
        return SolidColorBrush.Parse("#00000000"); // Transparent
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public class NavFgConverter : IValueConverter
{
    public static readonly NavFgConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool active && active)
            return SolidColorBrush.Parse("#69daff");
        return SolidColorBrush.Parse("#aaabad");
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public class CategoryBgConverter : IValueConverter
{
    public static readonly CategoryBgConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is string current && parameter is string target && current == target)
            return SolidColorBrush.Parse("#69daff");
        return SolidColorBrush.Parse("#1d2023");
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public class CategoryFgConverter : IValueConverter
{
    public static readonly CategoryFgConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is string current && parameter is string target && current == target)
            return SolidColorBrush.Parse("#004a5d");
        return SolidColorBrush.Parse("#aaabad");
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public class CategoryImageConverter : IValueConverter
{
    public static readonly CategoryImageConverter Instance = new();

    private static readonly Dictionary<string, Bitmap?> Cache = new();

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var category = value as string ?? "All";
        var uri = category switch
        {
            "Scenery" => "avares://GameCopilot/Assets/background_landscape.png",
            _ => "avares://GameCopilot/Assets/background_flight.png"
        };

        if (!Cache.TryGetValue(uri, out var bitmap))
        {
            try
            {
                var assets = Avalonia.Platform.AssetLoader.Open(new Uri(uri));
                bitmap = new Bitmap(assets);
            }
            catch
            {
                bitmap = null;
            }
            Cache[uri] = bitmap;
        }

        return bitmap;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>
/// Compares a preset Name to ActivePresetName → true when they match.
/// Used via MultiBinding in preset list to show "AKTIV" badge.
/// </summary>
public class ActivePresetConverter : IMultiValueConverter
{
    public static readonly ActivePresetConverter Instance = new();

    public object Convert(IList<object?> values, Type targetType, object? parameter, CultureInfo culture)
    {
        if (values.Count == 2 && values[0] is string name && values[1] is string active)
            return name == active;
        return false;
    }
}

/// <summary>
/// Converts a 0.0-1.0 value to a percentage width string for Grid ColumnDefinitions.
/// Parameter = total columns (default 200). Returns "*" proportional column def.
/// </summary>
public class BarWidthConverter : IValueConverter
{
    public static readonly BarWidthConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is double d)
        {
            var maxWidth = 200.0;
            if (parameter is string s && double.TryParse(s, out var mw))
                maxWidth = mw;
            return Math.Max(0, Math.Min(maxWidth, d * maxWidth));
        }
        return 0.0;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>
/// Converts bool to checkmark string for shader badge display.
/// </summary>
public class BoolToCheckConverter : IValueConverter
{
    public static readonly BoolToCheckConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool b)
            return b ? "AN" : "AUS";
        return "AUS";
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public class BoolToShaderColorConverter : IValueConverter
{
    public static readonly BoolToShaderColorConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool b && b)
            return SolidColorBrush.Parse("#69daff");
        return SolidColorBrush.Parse("#46484a");
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>Action button background: primary = cyan filled, secondary = transparent.</summary>
public class ActionBtnBgConverter : IValueConverter
{
    public static readonly ActionBtnBgConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool b && b)
            return SolidColorBrush.Parse("#69daff");
        return SolidColorBrush.Parse("#00000000");
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>Action button foreground: primary = dark, secondary = cyan.</summary>
public class ActionBtnFgConverter : IValueConverter
{
    public static readonly ActionBtnFgConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool b && b)
            return SolidColorBrush.Parse("#004a5d");
        return SolidColorBrush.Parse("#69daff");
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>Action button border: primary = 0, secondary = 1px cyan.</summary>
public class ActionBtnBorderConverter : IValueConverter
{
    public static readonly ActionBtnBorderConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool b && b)
            return new Thickness(0);
        return new Thickness(1);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>
/// Converts SplashProgress (0.0-1.0) to loading bar width (0-400px).
/// Replaces WPF DoubleAnimation on LoadingBar.Width.
/// </summary>
public class ProgressWidthConverter : IValueConverter
{
    public static readonly ProgressWidthConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is double d)
            return d * 400.0;
        return 0.0;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>
/// bool ShowApiKeyText → PasswordChar: true = '\0' (plain), false = '•' (masked).
/// Used on the Codex API key TextBox.
/// </summary>
public class PasswordCharConverter : IValueConverter
{
    public static readonly PasswordCharConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? '\0' : '•';

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>
/// bool ShowApiKeyText → eye icon: true = 👁 (visible), false = 🙈 (hidden).
/// </summary>
public class EyeIconConverter : IValueConverter
{
    public static readonly EyeIconConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? "👁" : "🙈";

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
