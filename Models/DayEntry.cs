using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AzuHelper_v2.Models;

public sealed class DayEntry : INotifyPropertyChanged
{
    private bool _enabled;
    private string _startTime = string.Empty;
    private string _endTime = string.Empty;
    private string _duration = string.Empty;

    public required string DayName { get; init; }
    public required DateOnly Date { get; init; }

    public string DateLabel => Date.ToString("dd.MM");

    public bool Enabled
    {
        get => _enabled;
        set
        {
            if (_enabled == value) return;
            _enabled = value;
            OnPropertyChanged();
        }
    }

    public string StartTime
    {
        get => _startTime;
        set
        {
            if (_startTime == value) return;
            _startTime = value;
            OnPropertyChanged();
        }
    }

    public string EndTime
    {
        get => _endTime;
        set
        {
            if (_endTime == value) return;
            _endTime = value;
            OnPropertyChanged();
        }
    }

    public string Duration
    {
        get => _duration;
        set
        {
            if (_duration == value) return;
            _duration = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public void RecalculateDuration(TimeSpan fixedBreak)
    {
        if (!Enabled)
        {
            Duration = string.Empty;
            return;
        }

        if (!TimeOnly.TryParse(StartTime, out var start) || !TimeOnly.TryParse(EndTime, out var end))
        {
            Duration = string.Empty;
            return;
        }

        if (end <= start)
        {
            Duration = string.Empty;
            return;
        }

        var delta = end.ToTimeSpan() - start.ToTimeSpan() - fixedBreak;
        if (delta <= TimeSpan.Zero)
        {
            Duration = string.Empty;
            return;
        }

        Duration = $"{(int)delta.TotalHours}:{delta.Minutes:00}";
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
