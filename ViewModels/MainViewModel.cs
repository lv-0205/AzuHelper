using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using AzuHelper_v2.Models;
using AzuHelper_v2.Services;

namespace AzuHelper_v2.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly ConfigService _configService;
    private readonly GitHubUpdateService _updateService;
    private AppConfig _config = AppConfig.Defaults();
    private Task? _initializeTask;

    private string _weekLabel = string.Empty;
    private string _mailTo = string.Empty;
    private string _mailCc = string.Empty;
    private bool _openInOutlook = true;
    private bool _isCreatingEmail;
    private double _emailProgress;
    private string _emailProgressText = string.Empty;
    private string _quickFillError = string.Empty;
    private bool _isUpdating;
    private string _updateStatusText = string.Empty;

    public MainViewModel() : this(new ConfigService())
    {
    }

    public MainViewModel(ConfigService configService)
    {
        _configService = configService;
        _updateService = new GitHubUpdateService();

        Days = new ObservableCollection<DayEntry>();
        EmailMessages = new ObservableCollection<string>();
        EmailMessages.CollectionChanged += EmailMessagesOnCollectionChanged;

        OpenSettingsCommand = new RelayCommand(OpenSettings);
        CreateEmailCommand = new RelayCommand(CreateEmailAsync, () => !IsCreatingEmail);
        QuickFillCommand = new RelayCommand(QuickFill);
        CheckUpdateCommand = new RelayCommand(CheckForUpdatesAsync, () => !IsUpdating);
    }

    public ObservableCollection<DayEntry> Days { get; }

    public ObservableCollection<string> EmailMessages { get; }

    public bool HasEmailMessages => EmailMessages.Count > 0;

    public bool HasUpdateStatus => !string.IsNullOrWhiteSpace(_updateStatusText);

    public string WeekLabel
    {
        get => _weekLabel;
        private set
        {
            if (_weekLabel == value) return;
            _weekLabel = value;
            OnPropertyChanged();
        }
    }

    public string MailTo
    {
        get => _mailTo;
        set
        {
            if (_mailTo == value) return;
            _mailTo = value;
            OnPropertyChanged();
        }
    }

    public string MailCc
    {
        get => _mailCc;
        set
        {
            if (_mailCc == value) return;
            _mailCc = value;
            OnPropertyChanged();
        }
    }

    public bool OpenInOutlook
    {
        get => _openInOutlook;
        set
        {
            if (_openInOutlook == value) return;
            _openInOutlook = value;
            OnPropertyChanged();
        }
    }

    public bool IsCreatingEmail
    {
        get => _isCreatingEmail;
        private set
        {
            if (_isCreatingEmail == value) return;
            _isCreatingEmail = value;
            OnPropertyChanged();
            CreateEmailCommand.RaiseCanExecuteChanged();
        }
    }

    public bool IsUpdating
    {
        get => _isUpdating;
        private set
        {
            if (_isUpdating == value) return;
            _isUpdating = value;
            OnPropertyChanged();
            CheckUpdateCommand.RaiseCanExecuteChanged();
        }
    }

    public double EmailProgress
    {
        get => _emailProgress;
        private set
        {
            if (Math.Abs(_emailProgress - value) < 0.01) return;
            _emailProgress = value;
            OnPropertyChanged();
        }
    }

    public string EmailProgressText
    {
        get => _emailProgressText;
        private set
        {
            if (_emailProgressText == value) return;
            _emailProgressText = value;
            OnPropertyChanged();
        }
    }

    public string QuickFillError
    {
        get => _quickFillError;
        private set
        {
            if (_quickFillError == value) return;
            _quickFillError = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasQuickFillError));
        }
    }

    public bool HasQuickFillError => !string.IsNullOrWhiteSpace(_quickFillError);

    public string UpdateStatusText
    {
        get => _updateStatusText;
        private set
        {
            if (_updateStatusText == value) return;
            _updateStatusText = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasUpdateStatus));
        }
    }

    public RelayCommand OpenSettingsCommand { get; }
    public RelayCommand CreateEmailCommand { get; }
    public RelayCommand QuickFillCommand { get; }
    public RelayCommand CheckUpdateCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public Task EnsureInitializedAsync()
    {
        return _initializeTask ??= InitializeAsync();
    }

    private async Task InitializeAsync()
    {
        var cfg = await _configService.LoadAsync().ConfigureAwait(false);
        _config = cfg;

        await Application.Current.Dispatcher.InvokeAsync(() =>
        {
            MailTo = cfg.MailTo;
            MailCc = cfg.MailCc;

            var now = DateOnly.FromDateTime(DateTime.Today);
            var startOfWeek = now.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
            if (DateTime.Today.DayOfWeek == DayOfWeek.Sunday)
            {
                startOfWeek = now.AddDays(-6);
            }

            var week = ISOWeek.GetWeekOfYear(DateTime.Today);
            WeekLabel = $"Kalenderwoche {week}";

            Days.Clear();
            for (var i = 0; i < 5; i++)
            {
                var date = startOfWeek.AddDays(i);

                var day = new DayEntry
                {
                    DayName = CultureInfo.CurrentCulture.DateTimeFormat.AbbreviatedDayNames[(int)(DayOfWeek)(((int)DayOfWeek.Monday + i) % 7)],
                    Date = date,
                    Enabled = false,
                    StartTime = cfg.Arbeitsbeginn,
                    EndTime = cfg.Arbeitsende,
                    Duration = string.Empty
                };

                day.PropertyChanged += DayOnPropertyChanged;
                day.RecalculateDuration(TimeSpan.FromMinutes(60));
                Days.Add(day);
            }
        });
    }

    private void DayOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not DayEntry day) return;

        if (e.PropertyName is nameof(DayEntry.Enabled) or nameof(DayEntry.StartTime) or nameof(DayEntry.EndTime))
        {
            day.RecalculateDuration(TimeSpan.FromMinutes(60));
        }
    }

    private async Task CheckForUpdatesAsync()
    {
        if (IsUpdating)
        {
            return;
        }

        await Application.Current.Dispatcher.InvokeAsync(() =>
        {
            IsUpdating = true;
            UpdateStatusText = "Suche nach Updates...";
        });

        try
        {
            var progress = new Progress<string>(message =>
                _ = Application.Current.Dispatcher.InvokeAsync(() => UpdateStatusText = message));

            var result = await _updateService.CheckAndPrepareUpdateAsync(_config, progress).ConfigureAwait(false);

            await Application.Current.Dispatcher.InvokeAsync(() => UpdateStatusText = result.Message);

            if (result.Kind == UpdateResultKind.ReadyToApply && !string.IsNullOrWhiteSpace(result.ScriptPath))
            {
                _updateService.LaunchUpdateScript(result.ScriptPath);
                await Application.Current.Dispatcher.InvokeAsync(() => Application.Current.Shutdown());
            }
        }
        catch
        {
            await Application.Current.Dispatcher.InvokeAsync(() => UpdateStatusText = "Update fehlgeschlagen.");
        }
        finally
        {
            await Application.Current.Dispatcher.InvokeAsync(() => IsUpdating = false);
        }
    }

    private void OpenSettings()
    {
        var oldStart = _config.Arbeitsbeginn;
        var oldEnd = _config.Arbeitsende;

        var clone = new AppConfig
        {
            Name = _config.Name,
            Vorname = _config.Vorname,
            Stammnummer = _config.Stammnummer,
            Berufsgruppe = _config.Berufsgruppe,
            Ausbilder = _config.Ausbilder,
            Arbeitsbeginn = _config.Arbeitsbeginn,
            Arbeitsende = _config.Arbeitsende,
            QuickFillStart = _config.QuickFillStart,
            QuickFillEnd = _config.QuickFillEnd,
            QuickFillDurationHours = _config.QuickFillDurationHours,
            QuickFillDurationMinutes = _config.QuickFillDurationMinutes,
            QuickFillMonday = _config.QuickFillMonday,
            QuickFillTuesday = _config.QuickFillTuesday,
            QuickFillWednesday = _config.QuickFillWednesday,
            QuickFillThursday = _config.QuickFillThursday,
            QuickFillFriday = _config.QuickFillFriday,
            Region = _config.Region,
            MailTo = _config.MailTo,
            MailCc = _config.MailCc,
            OutlookSignaturName = _config.OutlookSignaturName,
            MailSubjectTemplate = _config.MailSubjectTemplate,
            MailBodyTemplate = _config.MailBodyTemplate,
            FileNameTemplate = _config.FileNameTemplate
        };

        var dlg = new SettingsDialog(clone)
        {
            Owner = Application.Current.MainWindow
        };

        if (dlg.ShowDialog() != true)
        {
            return;
        }

        _configService.Update(dlg.Config);
        _config = _configService.Config;
        _ = _configService.SaveAsync();

        MailTo = _config.MailTo;
        MailCc = _config.MailCc;

        // If a day still matches the old defaults, update it to the new defaults.
        foreach (var day in Days)
        {
            if (string.IsNullOrWhiteSpace(day.StartTime) || day.StartTime == oldStart)
            {
                day.StartTime = _config.Arbeitsbeginn;
            }

            if (string.IsNullOrWhiteSpace(day.EndTime) || day.EndTime == oldEnd)
            {
                day.EndTime = _config.Arbeitsende;
            }

            day.RecalculateDuration(TimeSpan.FromMinutes(60));
        }
    }

    private void QuickFill()
    {
        var startText = string.IsNullOrWhiteSpace(_config.QuickFillStart) ? _config.Arbeitsbeginn : _config.QuickFillStart;
        var endText = string.IsNullOrWhiteSpace(_config.QuickFillEnd) ? _config.Arbeitsende : _config.QuickFillEnd;

        if (!TimeOnly.TryParse(startText, out var start))
        {
            QuickFillError = "Quickfill-Start ungültig. Bitte Einstellungen prüfen.";
            return;
        }

        if (!TimeOnly.TryParse(endText, out var end))
        {
            QuickFillError = "Quickfill-Ende ungültig. Bitte Einstellungen prüfen.";
            return;
        }

        if (end <= start)
        {
            QuickFillError = "Quickfill-Ende muss nach dem Start liegen.";
            return;
        }

        if (!(_config.QuickFillMonday || _config.QuickFillTuesday || _config.QuickFillWednesday || _config.QuickFillThursday || _config.QuickFillFriday))
        {
            QuickFillError = "Quickfill: Bitte mindestens einen Tag auswählen.";
            return;
        }

        QuickFillError = string.Empty;

        var startString = start.ToString("HH:mm", CultureInfo.InvariantCulture);
        var endString = end.ToString("HH:mm", CultureInfo.InvariantCulture);

        foreach (var day in Days)
        {
            if (!IsQuickFillEnabledForDay(day.Date.DayOfWeek))
            {
                continue;
            }

            day.Enabled = true;
            day.StartTime = startString;
            day.EndTime = endString;
            day.RecalculateDuration(TimeSpan.FromMinutes(60));
        }
    }

    private bool IsQuickFillEnabledForDay(DayOfWeek dayOfWeek)
    {
        return dayOfWeek switch
        {
            DayOfWeek.Monday => _config.QuickFillMonday,
            DayOfWeek.Tuesday => _config.QuickFillTuesday,
            DayOfWeek.Wednesday => _config.QuickFillWednesday,
            DayOfWeek.Thursday => _config.QuickFillThursday,
            DayOfWeek.Friday => _config.QuickFillFriday,
            _ => false
        };
    }

    private async Task CreateEmailAsync()
    {
        if (IsCreatingEmail)
        {
            return;
        }

        await Application.Current.Dispatcher.InvokeAsync(() =>
        {
            IsCreatingEmail = true;
            EmailMessages.Clear();
            EmailProgress = 0;
            EmailProgressText = "Starte...";
            EmailMessages.Add(EmailProgressText);
        });

        var to = MailTo?.Trim() ?? string.Empty;
        var cc = MailCc?.Trim() ?? string.Empty;

        var week = ISOWeek.GetWeekOfYear(DateTime.Today);
        var tokens = BuildTokens(_config, week);
        var subject = TemplateService.Apply(_config.MailSubjectTemplate, tokens);
        var body = TemplateService.Apply(_config.MailBodyTemplate, tokens);

        try
        {
            await UpdateEmailProgressAsync(10, "Erstelle Excel-Datei...");

            var excelPath = await Task.Run(() => ExcelExportService.CreateTimesheet(_config, Days)).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                await AddEmailMessageAsync("Fehler: Excel-Vorlage nicht gefunden oder konnte nicht erstellt werden.");
                await ShowErrorAsync("Excel-Vorlage nicht gefunden oder konnte nicht erstellt werden.");
                return;
            }

            await UpdateEmailProgressAsync(60, "Erstelle E-Mail...");

            var (success, errorMessage) = await Task.Run(() =>
            {
                var result = OutlookEmailService.SendEmail(
                    to,
                    cc,
                    subject,
                    body,
                    excelPath,
                    OpenInOutlook,
                    _config.OutlookSignaturName,
                    out var error);
                return (result, error);
            }).ConfigureAwait(false);

            if (!success)
            {
                await AddEmailMessageAsync($"Fehler: {errorMessage ?? "Fehler beim Erstellen der E-Mail."}");
                await ShowErrorAsync(errorMessage ?? "Fehler beim Erstellen der E-Mail.");
                return;
            }

            await UpdateEmailProgressAsync(100, "E-Mail erstellt.");
        }
        finally
        {
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                IsCreatingEmail = false;
                EmailProgress = 0;
                EmailProgressText = string.Empty;
            });
        }
    }

    private Task UpdateEmailProgressAsync(double progress, string text) => Application.Current.Dispatcher.InvokeAsync(() =>
    {
        EmailProgress = progress;
        EmailProgressText = text;
        EmailMessages.Add(text);
    }).Task;

    private Task AddEmailMessageAsync(string message) => Application.Current.Dispatcher.InvokeAsync(() =>
    {
        EmailMessages.Add(message);
    }).Task;

    private Task ShowErrorAsync(string message) => Application.Current.Dispatcher.InvokeAsync(() =>
        MessageBox.Show(message, "AzuHelper", MessageBoxButton.OK, MessageBoxImage.Error)).Task;

    private void EmailMessagesOnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(HasEmailMessages));
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    private static Dictionary<string, string> BuildTokens(AppConfig config, int week)
    {
        var today = DateTime.Today;
        var fullname = $"{config.Vorname} {config.Name}".Trim();
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["name"] = config.Name,
            ["vorname"] = config.Vorname,
            ["fullname"] = fullname,
            ["week"] = week.ToString(CultureInfo.InvariantCulture),
            ["kw"] = week.ToString(CultureInfo.InvariantCulture),
            ["year"] = today.Year.ToString(CultureInfo.InvariantCulture),
            ["date"] = today.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture)
        };
    }
}
