using System.IO;
using System.Text.Json;
using AzuHelper_v2.Models;

namespace AzuHelper_v2.Services;

public sealed class ConfigService
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true
    };

    public AppConfig Config { get; private set; } = AppConfig.Defaults();

    public async Task<AppConfig> LoadAsync(CancellationToken cancellationToken = default)
    {
        var path = PathService.ConfigPath();
        if (!File.Exists(path))
        {
            Config = AppConfig.Defaults();
            await SaveAsync(cancellationToken).ConfigureAwait(false);
            return Config;
        }

        try
        {
            var json = await File.ReadAllTextAsync(path, cancellationToken).ConfigureAwait(false);
            var loaded = JsonSerializer.Deserialize<AppConfig>(json, JsonOptions);
            Config = MergeDefaults(loaded);
            return Config;
        }
        catch
        {
            Config = AppConfig.Defaults();
            return Config;
        }
    }

    public async Task SaveAsync(CancellationToken cancellationToken = default)
    {
        var path = PathService.ConfigPath();
        var json = JsonSerializer.Serialize(Config, JsonOptions);
        await File.WriteAllTextAsync(path, json, cancellationToken).ConfigureAwait(false);
    }

    public void Update(AppConfig newConfig)
    {
        Config = MergeDefaults(newConfig);
    }

    private static readonly Dictionary<string, string> RegionMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["BW"] = "Baden-Württemberg",
        ["BADEN-WÜRTTEMBERG"] = "Baden-Württemberg",
        ["BY"] = "Bayern",
        ["BAYERN"] = "Bayern",
        ["BE"] = "Berlin",
        ["BERLIN"] = "Berlin",
        ["BB"] = "Brandenburg",
        ["BRANDENBURG"] = "Brandenburg",
        ["HB"] = "Bremen",
        ["BREMEN"] = "Bremen",
        ["HH"] = "Hamburg",
        ["HAMBURG"] = "Hamburg",
        ["HE"] = "Hessen",
        ["HESSEN"] = "Hessen",
        ["MV"] = "Mecklenburg-Vorpommern",
        ["MECKLENBURG-VORPOMMERN"] = "Mecklenburg-Vorpommern",
        ["NI"] = "Niedersachsen",
        ["NIEDERSACHSEN"] = "Niedersachsen",
        ["NW"] = "Nordrhein-Westfalen",
        ["NORDRHEIN-WESTFALEN"] = "Nordrhein-Westfalen",
        ["RP"] = "Rheinland-Pfalz",
        ["RHEINLAND-PFALZ"] = "Rheinland-Pfalz",
        ["SL"] = "Saarland",
        ["SAARLAND"] = "Saarland",
        ["SN"] = "Sachsen",
        ["SACHSEN"] = "Sachsen",
        ["ST"] = "Sachsen-Anhalt",
        ["SACHSEN-ANHALT"] = "Sachsen-Anhalt",
        ["SH"] = "Schleswig-Holstein",
        ["SCHLESWIG-HOLSTEIN"] = "Schleswig-Holstein",
        ["TH"] = "Thüringen",
        ["THÜRINGEN"] = "Thüringen"
    };

    private static AppConfig MergeDefaults(AppConfig? loaded)
    {
        var d = AppConfig.Defaults();
        if (loaded is null)
        {
            return d;
        }

        loaded.Name = string.IsNullOrWhiteSpace(loaded.Name) ? d.Name : loaded.Name;
        loaded.Vorname = string.IsNullOrWhiteSpace(loaded.Vorname) ? d.Vorname : loaded.Vorname;
        loaded.Stammnummer = string.IsNullOrWhiteSpace(loaded.Stammnummer) ? d.Stammnummer : loaded.Stammnummer;
        loaded.Berufsgruppe = string.IsNullOrWhiteSpace(loaded.Berufsgruppe) ? d.Berufsgruppe : loaded.Berufsgruppe;
        loaded.Ausbilder = string.IsNullOrWhiteSpace(loaded.Ausbilder) ? d.Ausbilder : loaded.Ausbilder;
        loaded.Arbeitsbeginn = string.IsNullOrWhiteSpace(loaded.Arbeitsbeginn) ? d.Arbeitsbeginn : loaded.Arbeitsbeginn;
        loaded.Arbeitsende = string.IsNullOrWhiteSpace(loaded.Arbeitsende) ? d.Arbeitsende : loaded.Arbeitsende;
        loaded.QuickFillStart = string.IsNullOrWhiteSpace(loaded.QuickFillStart) ? d.QuickFillStart : loaded.QuickFillStart;
        loaded.QuickFillEnd = string.IsNullOrWhiteSpace(loaded.QuickFillEnd) ? d.QuickFillEnd : loaded.QuickFillEnd;
        var durationHours = loaded.QuickFillDurationHours > 0
            ? loaded.QuickFillDurationHours
            : (loaded.QuickFillDurationMinutes > 0 ? loaded.QuickFillDurationMinutes / 60d : d.QuickFillDurationHours);

        loaded.QuickFillDurationHours = durationHours;
        loaded.QuickFillDurationMinutes = loaded.QuickFillDurationMinutes <= 0
            ? (int)Math.Round(durationHours * 60)
            : loaded.QuickFillDurationMinutes;

        var anyQuickFillDaySet = loaded.QuickFillMonday || loaded.QuickFillTuesday || loaded.QuickFillWednesday || loaded.QuickFillThursday || loaded.QuickFillFriday;
        if (!anyQuickFillDaySet)
        {
            loaded.QuickFillMonday = d.QuickFillMonday;
            loaded.QuickFillTuesday = d.QuickFillTuesday;
            loaded.QuickFillWednesday = d.QuickFillWednesday;
            loaded.QuickFillThursday = d.QuickFillThursday;
            loaded.QuickFillFriday = d.QuickFillFriday;
        }
        loaded.Region = NormalizeRegion(string.IsNullOrWhiteSpace(loaded.Region) ? d.Region : loaded.Region, d.Region);
        loaded.MailTo = string.IsNullOrWhiteSpace(loaded.MailTo) ? d.MailTo : loaded.MailTo;
        loaded.MailCc ??= d.MailCc;
        loaded.OutlookSignaturName ??= d.OutlookSignaturName;
        loaded.MailSubjectTemplate = string.IsNullOrWhiteSpace(loaded.MailSubjectTemplate) ? d.MailSubjectTemplate : loaded.MailSubjectTemplate;
        loaded.MailBodyTemplate = string.IsNullOrWhiteSpace(loaded.MailBodyTemplate) ? d.MailBodyTemplate : loaded.MailBodyTemplate;
        loaded.FileNameTemplate = string.IsNullOrWhiteSpace(loaded.FileNameTemplate) ? d.FileNameTemplate : loaded.FileNameTemplate;
        loaded.GitHubOwner = string.IsNullOrWhiteSpace(loaded.GitHubOwner) ? d.GitHubOwner : loaded.GitHubOwner;
        loaded.GitHubRepo = string.IsNullOrWhiteSpace(loaded.GitHubRepo) ? d.GitHubRepo : loaded.GitHubRepo;
        loaded.GitHubAssetName = string.IsNullOrWhiteSpace(loaded.GitHubAssetName) ? d.GitHubAssetName : loaded.GitHubAssetName;

        return loaded;
    }

    private static string NormalizeRegion(string region, string defaultRegion)
    {
        if (string.IsNullOrWhiteSpace(region))
        {
            return defaultRegion;
        }

        var key = region.Trim();
        return RegionMap.TryGetValue(key, out var mapped) ? mapped : key;
    }
}
