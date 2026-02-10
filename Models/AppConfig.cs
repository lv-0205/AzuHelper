using System.Text.Json.Serialization;

namespace AzuHelper_v2.Models;

public sealed class AppConfig
{
    [JsonPropertyName("NAME")]
    public string Name { get; set; } = "Mustermann";

    [JsonPropertyName("VORNAME")]
    public string Vorname { get; set; } = "Max";

    [JsonPropertyName("STAMMNUMMER")]
    public string Stammnummer { get; set; } = "12345";

    [JsonPropertyName("BERUFSGRUPPE")]
    public string Berufsgruppe { get; set; } = "FIA22";

    [JsonPropertyName("AUSBILDER")]
    public string Ausbilder { get; set; } = "Schmidt";

    [JsonPropertyName("ARBEITSBEGINN")]
    public string Arbeitsbeginn { get; set; } = "08:00";

    [JsonPropertyName("ARBEITSENDE")]
    public string Arbeitsende { get; set; } = "16:30";

    [JsonPropertyName("QUICKFILL_START")]
    public string QuickFillStart { get; set; } = "08:00";

    [JsonPropertyName("QUICKFILL_END")]
    public string QuickFillEnd { get; set; } = "16:30";

    [JsonPropertyName("QUICKFILL_DURATION_MINUTES")]
    public int QuickFillDurationMinutes { get; set; } = 480;

    [JsonPropertyName("QUICKFILL_DURATION_HOURS")]
    public double QuickFillDurationHours { get; set; } = 8;

    [JsonPropertyName("QUICKFILL_MONDAY")]
    public bool QuickFillMonday { get; set; } = true;

    [JsonPropertyName("QUICKFILL_TUESDAY")]
    public bool QuickFillTuesday { get; set; } = true;

    [JsonPropertyName("QUICKFILL_WEDNESDAY")]
    public bool QuickFillWednesday { get; set; } = true;

    [JsonPropertyName("QUICKFILL_THURSDAY")]
    public bool QuickFillThursday { get; set; } = true;

    [JsonPropertyName("QUICKFILL_FRIDAY")]
    public bool QuickFillFriday { get; set; } = true;

    [JsonPropertyName("REGION")]
    public string Region { get; set; } = "Baden-Württemberg";

    [JsonPropertyName("MAIL_TO")]
    public string MailTo { get; set; } = "ausbilder@firma.de";

    [JsonPropertyName("MAIL_CC")]
    public string MailCc { get; set; } = "";

    [JsonPropertyName("OUTLOOKSIGNATURNAME")]
    public string OutlookSignaturName { get; set; } = "";

    [JsonPropertyName("MAIL_SUBJECT_TEMPLATE")]
    public string MailSubjectTemplate { get; set; } = "Arbeitszeiten - {vorname} {name} - KW{week}";

    [JsonPropertyName("MAIL_BODY_TEMPLATE")]
    public string MailBodyTemplate { get; set; } = "Guten Tag,\n\nim Anhang befindet sich meine Arbeitszeiterfassung für Kalenderwoche {week}.\n\nMit freundlichen Grüßen,\n{vorname} {name}";

    [JsonPropertyName("FILENAME_TEMPLATE")]
    public string FileNameTemplate { get; set; } = "Arbeitszeiterfassung_{name}{vorname}_KW{week}.xlsx";

    [JsonPropertyName("GITHUB_OWNER")]
    public string GitHubOwner { get; set; } = string.Empty;

    [JsonPropertyName("GITHUB_REPO")]
    public string GitHubRepo { get; set; } = string.Empty;

    [JsonPropertyName("GITHUB_ASSET_NAME")]
    public string GitHubAssetName { get; set; } = string.Empty;

    public static AppConfig Defaults() => new();
}
