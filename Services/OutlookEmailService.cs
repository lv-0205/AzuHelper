using System.Runtime.InteropServices;
using System.IO;

namespace AzuHelper_v2.Services;

public static class OutlookEmailService
{
    public static bool SendEmail(
        string to,
        string cc,
        string subject,
        string body,
        string? attachmentPath,
        bool openOnly,
        string? signatureName,
        out string? errorMessage)
    {
        errorMessage = null;

        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType is null)
            {
                errorMessage = "Cannot start Outlook. The new Outlook app doesn't support COM. Please use classic Outlook.";
                return false;
            }

            dynamic? outlook = null;
            dynamic? mail = null;

            try
            {
                outlook = Activator.CreateInstance(outlookType);
                mail = outlook?.CreateItem(0);
                if (mail is null)
                {
                    errorMessage = "Failed to create Outlook mail item.";
                    return false;
                }

                mail.To = to;
                mail.CC = cc;
                mail.Subject = subject;

                var signature = LoadSignature(signatureName);
                mail.Body = string.IsNullOrWhiteSpace(signature)
                    ? body
                    : body + Environment.NewLine + Environment.NewLine + signature;

                if (!string.IsNullOrWhiteSpace(attachmentPath) && File.Exists(attachmentPath))
                {
                    mail.Attachments.Add(attachmentPath);
                }

                if (openOnly)
                {
                    mail.Display();
                }
                else
                {
                    mail.Send();
                }

                return true;
            }
            finally
            {
                if (mail is not null)
                {
                    Marshal.FinalReleaseComObject(mail);
                }

                if (outlook is not null)
                {
                    Marshal.FinalReleaseComObject(outlook);
                }
            }
        }
        catch (Exception ex)
        {
            errorMessage = $"Error sending email: {ex.Message}";
            return false;
        }
    }

    private static string? LoadSignature(string? signatureName)
    {
        if (string.IsNullOrWhiteSpace(signatureName))
        {
            return null;
        }

        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var signaturePath = Path.Combine(appData, "Microsoft", "Signatures", signatureName + ".txt");
        if (!File.Exists(signaturePath))
        {
            return null;
        }

        try
        {
            return File.ReadAllText(signaturePath);
        }
        catch
        {
            return null;
        }
    }
}
