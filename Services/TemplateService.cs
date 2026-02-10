using System.Collections.Generic;

namespace AzuHelper_v2.Services;

public static class TemplateService
{
    public static string Apply(string template, IDictionary<string, string> tokens)
    {
        if (string.IsNullOrWhiteSpace(template))
        {
            return string.Empty;
        }

        var result = template;
        foreach (var (key, value) in tokens)
        {
            var token = "{" + key + "}";
            result = result.Replace(token, value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        return result;
    }
}
