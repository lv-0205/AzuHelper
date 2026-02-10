using System.IO;
using System.Reflection;

namespace AzuHelper_v2.Services;

public static class PathService
{
    public static string ResourcePath(string relativePath)
    {
        // Resolve from the app base directory to work both in dev and after build.
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, relativePath));
    }

    public static string SavesDirectory()
    {
        var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var dir = Path.Combine(baseDir, "AzuHelper", "saves");

        Directory.CreateDirectory(dir);
        return dir;
    }

    public static string ConfigPath()
    {
        // Non-packaged behavior: current working directory.
        return Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "config.json"));
    }

    public static string AppIconPath(string relative) => ResourcePath(relative);

    public static string AssemblyDirectory()
        => Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? Environment.CurrentDirectory;
}
