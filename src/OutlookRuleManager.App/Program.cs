using OutlookRuleManager.Core;

namespace OutlookRuleManager.App;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        AppSettings.Load();
        Loc.Current = AppSettings.ResolveLanguage();
        Application.Run(new MainForm());
    }
}
