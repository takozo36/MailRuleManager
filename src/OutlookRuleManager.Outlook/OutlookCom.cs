using OutlookRuleManager.Core;

namespace OutlookRuleManager.Outlook;

/// <summary>
/// Shared helpers for calling the Outlook object model with late binding (dynamic / IDispatch).
/// The Outlook interop assembly (Microsoft.Office.Interop.Outlook) is not used, so there is no external dependency.
/// Reading speed was measured to be the same as with early binding (the cost is on Outlook's side).
///
/// Note: passing a dynamic argument turns even a call to a static helper into a runtime-bound call,
/// so cast arguments to (object) when passing them to the helpers in this project.
/// </summary>
internal static class OutlookCom
{
    // Outlook enumeration values (defined here because the interop assembly is not referenced)
    public const int OlMailItem = 0;          // OlItemType.olMailItem
    public const int OlFolderInbox = 6;       // OlDefaultFolders.olFolderInbox
    public const int OlHiddenItems = 1;       // OlTableContents.olHiddenItems

    /// <summary>
    /// Assigns a property whose value is an object or an array (e.g. MoveToFolder.Folder, Subject.Text).
    /// Assigning through dynamic (x.Folder = folder) fails with "The operation failed" (0x80020009), while calling
    /// IDispatch PROPERTYPUT directly succeeds (measured on Outlook Classic x64). Strings, booleans and numbers can be
    /// assigned through dynamic without problems.
    /// </summary>
    public static void SetProperty(object target, string name, object? value) =>
        target.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, target, [value]);

    /// <summary>Connects to the running Outlook (starts it if it is not running).</summary>
    public static dynamic CreateApplication()
    {
        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(Loc.T(
                "Outlook（クラシック）が見つかりません。新しい Outlook（New Outlook）には対応していません。",
                "Outlook Classic was not found. The new Outlook is not supported."));
        return Activator.CreateInstance(type)
            ?? throw new InvalidOperationException(Loc.T("Outlook を起動できませんでした。", "Outlook could not be started."));
    }
}
