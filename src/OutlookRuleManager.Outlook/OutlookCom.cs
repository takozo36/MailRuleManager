namespace OutlookRuleManager.Outlook;

/// <summary>
/// Outlook のオブジェクトモデルを遅延バインディング（dynamic / IDispatch）で呼ぶための共通部品。
/// 相互運用アセンブリ（Microsoft.Office.Interop.Outlook）を使わないので、外部ライブラリへの依存がない。
/// 早期バインディングと比べた読み込み速度は実測で同等（Outlook 側のコストが支配的なため）。
///
/// 注意: 引数に dynamic を渡すと静的メソッドの呼び出しまで実行時解決になるので、
/// このプロジェクトの補助メソッドへ渡すときは (object) にキャストしてから渡す。
/// </summary>
internal static class OutlookCom
{
    // Outlook の列挙値（相互運用アセンブリを使わないので自前で定義する）
    public const int OlMailItem = 0;          // OlItemType.olMailItem
    public const int OlImportanceLow = 0;     // OlImportance
    public const int OlImportanceHigh = 2;

    /// <summary>
    /// オブジェクトや配列を値に持つプロパティへ代入する（例: MoveToFolder.Folder、Subject.Text）。
    /// dynamic での代入（x.Folder = folder）は「操作は失敗しました」(0x80020009) になるが、
    /// IDispatch の PROPERTYPUT を直接呼ぶと成功する（Outlook クラシック x64 で実測）。
    /// 文字列・真偽値・数値の代入は dynamic のままで問題ない。
    /// </summary>
    public static void SetProperty(object target, string name, object? value) =>
        target.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, target, [value]);

    /// <summary>起動中の Outlook に接続する（起動していなければ起動する）。</summary>
    public static dynamic CreateApplication()
    {
        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook（クラシック）が見つかりません。新しい Outlook（New Outlook）には対応していません。");
        return Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Outlook を起動できませんでした。");
    }
}
