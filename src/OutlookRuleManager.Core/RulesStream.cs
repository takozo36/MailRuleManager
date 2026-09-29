using System.Text;

namespace OutlookRuleManager.Core;

/// <summary>ルールのまとめデータを解析できなかった（知らない要素・壊れたデータ）。</summary>
public sealed class RulesStreamFormatException(string message) : Exception(message);

/// <summary>
/// Outlook（クラシック）が受信トレイの隠しメッセージ（IPM.RuleOrganizer）の PR_RW_RULES_STREAM に
/// まとめて保存している、全ルールのバイナリを読む。ルールのエクスポート（.rwz）と同じ形式。
///
/// 形式は Microsoft の非公開仕様で、次のオープンソース（いずれも MIT License）の解析結果をもとに C# で実装した:
///   - asklar/rwzreader  https://github.com/asklar/rwzreader  （Copyright (c) 2021 Alexander Sklar）
///   - hughbe/OutlookRulesReader  https://github.com/hughbe/OutlookRulesReader  （Copyright (c) 2021 Hugh Bellamy）
/// 両ライセンスの全文は THIRD-PARTY-NOTICES.md にある。
///
/// 移動先フォルダーはバイナリの EntryID のまま返す（FolderRef.Path は空）。パスへの変換と
/// 存在確認は Outlook 連携側で行う。要素には長さの情報がないので、知らない要素があると
/// 以降を読めない。その場合は RulesStreamFormatException を投げる（呼び出し側は 1 件ずつの読み込みに切り替える）。
/// </summary>
public static class RulesStream
{
    /// <summary>
    /// 全ルールを読む。戻り値の Index は 1 始まりの実行順。
    /// 「このコンピューターのみ」の条件は、どのコンピューターかをこのデータだけでは判定できないため
    /// NeedsMachineCheck を立てて返す（呼び出し側で Outlook に確認して LocalMachineOnly / OtherMachine を決める）。
    /// </summary>
    public static IReadOnlyList<StreamRule> Parse(byte[] data)
    {
        var r = new Reader(data);
        try
        {
            int count = ReadHeader(r);
            var rules = new List<StreamRule>(count);
            for (int i = 0; i < count; i++)
            {
                if (i > 0) r.U16(); // ルール間の区切り（.rwz では 0、ルールのまとめデータでは別の値。使わない）
                rules.Add(ReadRule(r, i + 1));
            }
            return rules;
        }
        catch (RulesStreamFormatException) { throw; }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or IndexOutOfRangeException or DecoderFallbackException)
        {
            throw new RulesStreamFormatException($"位置 {r.Offset} でデータが途切れています: {ex.Message}");
        }
    }

    private static int ReadHeader(Reader r)
    {
        uint signature = r.U32();
        bool hasSignature = signature is 1310720 or 1200000 or 1100000 or 1000000 or 980413 or 970812 or 0;
        if (!hasSignature) throw new RulesStreamFormatException($"ルールデータの先頭が想定と違います (0x{signature:X8})。");
        bool newer = signature is 1310720 or 1200000 or 1100000 or 1000000;
        if (newer) r.U32(); // flags
        for (int k = 1; k <= 8; k++) r.U32();
        if (newer || signature == 0) r.U32();
        int count = r.U16();
        r.U16();
        return count;
    }

    private static StreamRule ReadRule(Reader r, int index)
    {
        r.U16(); // 署名（使わない）
        string name = r.StringObject();
        bool enabled = r.U32() != 0;
        for (int k = 0; k < 4; k++) r.U32();
        r.U32(); // dataSize
        int elementCount = r.U16();
        ushort separator = r.U16();
        if (separator == 0xFFFF)
        {
            r.U16();
            int len = r.U16();
            r.Bytes(len); // "CRuleElement"
        }

        var rule = new StreamRule { Index = index, Name = name, Enabled = enabled };
        for (int e = 0; e < elementCount; e++)
        {
            if (e > 0) r.U16(); // 要素間の区切り 0x8001
            ReadElement(r, rule);
        }
        return rule;
    }

    private static void ReadElement(Reader r, StreamRule rule)
    {
        uint id = r.U32();
        switch (id)
        {
            case 0x64: r.U32(); r.U32(); r.U32(); return;
            case 0x190:
                r.U32(); r.U32();
                rule.Kind = r.U32() == 0x4 ? RuleKind.Send : RuleKind.Receive;
                return;
        }

        if (ConditionIds.TryGetValue(id, out var cond))
        {
            rule.Conditions.Add(ReadCondition(r, id, cond.Type, cond.Data));
            if (cond.Type == ConditionType.LocalMachineOnly) rule.NeedsMachineCheck = true;
            return;
        }
        if (ExceptionIds.TryGetValue(id, out var exc))
        {
            rule.Exceptions.Add(ReadCondition(r, id, exc.Type, exc.Data));
            return;
        }
        if (ActionIds.TryGetValue(id, out var act))
        {
            rule.Actions.Add(ReadAction(r, act.Type, act.Data));
            return;
        }
        throw new RulesStreamFormatException($"ルール「{rule.Name}」に未知の要素 0x{id:X} があります。");
    }

    private static RuleCondition ReadCondition(Reader r, uint id, ConditionType type, DataKind kind)
    {
        var values = ReadData(r, kind, out _, out _);
        return new RuleCondition(type, values);
    }

    private static RuleAction ReadAction(Reader r, ActionType type, DataKind kind)
    {
        var values = ReadData(r, kind, out var folderEid, out var storeEid);
        if (kind == DataKind.Folder)
            return new RuleAction(type, Array.Empty<RuleValue>(),
                folderEid is null ? null : new FolderRef("", folderEid, storeEid ?? ""), folderEid is null);
        return new RuleAction(type, values);
    }

    /// <summary>要素の中身を読む。移動・コピー先は EntryID（16 進文字列）を返す。</summary>
    private static IReadOnlyList<RuleValue> ReadData(Reader r, DataKind kind, out string? folderEid, out string? storeEid)
    {
        folderEid = storeEid = null;
        switch (kind)
        {
            case DataKind.Simple:
                r.U32();
                return Array.Empty<RuleValue>();

            case DataKind.People:
            {
                r.U32(); r.U32();
                uint n = r.U32();
                var list = new List<RuleValue>((int)n);
                for (int i = 0; i < n; i++) list.Add(ReadRecipient(r));
                r.U32(); r.U32();
                return list;
            }

            case DataKind.Strings:
            {
                uint n = r.U32();
                var list = new List<RuleValue>((int)n);
                for (int i = 0; i < n; i++)
                {
                    r.U32();
                    string s = r.StringObject();
                    if (s.Length > 0) list.Add(RuleValue.Text(s));
                }
                return list;
            }

            case DataKind.OneString:
            {
                r.U32(); r.U32();
                string s = r.StringObject();
                return s.Length == 0 ? Array.Empty<RuleValue>() : [RuleValue.Text(s)];
            }

            case DataKind.Categories:
            {
                r.U32(); r.U32();
                return r.StringObject().Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Select(RuleValue.Text).ToList();
            }

            case DataKind.Importance:
            {
                r.U32(); r.U32();
                return [RuleValue.Text(r.U32() switch { 0 => "低", 2 => "高", _ => "標準" })];
            }

            case DataKind.OneNumber:
                r.U32(); r.U32(); r.U32();
                return Array.Empty<RuleValue>();

            case DataKind.Folder:
            {
                r.U32(); r.U32();
                folderEid = SizedHex(r);
                storeEid = SizedHex(r);
                r.StringObject(); // フォルダー名（パスは Outlook に問い合わせる）
                r.U32();
                return Array.Empty<RuleValue>();
            }

            case DataKind.Machine:
                r.U32(); r.U32(); r.Bytes(16);
                return Array.Empty<RuleValue>();

            case DataKind.Flag:
                r.U32(); r.U32(); r.U32(); r.StringObject(); r.U32();
                return Array.Empty<RuleValue>();

            case DataKind.NumberAndString:
                r.U32(); r.U32(); r.U32(); r.StringObject();
                return Array.Empty<RuleValue>();

            case DataKind.CustomAction:
                r.U32(); r.U32(); r.StringObject(); r.StringObject(); r.StringObject(); r.StringObject();
                return Array.Empty<RuleValue>();

            case DataKind.ServerReply:
            {
                r.U32(); r.U32();
                uint size = r.U32();
                r.Bytes((int)size);
                r.StringObject();
                return Array.Empty<RuleValue>();
            }

            case DataKind.Script:
                r.U32(); r.U32(); r.StringObject(); r.StringObject();
                return Array.Empty<RuleValue>();

            case DataKind.Retention:
                r.U32(); r.U32(); r.U32(); r.Bytes(16); r.StringObject();
                return Array.Empty<RuleValue>();

            case DataKind.SizeRange:
                r.U32(); r.U32(); r.U32(); r.U32();
                return Array.Empty<RuleValue>();

            case DataKind.DateRange:
                r.U32(); r.U32();
                r.U32(); r.U32(); r.U64();
                r.U32(); r.U32(); r.U64();
                return Array.Empty<RuleValue>();

            case DataKind.FormType:
            {
                uint n = r.U32(); r.U32();
                var list = new List<RuleValue>();
                for (int i = 0; i < n; i++)
                {
                    list.Add(RuleValue.Text(r.StringObject()));
                    int len = r.U8();
                    r.Bytes(len);
                    while (r.Remaining > 0 && r.PeekU8() == 0) r.U8(); // 詰め物
                }
                return list;
            }

            case DataKind.DocumentProperties:
            {
                r.U32(); r.U32();
                r.StringObject();
                int props = r.U16();
                for (int i = 0; i < props; i++)
                {
                    r.StringObject(); r.U32(); r.U32(); r.StringObject();
                    r.U32(); r.U32(); r.U32(); r.U32(); r.U32(); r.U32();
                    r.U32(); r.U32(); r.U64(); r.U32();
                }
                uint classes = r.U32();
                for (int i = 0; i < classes; i++) r.Bytes(r.U8());
                return Array.Empty<RuleValue>();
            }

            default:
                throw new RulesStreamFormatException($"内部エラー: 未対応のデータ種別 {kind}");
        }
    }

    private static string? SizedHex(Reader r)
    {
        uint size = r.U32();
        if (size == 0) return null;
        return Convert.ToHexString(r.Bytes((int)size));
    }

    /// <summary>
    /// 宛先 1 件（MAPI のプロパティ値の並び）。表示名・メールアドレスを取り出す。
    /// EntryID を持たない宛先は「アドレス帳で解決できていない」とみなす。
    /// </summary>
    private static RuleValue ReadRecipient(Reader r)
    {
        r.U32();
        uint props = r.U32();
        uint dataSize = r.U32();
        int start = r.Offset;
        int end = checked(start + (int)dataSize);
        string? name = null, email = null, smtp = null;
        bool hasEntryId = false;
        for (int i = 0; i < props; i++)
        {
            ushort type = r.U16();
            ushort pid = r.U16();
            r.U32();
            uint d1 = r.U32();
            uint d2 = r.U32();
            switch (type)
            {
                case 0x1F: // PtypString（UTF-16、NUL 終端）
                {
                    string s = r.Utf16ZAt(start + (int)d1, end);
                    if (pid == 0x3001) name = s;
                    else if (pid == 0x3003) email = s;
                    else if (pid == 0x39FE) smtp = s;
                    break;
                }
                case 0x1E: // PtypString8
                {
                    string s = r.AsciiZAt(start + (int)d1, end);
                    if (pid == 0x3001) name ??= s;
                    else if (pid == 0x3003) email ??= s;
                    else if (pid == 0x39FE) smtp ??= s;
                    break;
                }
                case 0x102: // PtypBinary
                    if (pid == 0x0FFF && d2 > 0) hasEntryId = true;
                    break;
            }
        }
        r.Offset = end;
        return RuleValue.Address(name, email ?? smtp, hasEntryId);
    }

    // ---- 要素 ID と、Outlook のオブジェクトモデルでの種類の対応 ----

    private enum DataKind
    {
        Simple, People, Strings, OneString, Categories, Importance, OneNumber, Folder, Machine,
        Flag, NumberAndString, CustomAction, ServerReply, Script, Retention, SizeRange, DateRange,
        FormType, DocumentProperties,
    }

    private static readonly Dictionary<uint, (ConditionType Type, DataKind Data)> ConditionIds = new()
    {
        [0xC8] = (ConditionType.To, DataKind.Simple),
        [0xC9] = (ConditionType.OnlyToMe, DataKind.Simple),
        [0xCA] = (ConditionType.NotTo, DataKind.Simple),
        [0xCB] = (ConditionType.From, DataKind.People),
        [0xCC] = (ConditionType.SentTo, DataKind.People),
        [0xCD] = (ConditionType.Subject, DataKind.Strings),
        [0xCE] = (ConditionType.Body, DataKind.Strings),
        [0xCF] = (ConditionType.BodyOrSubject, DataKind.Strings),
        [0xD0] = (ConditionType.FlaggedForAction, DataKind.OneString),
        [0xD2] = (ConditionType.Importance, DataKind.Importance),
        [0xD3] = (ConditionType.Sensitivity, DataKind.OneNumber),
        [0xD7] = (ConditionType.Category, DataKind.Categories),
        [0xDC] = (ConditionType.OutOfOffice, DataKind.Simple),
        [0xDE] = (ConditionType.HasAttachment, DataKind.Simple),
        [0xDF] = (ConditionType.Property, DataKind.DocumentProperties),
        [0xE0] = (ConditionType.SizeRange, DataKind.SizeRange),
        [0xE1] = (ConditionType.DateRange, DataKind.DateRange),
        [0xE2] = (ConditionType.Cc, DataKind.Simple),
        [0xE3] = (ConditionType.ToOrCc, DataKind.Simple),
        [0xE4] = (ConditionType.FormName, DataKind.FormType),
        [0xE5] = (ConditionType.RecipientAddress, DataKind.Strings),
        [0xE6] = (ConditionType.SenderAddress, DataKind.Strings),
        [0xE8] = (ConditionType.MessageHeader, DataKind.Strings),
        [0xE9] = (ConditionType.Unknown, DataKind.Simple),   // 例外リストの差出人
        [0xEB] = (ConditionType.Unknown, DataKind.Simple),   // 迷惑メールの疑い
        [0xEC] = (ConditionType.Unknown, DataKind.Simple),   // 成人向けコンテンツ
        [0xED] = (ConditionType.Unknown, DataKind.SizeRange), // 関連度
        [0xEE] = (ConditionType.Account, DataKind.OneString),
        [0xEF] = (ConditionType.LocalMachineOnly, DataKind.Machine),
        [0xF0] = (ConditionType.SenderInAddressBook, DataKind.OneString),
        [0xF1] = (ConditionType.MeetingInviteOrUpdate, DataKind.Simple),
        [0xF2] = (ConditionType.Unknown, DataKind.Simple),   // 連絡先から
        [0xF3] = (ConditionType.Unknown, DataKind.Simple),   // 購読から
        [0xF4] = (ConditionType.FormName, DataKind.FormType),
        [0xF5] = (ConditionType.FromRssFeed, DataKind.Strings),
        [0xF6] = (ConditionType.AnyCategory, DataKind.Simple),
        [0xF7] = (ConditionType.FromAnyRssFeed, DataKind.Simple),
    };

    private static readonly Dictionary<uint, (ConditionType Type, DataKind Data)> ExceptionIds = new()
    {
        [0x1F4] = (ConditionType.To, DataKind.Simple),
        [0x1F5] = (ConditionType.OnlyToMe, DataKind.Simple),
        [0x1F6] = (ConditionType.NotTo, DataKind.Simple),
        [0x1F7] = (ConditionType.From, DataKind.People),
        [0x1F8] = (ConditionType.SentTo, DataKind.People),
        [0x1F9] = (ConditionType.Subject, DataKind.Strings),
        [0x1FA] = (ConditionType.Body, DataKind.Strings),
        [0x1FB] = (ConditionType.BodyOrSubject, DataKind.Strings),
        [0x1FC] = (ConditionType.FlaggedForAction, DataKind.OneString),
        [0x1FE] = (ConditionType.Importance, DataKind.Importance),
        [0x1FF] = (ConditionType.Sensitivity, DataKind.OneNumber),
        [0x203] = (ConditionType.Category, DataKind.Categories),
        [0x208] = (ConditionType.OutOfOffice, DataKind.Simple),
        [0x20A] = (ConditionType.HasAttachment, DataKind.Simple),
        [0x20B] = (ConditionType.Property, DataKind.DocumentProperties),
        [0x20C] = (ConditionType.SizeRange, DataKind.SizeRange),
        [0x20D] = (ConditionType.DateRange, DataKind.DateRange),
        [0x20E] = (ConditionType.Cc, DataKind.Simple),
        [0x20F] = (ConditionType.ToOrCc, DataKind.Simple),
        [0x210] = (ConditionType.FormName, DataKind.FormType),
        [0x211] = (ConditionType.RecipientAddress, DataKind.Strings),
        [0x212] = (ConditionType.SenderAddress, DataKind.Strings),
        [0x213] = (ConditionType.MessageHeader, DataKind.Strings),
        [0x214] = (ConditionType.Account, DataKind.OneString),
        [0x215] = (ConditionType.SenderInAddressBook, DataKind.OneString),
        [0x216] = (ConditionType.MeetingInviteOrUpdate, DataKind.Simple),
        [0x217] = (ConditionType.Unknown, DataKind.Simple),  // 連絡先から
        [0x218] = (ConditionType.FormName, DataKind.FormType),
        [0x219] = (ConditionType.FromRssFeed, DataKind.Strings),
        [0x21A] = (ConditionType.AnyCategory, DataKind.Simple),
        [0x21B] = (ConditionType.FromAnyRssFeed, DataKind.Simple),
    };

    private static readonly Dictionary<uint, (ActionType Type, DataKind Data)> ActionIds = new()
    {
        [0x12C] = (ActionType.MoveToFolder, DataKind.Folder),
        [0x12D] = (ActionType.Delete, DataKind.Simple),
        [0x12E] = (ActionType.Forward, DataKind.People),
        [0x12F] = (ActionType.Template, DataKind.OneString),
        [0x130] = (ActionType.NewItemAlert, DataKind.OneString),
        [0x131] = (ActionType.FlagForActionInDays, DataKind.Flag),
        [0x132] = (ActionType.FlagClear, DataKind.Simple),
        [0x133] = (ActionType.AssignToCategory, DataKind.Categories),
        [0x136] = (ActionType.PlaySound, DataKind.OneString),
        [0x137] = (ActionType.Importance, DataKind.Importance),
        [0x138] = (ActionType.Sensitivity, DataKind.OneNumber),
        [0x139] = (ActionType.CopyToFolder, DataKind.Folder),
        [0x13A] = (ActionType.NotifyRead, DataKind.Simple),
        [0x13B] = (ActionType.NotifyDelivery, DataKind.Simple),
        [0x13C] = (ActionType.CcMessage, DataKind.People),
        [0x13E] = (ActionType.Defer, DataKind.OneNumber),
        [0x13F] = (ActionType.CustomAction, DataKind.CustomAction),
        [0x142] = (ActionType.Stop, DataKind.Simple),
        [0x143] = (ActionType.Unknown, DataKind.Simple),     // 商用・成人向けの検査をしない
        [0x144] = (ActionType.Redirect, DataKind.People),
        [0x145] = (ActionType.Unknown, DataKind.Simple),     // 関連度を加算
        [0x146] = (ActionType.ServerReply, DataKind.ServerReply),
        [0x147] = (ActionType.ForwardAsAttachment, DataKind.People),
        [0x148] = (ActionType.Print, DataKind.Simple),
        [0x149] = (ActionType.StartApplication, DataKind.OneString),
        [0x14A] = (ActionType.DeletePermanently, DataKind.Simple),
        [0x14B] = (ActionType.RunScript, DataKind.Script),
        [0x14C] = (ActionType.MarkRead, DataKind.Simple),
        [0x14F] = (ActionType.DesktopAlert, DataKind.Simple),
        [0x150] = (ActionType.FlagColor, DataKind.Simple),
        [0x151] = (ActionType.MarkAsTask, DataKind.NumberAndString),
        [0x152] = (ActionType.ClearCategories, DataKind.Simple),
        [0x153] = (ActionType.Unknown, DataKind.Retention),  // 保持ポリシー
    };

    /// <summary>リトルエンディアンの読み取り位置付きリーダー。</summary>
    private sealed class Reader(byte[] data)
    {
        public int Offset { get; set; }
        public int Remaining => data.Length - Offset;

        public byte U8() => data[Offset++];
        public byte PeekU8() => data[Offset];
        public ushort U16() { var v = BitConverter.ToUInt16(Take(2)); return v; }
        public uint U32() { var v = BitConverter.ToUInt32(Take(4)); return v; }
        public ulong U64() { var v = BitConverter.ToUInt64(Take(8)); return v; }
        public byte[] Bytes(int n) => Take(n).ToArray();

        private ReadOnlySpan<byte> Take(int n)
        {
            if (n < 0 || Offset + n > data.Length)
                throw new RulesStreamFormatException($"位置 {Offset} から {n} バイトを読めません（全体 {data.Length} バイト）。");
            var span = data.AsSpan(Offset, n);
            Offset += n;
            return span;
        }

        /// <summary>長さ付きの UTF-16 文字列（長さ 1 バイト。0xFF のときは続く 2 バイトが長さで、その後 2 バイトを読み飛ばす）。</summary>
        public string StringObject()
        {
            int len = U8();
            if (len == 0xFF)
            {
                len = U16();
                Take(2);
            }
            return Encoding.Unicode.GetString(Take(len * 2));
        }

        public string Utf16ZAt(int position, int limit)
        {
            int p = position;
            while (p + 1 < limit && (data[p] != 0 || data[p + 1] != 0)) p += 2;
            return Encoding.Unicode.GetString(data, position, p - position);
        }

        public string AsciiZAt(int position, int limit)
        {
            int p = position;
            while (p < limit && data[p] != 0) p++;
            return Encoding.ASCII.GetString(data, position, p - position);
        }
    }
}

/// <summary>ルールのまとめデータから読んだルール 1 件（移動先は EntryID のまま）。</summary>
public sealed class StreamRule
{
    public required int Index { get; init; }
    public required string Name { get; init; }
    public bool Enabled { get; init; }
    public RuleKind Kind { get; set; } = RuleKind.Receive;
    public List<RuleCondition> Conditions { get; } = new();
    public List<RuleCondition> Exceptions { get; } = new();
    public List<RuleAction> Actions { get; } = new();
    /// <summary>「このコンピューターのみ」を含む（このPCか別のPCかは Outlook に確認が必要）。</summary>
    public bool NeedsMachineCheck { get; set; }

    /// <summary>
    /// RuleData にする。resolveFolder で移動・コピー先の EntryID をフォルダーに変換する（見つからなければ null）。
    /// onThisMachine は「このコンピューターのみ」が このPC なら true、別のPCなら false。
    /// isLocalRule は Outlook の Rule.IsLocalRule（サウンド・新着通知などクライアントでしか動かない処理を含むと true。
    /// まとめデータからは判定しきれないので Outlook から受け取る）。
    /// 条件・例外・処理は、Outlook のオブジェクトモデルが返す順（種類ごとに固定）に並べ直す。
    /// </summary>
    public RuleData ToRuleData(Func<FolderRef, FolderRef?> resolveFolder, bool onThisMachine = true, bool isLocalRule = false)
    {
        var conditions = Conditions
            .Select(c => c.Type == ConditionType.LocalMachineOnly && !onThisMachine ? c with { Type = ConditionType.OtherMachine } : c)
            .OrderBy(c => Rank(ConditionOrder, (int)c.Type))
            .ToList();
        var actions = Actions
            .Select(a => a.Folder is { } f ? RuleAction.ToFolder(a.Type, resolveFolder(f)) : a)
            .OrderBy(a => Rank(ActionOrder, (int)a.Type))
            .ToList();
        return new RuleData
        {
            Index = Index,
            Name = Name,
            Enabled = Enabled,
            Kind = Kind,
            IsLocalRule = isLocalRule,
            Conditions = conditions,
            Exceptions = Exceptions.OrderBy(c => Rank(ConditionOrder, (int)c.Type)).ToList(),
            Actions = actions,
        };
    }

    // Outlook の RuleConditions / RuleActions を列挙したときの種類の順（Outlook クラシック x64 16.0 で実測）
    private static readonly int[] ConditionOrder = [4, 26, 9, 20, 11, 5, 10, 6, 3, 18, 13, 14, 15, 16, 17, 2, 23, 1, 12, 25, 27, 7, 8, 19, 21, 22, 24, 28, 29, 31, 30];
    private static readonly int[] ActionOrder = [1, 5, 4, 3, 24, 21, 25, 26, 27, 6, 7, 8, 2, 17, 29, 23, 9, 10, 11, 12, 13, 14, 15, 16, 19, 28, 18, 30];

    private static int Rank(int[] order, int type)
    {
        int i = Array.IndexOf(order, type);
        return i < 0 ? order.Length + type : i;
    }
}
