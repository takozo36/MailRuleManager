using System.Text;
using static OutlookRuleManager.Core.Loc;

namespace OutlookRuleManager.Core;

/// <summary>The rules stream could not be parsed (unknown element or broken data).</summary>
public sealed class RulesStreamFormatException(string message) : Exception(message);

/// <summary>
/// Reads the binary that classic Outlook stores for all rules in PR_RW_RULES_STREAM of the hidden message
/// (message class IPM.RuleOrganizer) in the Inbox. It has the same format as an exported .rwz file.
///
/// The format is not documented by Microsoft. This C# implementation is based on the format analysis of the
/// following open-source projects (both MIT License):
///   - asklar/rwzreader  https://github.com/asklar/rwzreader  (Copyright (c) 2021 Alexander Sklar)
///   - hughbe/OutlookRulesReader  https://github.com/hughbe/OutlookRulesReader  (Copyright (c) 2021 Hugh Bellamy)
/// The full license texts are in THIRD-PARTY-NOTICES.md.
///
/// Move-to folders are returned as raw EntryIDs (FolderRef.Path is empty); resolving them to paths and checking
/// that they exist is done by the Outlook layer. Elements carry no length, so an unknown element makes the rest
/// unreadable; in that case RulesStreamFormatException is thrown (the caller falls back to reading rule by rule).
/// </summary>
public static class RulesStream
{
    /// <summary>
    /// Parses all rules. Index of each result is the 1-based execution order.
    /// For "on this computer only", this data alone cannot tell which computer is meant, so NeedsMachineCheck is set
    /// (the caller asks Outlook and decides between LocalMachineOnly and OtherMachine).
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
                if (i > 0) r.U16(); // separator between rules (0 in .rwz, another value in the rules stream; unused)
                rules.Add(ReadRule(r, i + 1));
            }
            return rules;
        }
        catch (RulesStreamFormatException) { throw; }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or IndexOutOfRangeException or DecoderFallbackException)
        {
            throw new RulesStreamFormatException(T($"位置 {r.Offset} でデータが途切れています: {ex.Message}", $"Data ends unexpectedly at offset {r.Offset}: {ex.Message}"));
        }
    }

    private static int ReadHeader(Reader r)
    {
        uint signature = r.U32();
        bool hasSignature = signature is 1310720 or 1200000 or 1100000 or 1000000 or 980413 or 970812 or 0;
        if (!hasSignature)
            throw new RulesStreamFormatException(T($"ルールデータの先頭が想定と違います (0x{signature:X8})。", $"Unexpected rules data signature (0x{signature:X8})."));
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
        r.U16(); // signature (unused)
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
            if (e > 0) r.U16(); // separator between elements (0x8001)
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
            rule.Conditions.Add(new RuleCondition(cond.Type, ReadData(r, cond.Data, out _, out _)));
            if (cond.Type == ConditionType.LocalMachineOnly) rule.NeedsMachineCheck = true;
            return;
        }
        if (ExceptionIds.TryGetValue(id, out var exc))
        {
            rule.Exceptions.Add(new RuleCondition(exc.Type, ReadData(r, exc.Data, out _, out _)));
            return;
        }
        if (ActionIds.TryGetValue(id, out var act))
        {
            rule.Actions.Add(ReadAction(r, act.Type, act.Data));
            return;
        }
        throw new RulesStreamFormatException(T($"ルール「{rule.Name}」に未知の要素 0x{id:X} があります。", $"Rule \"{rule.Name}\" contains an unknown element 0x{id:X}."));
    }

    private static RuleAction ReadAction(Reader r, ActionType type, DataKind kind)
    {
        var values = ReadData(r, kind, out var folderEid, out var storeEid);
        if (kind == DataKind.Folder)
            return new RuleAction(type, Array.Empty<RuleValue>(),
                folderEid is null ? null : new FolderRef("", folderEid, storeEid ?? ""), folderEid is null);
        return new RuleAction(type, values);
    }

    /// <summary>Reads the payload of an element. For move / copy, returns the EntryIDs as hex strings.</summary>
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
                return [RuleValue.Importance((int)r.U32())];
            }

            case DataKind.OneNumber:
                r.U32(); r.U32(); r.U32();
                return Array.Empty<RuleValue>();

            case DataKind.Folder:
            {
                r.U32(); r.U32();
                folderEid = SizedHex(r);
                storeEid = SizedHex(r);
                r.StringObject(); // folder name (the full path is resolved through Outlook)
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
                    while (r.Remaining > 0 && r.PeekU8() == 0) r.U8(); // padding
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
                throw new RulesStreamFormatException($"Internal error: unsupported data kind {kind}");
        }
    }

    private static string? SizedHex(Reader r)
    {
        uint size = r.U32();
        if (size == 0) return null;
        return Convert.ToHexString(r.Bytes((int)size));
    }

    /// <summary>
    /// One recipient (an array of MAPI property values). Extracts the display name and the e-mail address.
    /// A recipient without an EntryID is treated as not resolved in the address book.
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
                case 0x1F: // PtypString (UTF-16, NUL-terminated)
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

    // ---- Element IDs and the corresponding types in the Outlook object model ----

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
        [0xE9] = (ConditionType.Unknown, DataKind.Simple),   // from senders on my exception list
        [0xEB] = (ConditionType.Unknown, DataKind.Simple),   // suspected junk e-mail
        [0xEC] = (ConditionType.Unknown, DataKind.Simple),   // adult content
        [0xED] = (ConditionType.Unknown, DataKind.SizeRange), // relevance range
        [0xEE] = (ConditionType.Account, DataKind.OneString),
        [0xEF] = (ConditionType.LocalMachineOnly, DataKind.Machine),
        [0xF0] = (ConditionType.SenderInAddressBook, DataKind.OneString),
        [0xF1] = (ConditionType.MeetingInviteOrUpdate, DataKind.Simple),
        [0xF2] = (ConditionType.Unknown, DataKind.Simple),   // from my contacts
        [0xF3] = (ConditionType.Unknown, DataKind.Simple),   // from a subscription
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
        [0x217] = (ConditionType.Unknown, DataKind.Simple),  // from my contacts
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
        [0x143] = (ActionType.Unknown, DataKind.Simple),     // do not search for commercial or adult content
        [0x144] = (ActionType.Redirect, DataKind.People),
        [0x145] = (ActionType.Unknown, DataKind.Simple),     // add to relevance
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
        [0x153] = (ActionType.Unknown, DataKind.Retention),  // apply retention policy
    };

    /// <summary>Little-endian reader with a position.</summary>
    private sealed class Reader(byte[] data)
    {
        public int Offset { get; set; }
        public int Remaining => data.Length - Offset;

        public byte U8() => data[Offset++];
        public byte PeekU8() => data[Offset];
        public ushort U16() => BitConverter.ToUInt16(Take(2));
        public uint U32() => BitConverter.ToUInt32(Take(4));
        public ulong U64() => BitConverter.ToUInt64(Take(8));
        public byte[] Bytes(int n) => Take(n).ToArray();

        private ReadOnlySpan<byte> Take(int n)
        {
            if (n < 0 || Offset + n > data.Length)
                throw new RulesStreamFormatException(T(
                    $"位置 {Offset} から {n} バイトを読めません（全体 {data.Length} バイト）。",
                    $"Cannot read {n} bytes at offset {Offset} (total {data.Length} bytes)."));
            var span = data.AsSpan(Offset, n);
            Offset += n;
            return span;
        }

        /// <summary>Length-prefixed UTF-16 string (1-byte length; if it is 0xFF, the next 2 bytes are the length and 2 more bytes are skipped).</summary>
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

/// <summary>A rule read from the rules stream (move-to folders are still raw EntryIDs).</summary>
public sealed class StreamRule
{
    public required int Index { get; init; }
    public required string Name { get; init; }
    public bool Enabled { get; init; }
    public RuleKind Kind { get; set; } = RuleKind.Receive;
    public List<RuleCondition> Conditions { get; } = new();
    public List<RuleCondition> Exceptions { get; } = new();
    public List<RuleAction> Actions { get; } = new();
    /// <summary>Contains "on this computer only" (whether it means this computer must be asked from Outlook).</summary>
    public bool NeedsMachineCheck { get; set; }

    /// <summary>
    /// Converts to RuleData. resolveFolder turns the EntryID of a move / copy destination into a folder (null if not found).
    /// onThisMachine is true when "on this computer only" refers to this computer, false for another computer.
    /// isLocalRule is Outlook's Rule.IsLocalRule (true when the rule contains client-only actions such as a sound or a
    /// new item alert; it cannot be fully derived from the stream, so it is taken from Outlook).
    /// Conditions, exceptions and actions are sorted into the order the Outlook object model returns them (fixed per type).
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

    // Order of types when enumerating Outlook's RuleConditions / RuleActions (measured on classic Outlook x64 16.0)
    private static readonly int[] ConditionOrder = [4, 26, 9, 20, 11, 5, 10, 6, 3, 18, 13, 14, 15, 16, 17, 2, 23, 1, 12, 25, 27, 7, 8, 19, 21, 22, 24, 28, 29, 31, 30];
    private static readonly int[] ActionOrder = [1, 5, 4, 3, 24, 21, 25, 26, 27, 6, 7, 8, 2, 17, 29, 23, 9, 10, 11, 12, 13, 14, 15, 16, 19, 28, 18, 30];

    private static int Rank(int[] order, int type)
    {
        int i = Array.IndexOf(order, type);
        return i < 0 ? order.Length + type : i;
    }
}
