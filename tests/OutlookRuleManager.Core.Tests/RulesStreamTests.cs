using System.Text;
using OutlookRuleManager.Core;

namespace OutlookRuleManager.Core.Tests;

public class RulesStreamTests : JapaneseTestBase
{
    private static readonly byte[] FolderA = [0x00, 0x00, 0x00, 0x00, 0xAA, 0xBB, 0x01];
    private static readonly byte[] FolderB = [0x00, 0x00, 0x00, 0x00, 0xCC, 0xDD, 0x02];
    private static readonly byte[] Store = [0x01, 0x02, 0x03];

    /// <summary>Two typical rules (sender → move → stop; subject + this computer only + sound).</summary>
    private static byte[] Sample() => new StreamBuilder()
        .Rule("社内　山田", enabled: true, b => b
            .From(("山田 太郎", "yamada@example.com", true))
            .MoveTo(FolderA)
            .Stop())
        .Rule("通知", enabled: false, b => b
            .Subject("請求書", "ご請求")
            .ThisComputerOnly()
            .Sound(@"C:\Windows\Media\notify.wav")
            .Alert("請求書が届きました")
            .MoveTo(FolderB)
            .Stop())
        .Build();

    [Fact]
    public void ReadsNameEnabledConditionsAndActions()
    {
        var rules = RulesStream.Parse(Sample());
        Assert.Equal(2, rules.Count);

        var r1 = rules[0];
        Assert.Equal((1, "社内　山田", true), (r1.Index, r1.Name, r1.Enabled));
        var from = Assert.Single(r1.Conditions);
        Assert.Equal(ConditionType.From, from.Type);
        Assert.Equal("山田 太郎 <yamada@example.com>", from.Values[0].Display);
        Assert.True(from.Values[0].Resolved);
        Assert.Equal([ActionType.MoveToFolder, ActionType.Stop], r1.Actions.Select(a => a.Type));
        Assert.Equal(Convert.ToHexString(FolderA), r1.Actions[0].Folder!.EntryId);
        Assert.Equal(Convert.ToHexString(Store), r1.Actions[0].Folder!.StoreId);
        Assert.False(r1.NeedsMachineCheck);

        var r2 = rules[1];
        Assert.False(r2.Enabled);
        Assert.True(r2.NeedsMachineCheck);
        Assert.Equal(["請求書", "ご請求"], r2.Conditions.First(c => c.Type == ConditionType.Subject).Values.Select(v => v.Display));
        Assert.Equal(@"C:\Windows\Media\notify.wav", r2.Actions.First(a => a.Type == ActionType.PlaySound).Values[0].Display);
        Assert.Equal("請求書が届きました", r2.Actions.First(a => a.Type == ActionType.NewItemAlert).Values[0].Display);
    }

    [Fact]
    public void ToRuleData_ResolvesFoldersAndSortsLikeOutlook()
    {
        var r2 = RulesStream.Parse(Sample())[1];
        var data = r2.ToRuleData(
            f => f.EntryId == Convert.ToHexString(FolderB) ? new FolderRef(@"\\me\受信トレイ\通知", "B", "S") : null,
            onThisMachine: true, isLocalRule: true);

        // Outlook returns subject → this computer only, and move → stop → sound → new item alert
        Assert.Equal([ConditionType.Subject, ConditionType.LocalMachineOnly], data.Conditions.Select(c => c.Type));
        Assert.Equal([ActionType.MoveToFolder, ActionType.Stop, ActionType.PlaySound, ActionType.NewItemAlert], data.Actions.Select(a => a.Type));
        Assert.Equal(@"受信トレイ\通知", data.Actions[0].Folder!.DisplayPath);
        Assert.True(data.IsLocalRule);
        Assert.Equal(2, data.Index);
    }

    [Fact]
    public void MissingFolderMakesTheRuleAnError()
    {
        var data = RulesStream.Parse(Sample())[0].ToRuleData(_ => null);
        Assert.True(data.Actions[0].FolderMissing);
        Assert.Contains(RuleDiagnostics.Analyze([RuleEntry.FromSource(data)])["R1"], d => d.Severity == Severity.Error);
    }

    [Fact]
    public void OtherComputerOnly_BecomesOtherMachine()
    {
        var data = RulesStream.Parse(Sample())[1].ToRuleData(_ => null, onThisMachine: false);
        Assert.Contains(data.Conditions, c => c.Type == ConditionType.OtherMachine);
        Assert.DoesNotContain(data.Conditions, c => c.Type == ConditionType.LocalMachineOnly);
    }

    [Fact]
    public void RecipientWithoutEntryId_IsUnresolved()
    {
        var bytes = new StreamBuilder()
            .Rule("未解決", true, b => b.From(("誰か", "someone@example.com", false)).MoveTo(FolderA))
            .Build();
        Assert.False(RulesStream.Parse(bytes)[0].Conditions[0].Values[0].Resolved);
    }

    [Fact]
    public void ReadsLongNames()
    {
        string longName = new string('あ', 300);
        var bytes = new StreamBuilder().Rule(longName, true, b => b.Stop()).Build();
        Assert.Equal(longName, RulesStream.Parse(bytes)[0].Name);
    }

    [Fact]
    public void UnknownElementThrows()
    {
        var bytes = new StreamBuilder().Rule("謎", true, b => b.Raw(0x999, [0, 0, 0, 0])).Build();
        var ex = Assert.Throws<RulesStreamFormatException>(() => RulesStream.Parse(bytes));
        Assert.Contains("0x999", ex.Message);
    }

    [Fact]
    public void TruncatedDataThrows()
    {
        var bytes = Sample();
        Assert.Throws<RulesStreamFormatException>(() => RulesStream.Parse(bytes[..(bytes.Length / 2)]));
    }

    [Fact]
    public void WrongSignatureThrows()
    {
        var bytes = Sample();
        bytes[0] = 0x12; bytes[1] = 0x34;
        Assert.Throws<RulesStreamFormatException>(() => RulesStream.Parse(bytes));
    }

    /// <summary>
    /// Builds a rules stream (same format as .rwz) for tests.
    /// Real data cannot be included in the tests, so fictitious rules are written in the same format.
    /// </summary>
    private sealed class StreamBuilder
    {
        private readonly List<(string Name, bool Enabled, List<byte[]> Elements)> _rules = new();

        public StreamBuilder Rule(string name, bool enabled, Action<RuleBuilder> build)
        {
            var rb = new RuleBuilder();
            build(rb);
            _rules.Add((name, enabled, rb.Elements));
            return this;
        }

        public byte[] Build()
        {
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            w.Write(1310720u);        // signature (Outlook 2019 and later)
            w.Write(0x06140000u);     // flags
            for (int i = 0; i < 9; i++) w.Write(i == 7 ? 1u : 0u);
            w.Write((ushort)_rules.Count);
            w.Write((ushort)0);
            for (int i = 0; i < _rules.Count; i++)
            {
                var (name, enabled, elements) = _rules[i];
                if (i > 0) w.Write((ushort)0x4240);
                w.Write((ushort)0x060F);
                WriteStringObject(w, name);
                w.Write(enabled ? 1u : 0u);
                for (int k = 0; k < 4; k++) w.Write(0u);
                w.Write(0u); // dataSize (not used by the reader)
                var all = new List<byte[]> { Element(0x64, 1u, 0u, 0u), Element(0x190, 1u, 0u, 1u) };
                all.AddRange(elements);
                w.Write((ushort)all.Count);
                if (i == 0)
                {
                    w.Write((ushort)0xFFFF);
                    w.Write((ushort)1);
                    w.Write((ushort)"CRuleElement".Length);
                    w.Write(Encoding.ASCII.GetBytes("CRuleElement"));
                }
                else w.Write((ushort)0x8001);
                for (int e = 0; e < all.Count; e++)
                {
                    if (e > 0) w.Write((ushort)0x8001);
                    w.Write(all[e]);
                }
            }
            w.Write(0u); // footer (template folder length 0)
            return ms.ToArray();
        }

        private static byte[] Element(uint id, params uint[] values)
        {
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            w.Write(id);
            foreach (var v in values) w.Write(v);
            return ms.ToArray();
        }

        public static void WriteStringObject(BinaryWriter w, string s)
        {
            if (s.Length < 0xFF) w.Write((byte)s.Length);
            else { w.Write((byte)0xFF); w.Write((ushort)s.Length); w.Write((ushort)0); }
            w.Write(Encoding.Unicode.GetBytes(s));
        }
    }

    private sealed class RuleBuilder
    {
        public List<byte[]> Elements { get; } = new();

        private RuleBuilder Add(uint id, Action<BinaryWriter> body)
        {
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            w.Write(id);
            body(w);
            Elements.Add(ms.ToArray());
            return this;
        }

        public RuleBuilder From(params (string Name, string Address, bool HasEntryId)[] people) => Add(0xCB, w =>
        {
            w.Write(1u); w.Write(0u); w.Write((uint)people.Length);
            foreach (var p in people) WriteRecipient(w, p.Name, p.Address, p.HasEntryId);
            w.Write(1u); w.Write(0u);
        });

        public RuleBuilder Subject(params string[] words) => Add(0xCD, w =>
        {
            w.Write((uint)words.Length);
            foreach (var s in words) { w.Write(0u); StreamBuilder.WriteStringObject(w, s); }
        });

        public RuleBuilder ThisComputerOnly() => Add(0xEF, w => { w.Write(1u); w.Write(0u); w.Write(new byte[16]); });

        public RuleBuilder MoveTo(byte[] folder) => Add(0x12C, w =>
        {
            w.Write(1u); w.Write(0u);
            w.Write((uint)folder.Length); w.Write(folder);
            w.Write((uint)Store.Length); w.Write(Store);
            StreamBuilder.WriteStringObject(w, "Folder");
            w.Write(0u);
        });

        public RuleBuilder Sound(string path) => Add(0x136, w => { w.Write(1u); w.Write(0u); StreamBuilder.WriteStringObject(w, path); });

        public RuleBuilder Alert(string text) => Add(0x130, w => { w.Write(1u); w.Write(0u); StreamBuilder.WriteStringObject(w, text); });

        public RuleBuilder Stop() => Add(0x142, w => w.Write(0u));

        public RuleBuilder Raw(uint id, byte[] body) => Add(id, w => w.Write(body));

        /// <summary>Writes one recipient (display name, e-mail address and optionally an EntryID) as an array of MAPI property values.</summary>
        private static void WriteRecipient(BinaryWriter w, string name, string address, bool hasEntryId)
        {
            var props = new List<(ushort Type, ushort Id, byte[] Data)>
            {
                (0x1F, 0x3001, Encoding.Unicode.GetBytes(name + "\0")),
                (0x1F, 0x3003, Encoding.Unicode.GetBytes(address + "\0")),
            };
            if (hasEntryId) props.Add((0x102, 0x0FFF, [1, 2, 3, 4]));

            int headerSize = props.Count * 16;
            var data = new MemoryStream();
            var headers = new MemoryStream();
            var hw = new BinaryWriter(headers);
            foreach (var p in props)
            {
                hw.Write(p.Type); hw.Write(p.Id); hw.Write(0u);
                hw.Write((uint)(headerSize + data.Length)); // position of the value (from the start of the array)
                hw.Write((uint)p.Data.Length);
                data.Write(p.Data);
            }
            w.Write(0u);
            w.Write((uint)props.Count);
            w.Write((uint)(headerSize + data.Length));
            w.Write(headers.ToArray());
            w.Write(data.ToArray());
        }
    }
}
