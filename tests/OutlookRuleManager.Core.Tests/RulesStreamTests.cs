using System.Text;
using OutlookRuleManager.Core;

namespace OutlookRuleManager.Core.Tests;

public class RulesStreamTests
{
    private static readonly byte[] FolderA = [0x00, 0x00, 0x00, 0x00, 0xAA, 0xBB, 0x01];
    private static readonly byte[] FolderB = [0x00, 0x00, 0x00, 0x00, 0xCC, 0xDD, 0x02];
    private static readonly byte[] Store = [0x01, 0x02, 0x03];

    /// <summary>典型的な 2 件（差出人→移動→中止、件名＋このコンピューターのみ＋サウンド）。</summary>
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
    public void 名前_有効_条件_処理を読める()
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
    public void RuleDataにすると移動先を解決しOutlookと同じ順に並ぶ()
    {
        var r2 = RulesStream.Parse(Sample())[1];
        var data = r2.ToRuleData(
            f => f.EntryId == Convert.ToHexString(FolderB) ? new FolderRef(@"\\me\受信トレイ\通知", "B", "S") : null,
            onThisMachine: true, isLocalRule: true);

        // Outlook は 件名 → このコンピューターのみ、移動 → 中止 → サウンド → 新着通知 の順で返す
        Assert.Equal([ConditionType.Subject, ConditionType.LocalMachineOnly], data.Conditions.Select(c => c.Type));
        Assert.Equal([ActionType.MoveToFolder, ActionType.Stop, ActionType.PlaySound, ActionType.NewItemAlert], data.Actions.Select(a => a.Type));
        Assert.Equal(@"受信トレイ\通知", data.Actions[0].Folder!.DisplayPath);
        Assert.True(data.IsLocalRule);
        Assert.Equal(2, data.Index);
    }

    [Fact]
    public void 移動先が見つからなければエラーのルールになる()
    {
        var data = RulesStream.Parse(Sample())[0].ToRuleData(_ => null);
        Assert.True(data.Actions[0].FolderMissing);
        Assert.Contains(RuleDiagnostics.Analyze([RuleEntry.FromSource(data)])["R1"], d => d.Severity == Severity.Error);
    }

    [Fact]
    public void 別のコンピューターのみのルールはOtherMachineになる()
    {
        var data = RulesStream.Parse(Sample())[1].ToRuleData(_ => null, onThisMachine: false);
        Assert.Contains(data.Conditions, c => c.Type == ConditionType.OtherMachine);
        Assert.DoesNotContain(data.Conditions, c => c.Type == ConditionType.LocalMachineOnly);
    }

    [Fact]
    public void EntryIDの無い宛先は未解決として読む()
    {
        var bytes = new StreamBuilder()
            .Rule("未解決", true, b => b.From(("誰か", "someone@example.com", false)).MoveTo(FolderA))
            .Build();
        Assert.False(RulesStream.Parse(bytes)[0].Conditions[0].Values[0].Resolved);
    }

    [Fact]
    public void 長い名前も読める()
    {
        string longName = new string('あ', 300);
        var bytes = new StreamBuilder().Rule(longName, true, b => b.Stop()).Build();
        Assert.Equal(longName, RulesStream.Parse(bytes)[0].Name);
    }

    [Fact]
    public void 知らない要素があれば例外にする()
    {
        var bytes = new StreamBuilder().Rule("謎", true, b => b.Raw(0x999, [0, 0, 0, 0])).Build();
        var ex = Assert.Throws<RulesStreamFormatException>(() => RulesStream.Parse(bytes));
        Assert.Contains("0x999", ex.Message);
    }

    [Fact]
    public void 途中で切れたデータは例外にする()
    {
        var bytes = Sample();
        Assert.Throws<RulesStreamFormatException>(() => RulesStream.Parse(bytes[..(bytes.Length / 2)]));
    }

    [Fact]
    public void 先頭の署名が違えば例外にする()
    {
        var bytes = Sample();
        bytes[0] = 0x12; bytes[1] = 0x34;
        Assert.Throws<RulesStreamFormatException>(() => RulesStream.Parse(bytes));
    }

    /// <summary>
    /// テスト用に、ルールのまとめデータ（.rwz と同じ形式）を組み立てる。
    /// 実データはテストに入れられないため、形式どおりに架空のルールを書き出す。
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
            w.Write(1310720u);        // 署名（Outlook 2019 以降）
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
                w.Write(0u); // dataSize（読み手は使わない）
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
            w.Write(0u); // フッター（テンプレートフォルダー長 0）
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
            StreamBuilder.WriteStringObject(w, "フォルダー");
            w.Write(0u);
        });

        public RuleBuilder Sound(string path) => Add(0x136, w => { w.Write(1u); w.Write(0u); StreamBuilder.WriteStringObject(w, path); });

        public RuleBuilder Alert(string text) => Add(0x130, w => { w.Write(1u); w.Write(0u); StreamBuilder.WriteStringObject(w, text); });

        public RuleBuilder Stop() => Add(0x142, w => w.Write(0u));

        public RuleBuilder Raw(uint id, byte[] body) => Add(id, w => w.Write(body));

        /// <summary>宛先 1 件（表示名・メールアドレス、任意で EntryID）を MAPI のプロパティ値の並びで書く。</summary>
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
                hw.Write((uint)(headerSize + data.Length)); // 値の位置（並びの先頭から）
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
