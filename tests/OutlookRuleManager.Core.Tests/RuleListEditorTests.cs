using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

public class RuleListEditorTests
{
    [Fact]
    public void 読み込み直後は変更なし()
    {
        var ed = new RuleListEditor(Simple(3));
        var plan = ed.BuildPlan();
        Assert.False(plan.HasChanges);
        Assert.Equal(0, plan.ChangeCount);
        Assert.False(ed.CanUndo);
    }

    [Fact]
    public void 上へ_連続した選択はまとまって動く()
    {
        var ed = new RuleListEditor(Simple(5));
        Assert.True(ed.MoveUp(["R3", "R4"]));
        Assert.Equal("R1,R3,R4,R2,R5", Names(ed.Entries));
    }

    [Fact]
    public void 上へ_先頭にあるものは動かない()
    {
        var ed = new RuleListEditor(Simple(3));
        Assert.False(ed.MoveUp(["R1"]));
        Assert.Equal("R1,R2,R3", Names(ed.Entries));
        Assert.False(ed.CanUndo);
    }

    [Fact]
    public void 上へ_絞り込み中は表示中のルールを飛び越え非表示のルールは動かない()
    {
        var ed = new RuleListEditor(Simple(5));
        // R2, R3 が非表示。R4 を上へ → 表示中の直前 R1 の前へ入る。R2, R3 の位置は変わらない
        Assert.True(ed.MoveUp(["R4"], ["R1", "R4", "R5"]));
        Assert.Equal("R4,R2,R3,R1,R5", Names(ed.Entries));
    }

    [Fact]
    public void 下へ()
    {
        var ed = new RuleListEditor(Simple(4));
        Assert.True(ed.MoveDown(["R1", "R3"]));
        Assert.Equal("R2,R1,R4,R3", Names(ed.Entries));
    }

    [Fact]
    public void 先頭へ_末尾へ_選択の順序は保つ()
    {
        var ed = new RuleListEditor(Simple(5));
        ed.MoveToTop(["R4", "R2"]);
        Assert.Equal("R2,R4,R1,R3,R5", Names(ed.Entries));
        ed.MoveToBottom(["R2", "R1"]);
        Assert.Equal("R4,R3,R5,R2,R1", Names(ed.Entries));
    }

    [Fact]
    public void 位置を指定して移動_範囲外は端に寄せる()
    {
        var ed = new RuleListEditor(Simple(5));
        ed.MoveTo(["R5"], 2);
        Assert.Equal("R1,R5,R2,R3,R4", Names(ed.Entries));
        ed.MoveTo(["R1"], 999);
        Assert.Equal("R5,R2,R3,R4,R1", Names(ed.Entries));
    }

    [Fact]
    public void 入れ替え()
    {
        var ed = new RuleListEditor(Simple(4));
        Assert.True(ed.Swap("R1", "R4"));
        Assert.Equal("R4,R2,R3,R1", Names(ed.Entries));
    }

    [Fact]
    public void ドラッグ_指定ルールの直前へ_自分自身へのドロップは変化なし()
    {
        var ed = new RuleListEditor(Simple(5));
        Assert.True(ed.MoveBefore(["R5"], "R2"));
        Assert.Equal("R1,R5,R2,R3,R4", Names(ed.Entries));
        Assert.False(ed.MoveBefore(["R5"], "R5"));
        Assert.True(ed.MoveBefore(["R1"], null));
        Assert.Equal("R5,R2,R3,R4,R1", Names(ed.Entries));
    }

    [Fact]
    public void 有効無効_名前変更_元に戻す()
    {
        var ed = new RuleListEditor(Simple(3));
        Assert.Equal(2, ed.SetEnabled(["R1", "R2"], false));
        Assert.True(ed.Rename("R3", "  新しい名前 "));
        Assert.Equal("新しい名前", ed.Entries[2].Name);
        Assert.Equal(3, ed.BuildPlan().ChangeCount);

        Assert.True(ed.Undo());
        Assert.Equal("R3", ed.Entries[2].Name);
        Assert.True(ed.Undo());
        Assert.True(ed.Entries.All(e => e.Enabled));
        Assert.False(ed.BuildPlan().HasChanges);
    }

    [Fact]
    public void 同じ値を設定しても変更扱いにならない()
    {
        var ed = new RuleListEditor(Simple(2));
        Assert.Equal(0, ed.SetEnabled(["R1"], true));
        Assert.False(ed.Rename("R1", "R1"));
        Assert.False(ed.CanUndo);
    }

    [Fact]
    public void 空の名前は拒否()
    {
        var ed = new RuleListEditor(Simple(1));
        Assert.Throws<ArgumentException>(() => ed.Rename("R1", "  "));
    }

    [Fact]
    public void 複製は元の直後に入り_計画では新規として数える()
    {
        var ed = new RuleListEditor(Simple(3));
        var created = ed.Duplicate(["R1", "R3"], out var skipped);
        Assert.Empty(skipped);
        Assert.Equal(2, created.Count);
        Assert.Equal("R1,R1 のコピー,R2,R3,R3 のコピー", Names(ed.Entries));

        var plan = ed.BuildPlan();
        Assert.Equal(2, plan.Created.Count());
        Assert.True(plan.OrderChanged); // 途中に新規が入る
        Assert.Empty(plan.Validate());
    }

    [Fact]
    public void 末尾だけに複製が入ったときは並び替えなし扱い()
    {
        var ed = new RuleListEditor(Simple(2));
        ed.Duplicate(["R2"], out _);
        Assert.False(ed.BuildPlan().OrderChanged);
    }

    [Fact]
    public void 画面でしか設定できない処理を含むルールは複製しない()
    {
        var rule = new RuleData
        {
            Index = 1,
            Name = "スクリプト",
            Actions = [new RuleAction(ActionType.RunScript)],
        };
        var ed = new RuleListEditor([rule]);
        var created = ed.Duplicate(["R1"], out var skipped);
        Assert.Empty(created);
        Assert.Single(skipped);
        Assert.False(ed.CanUndo);
    }

    [Fact]
    public void 移動先が消えたルールの複製は保存前に止める()
    {
        var ed = new RuleListEditor([FromRule(1, "壊れ", ["a@example.com"], folder: null)]);
        ed.Duplicate(["R1"], out _);
        Assert.NotEmpty(ed.BuildPlan().Validate());

        // 移動先を指定し直せば通る
        ed.SetMoveFolder([ed.Entries[1].Id], Folder("新"));
        Assert.Empty(ed.BuildPlan().Validate());
    }

    [Fact]
    public void 削除は計画の削除一覧に読み込み時の順で入る()
    {
        var ed = new RuleListEditor(Simple(5));
        Assert.Equal(2, ed.Delete(["R4", "R2"]));
        var plan = ed.BuildPlan();
        Assert.Equal([2, 4], plan.Deleted.Select(d => d.Index));
        Assert.False(plan.OrderChanged); // 残りの前後関係は変わらない
        Assert.Equal(2, plan.ChangeCount);
    }

    [Fact]
    public void 移動先の変更_元のフォルダーに戻すと変更なし()
    {
        var ed = new RuleListEditor(Simple(2)); // 移動先は A
        Assert.Equal(1, ed.SetMoveFolder(["R1"], Folder("B")));
        Assert.True(ed.Entries[0].IsFolderChanged);
        Assert.Equal("B", RuleText.MoveTarget(ed.Entries[0].Actions).Split('\\')[^1]);

        ed.SetMoveFolder(["R1"], Folder("A"));
        Assert.False(ed.Entries[0].IsFolderChanged);
    }

    [Fact]
    public void 移動処理のないルールには移動先を設定しない()
    {
        var rule = new RuleData { Index = 1, Name = "削除だけ", Actions = [new RuleAction(ActionType.Delete)] };
        var ed = new RuleListEditor([rule]);
        Assert.Equal(0, ed.SetMoveFolder(["R1"], Folder("B")));
    }

    [Fact]
    public void 保存後の状態を読み込み直さずに組み立てる()
    {
        var ed = new RuleListEditor(Simple(4));
        ed.Delete(["R2"]);
        ed.Duplicate(["R4"], out _);
        ed.MoveToTop(["R4"]);
        ed.Rename("R1", "一番");
        ed.SetEnabled(["R3"], false);
        ed.SetMoveFolder(["R3"], Folder("B"));

        var saved = ed.BuildPlan().ToSavedSnapshot();
        Assert.Equal(["R4", "一番", "R3", "R4 のコピー"], saved.Select(r => r.Name));
        Assert.Equal([1, 2, 3, 4], saved.Select(r => r.Index));
        Assert.False(saved[2].Enabled);
        Assert.EndsWith(@"\B", saved[2].Actions[0].Folder!.Path);

        // 組み立てた状態から始めると「変更なし」になる
        Assert.False(new RuleListEditor(saved).BuildPlan().HasChanges);
    }

    [Fact]
    public void 並び替えの説明に件数が出る()
    {
        var ed = new RuleListEditor(Simple(3));
        ed.Swap("R1", "R3");
        var plan = ed.BuildPlan();
        Assert.True(plan.OrderChanged);
        Assert.Contains(plan.Describe(), l => l.Contains("2 件のルールの実行順が変わります"));
    }
}
