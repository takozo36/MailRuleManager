using OutlookRuleManager.Core;
using static OutlookRuleManager.Core.Tests.TestRules;

namespace OutlookRuleManager.Core.Tests;

public class RuleListEditorTests : JapaneseTestBase
{
    [Fact]
    public void NoChangesRightAfterLoading()
    {
        var ed = new RuleListEditor(Simple(3));
        var plan = ed.BuildPlan();
        Assert.False(plan.HasChanges);
        Assert.Equal(0, plan.ChangeCount);
        Assert.False(ed.CanUndo);
    }

    [Fact]
    public void MoveUp_AdjacentSelectionMovesTogether()
    {
        var ed = new RuleListEditor(Simple(5));
        Assert.True(ed.MoveUp(["R3", "R4"]));
        Assert.Equal("R1,R3,R4,R2,R5", Names(ed.Entries));
    }

    [Fact]
    public void MoveUp_FirstRuleDoesNotMove()
    {
        var ed = new RuleListEditor(Simple(3));
        Assert.False(ed.MoveUp(["R1"]));
        Assert.Equal("R1,R2,R3", Names(ed.Entries));
        Assert.False(ed.CanUndo);
    }

    [Fact]
    public void MoveUp_WhileFiltered_SkipsHiddenRulesAndKeepsThemInPlace()
    {
        var ed = new RuleListEditor(Simple(5));
        // R2 and R3 are hidden. Moving R4 up puts it before the previous visible rule R1; R2 and R3 stay where they are
        Assert.True(ed.MoveUp(["R4"], ["R1", "R4", "R5"]));
        Assert.Equal("R4,R2,R3,R1,R5", Names(ed.Entries));
    }

    [Fact]
    public void MoveDown()
    {
        var ed = new RuleListEditor(Simple(4));
        Assert.True(ed.MoveDown(["R1", "R3"]));
        Assert.Equal("R2,R1,R4,R3", Names(ed.Entries));
    }

    [Fact]
    public void MoveToTopAndBottom_KeepTheOrderOfTheSelection()
    {
        var ed = new RuleListEditor(Simple(5));
        ed.MoveToTop(["R4", "R2"]);
        Assert.Equal("R2,R4,R1,R3,R5", Names(ed.Entries));
        ed.MoveToBottom(["R2", "R1"]);
        Assert.Equal("R4,R3,R5,R2,R1", Names(ed.Entries));
    }

    [Fact]
    public void MoveToPosition_OutOfRangeIsClamped()
    {
        var ed = new RuleListEditor(Simple(5));
        ed.MoveTo(["R5"], 2);
        Assert.Equal("R1,R5,R2,R3,R4", Names(ed.Entries));
        ed.MoveTo(["R1"], 999);
        Assert.Equal("R5,R2,R3,R4,R1", Names(ed.Entries));
    }

    [Fact]
    public void Swap()
    {
        var ed = new RuleListEditor(Simple(4));
        Assert.True(ed.Swap("R1", "R4"));
        Assert.Equal("R4,R2,R3,R1", Names(ed.Entries));
    }

    [Fact]
    public void DragAndDrop_MovesBeforeTarget_DroppingOnItselfChangesNothing()
    {
        var ed = new RuleListEditor(Simple(5));
        Assert.True(ed.MoveBefore(["R5"], "R2"));
        Assert.Equal("R1,R5,R2,R3,R4", Names(ed.Entries));
        Assert.False(ed.MoveBefore(["R5"], "R5"));
        Assert.True(ed.MoveBefore(["R1"], null));
        Assert.Equal("R5,R2,R3,R4,R1", Names(ed.Entries));
    }

    [Fact]
    public void EnableDisable_Rename_Undo()
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
    public void SettingTheSameValueIsNotAChange()
    {
        var ed = new RuleListEditor(Simple(2));
        Assert.Equal(0, ed.SetEnabled(["R1"], true));
        Assert.False(ed.Rename("R1", "R1"));
        Assert.False(ed.CanUndo);
    }

    [Fact]
    public void EmptyNameIsRejected()
    {
        var ed = new RuleListEditor(Simple(1));
        Assert.Throws<ArgumentException>(() => ed.Rename("R1", "  "));
    }

    [Fact]
    public void DuplicateIsInsertedRightAfterTheOriginal_AndCountsAsNewInThePlan()
    {
        var ed = new RuleListEditor(Simple(3));
        var created = ed.Duplicate(["R1", "R3"], out var skipped);
        Assert.Empty(skipped);
        Assert.Equal(2, created.Count);
        Assert.Equal("R1,R1 のコピー,R2,R3,R3 のコピー", Names(ed.Entries));

        var plan = ed.BuildPlan();
        Assert.Equal(2, plan.Created.Count());
        Assert.True(plan.OrderChanged); // a new rule was inserted in the middle
        Assert.Empty(plan.Validate());
    }

    [Fact]
    public void DuplicateOnlyAtTheEnd_IsNotAReorder()
    {
        var ed = new RuleListEditor(Simple(2));
        ed.Duplicate(["R2"], out _);
        Assert.False(ed.BuildPlan().OrderChanged);
    }

    [Fact]
    public void RuleWithActionsOnlyOutlookCanSet_IsNotDuplicated()
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
    public void DuplicateOfARuleWithAMissingFolder_IsStoppedBeforeSaving()
    {
        var ed = new RuleListEditor([FromRule(1, "壊れ", ["a@example.com"], folder: null)]);
        ed.Duplicate(["R1"], out _);
        Assert.NotEmpty(ed.BuildPlan().Validate());

        // Choosing a folder makes it valid
        ed.SetMoveFolder([ed.Entries[1].Id], Folder("新"));
        Assert.Empty(ed.BuildPlan().Validate());
    }

    [Fact]
    public void DeletedRulesAreListedInTheirOriginalOrder()
    {
        var ed = new RuleListEditor(Simple(5));
        Assert.Equal(2, ed.Delete(["R4", "R2"]));
        var plan = ed.BuildPlan();
        Assert.Equal([2, 4], plan.Deleted.Select(d => d.Index));
        Assert.False(plan.OrderChanged); // the remaining rules keep their relative order
        Assert.Equal(2, plan.ChangeCount);
    }

    [Fact]
    public void ChangeFolder_SettingTheOriginalFolderAgainIsNoChange()
    {
        var ed = new RuleListEditor(Simple(2)); // move-to folder is A
        Assert.Equal(1, ed.SetMoveFolder(["R1"], Folder("B")));
        Assert.True(ed.Entries[0].IsFolderChanged);
        Assert.Equal("B", RuleText.MoveTarget(ed.Entries[0].Actions).Split('\\')[^1]);

        ed.SetMoveFolder(["R1"], Folder("A"));
        Assert.False(ed.Entries[0].IsFolderChanged);
    }

    [Fact]
    public void RulesWithoutAMoveAction_DoNotGetAFolder()
    {
        var rule = new RuleData { Index = 1, Name = "削除だけ", Actions = [new RuleAction(ActionType.Delete)] };
        var ed = new RuleListEditor([rule]);
        Assert.Equal(0, ed.SetMoveFolder(["R1"], Folder("B")));
    }

    [Fact]
    public void SavedStateIsBuiltWithoutReloading()
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

        // Starting from the rebuilt state shows "no changes"
        Assert.False(new RuleListEditor(saved).BuildPlan().HasChanges);
    }

    [Fact]
    public void ReorderDescriptionShowsTheNumberOfMovedRules()
    {
        var ed = new RuleListEditor(Simple(3));
        ed.Swap("R1", "R3");
        var plan = ed.BuildPlan();
        Assert.True(plan.OrderChanged);
        Assert.Contains(plan.Describe(), l => l.Contains("2 件のルールの実行順が変わります"));
    }
}
