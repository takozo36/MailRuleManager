# Technical notes (what was measured about the Outlook object model)

Created: 2026-09-29
AI agent: Claude Code
Model: Claude Opus 5.5 / claude-opus-5-5

**English** | [日本語](TECHNICAL_NOTES.ja.md)

Everything here is behavior observed in the following environment. Other Outlook environments may behave differently.

- Windows 11 x64, classic Outlook x64 (Microsoft 365, 16.0.20430)
- POP / IMAP .pst store, about 340 client-side rules

## How saving works, and the order of steps

- Changes to the `Rules` collection returned by `Store.GetRules()` (add, remove, reorder, property changes) are not stored in Outlook
  until `Rules.Save()` is called. Getting `Rules` again shows the state before the changes. If a step fails and Save is not called, nothing is kept.
- **Touching a rule object after `Rules.Remove()` crashes Outlook** (access violation; it actually crashed).
- A rule created with `Rules.Create()` is inserted at execution order 1 (the top).
- Changing `Rule.ExecutionOrder` immediately re-sorts the indexes of `Rules`.
- Getting the same rule again with `Rules[i]` returns the same RCW (the .NET wrapper of the COM object), so identity can be checked with `ReferenceEquals`.

Based on this, saving takes these steps (`OutlookRuleGateway.Apply`):

1. Get `GetRules()` again and check that the count and the name at each position match the state at load time (cancel if not)
2. Change the name, move-to folder and enabled flag of existing rules
3. Create duplicates with `Rules.Create()` and copy the conditions, exceptions and actions. **Read them back and compare with the original**
4. Place the kept rules from position 1 with `ExecutionOrder` (the rules to delete are pushed to the end)
5. Confirm with `ReferenceEquals` that the last rule is one to delete, then `Remove()` it. Never touch that object again
6. Check the final order with `ReferenceEquals`
7. `Rules.Save()`

With `dryRun: true`, `Apply` skips only step 7. Steps 1–6 were confirmed on a real mailbox (including the comparison of 3 duplicates).

## Calling Outlook with late binding (dynamic)

The interop assembly (NuGet `Microsoft.Office.Interop.Outlook`) is not used. Outlook is called with
`Type.GetTypeFromProgID("Outlook.Application")` and `dynamic`, because that package is a repackaging by an individual
and its description says it is unsupported and has no license.

- Loading speed was the same as with early binding (average over 20 rules: early binding 208 ms/rule, dynamic 196 ms/rule).
- **Assigning a property whose value is an object or an array fails with `dynamic`**
  (e.g. `rule.Actions.MoveToFolder.Folder = folder` fails with "The operation failed", 0x80020009).
  Calling IDispatch PROPERTYPUT directly with `Type.InvokeMember(name, BindingFlags.SetProperty, ...)` succeeds (`OutlookCom.SetProperty`).
  The same assignment succeeds with early binding. Strings, booleans and numbers can be assigned with `dynamic` without problems.

## Loading: reading rule by rule through the object model is slow

- Each rule has 31 condition, 31 exception and 28 action objects, and which ones are enabled can only be found out by asking one by one.
- The first access to each object is slow (scanning the 31 conditions: 111.5 ms/rule the first time, 39 ms/rule the second time).
- Switching the caller between early and late binding, or between STA and MTA, did not bring it below 0.2–0.35 s per rule (about 100 s for 342 rules).

## Loading: reading the rules data directly (fast load)

Outlook stores all rules in `PR_RW_RULES_STREAM` (`0x68020102`, binary) of a hidden message
(message class `IPM.RuleOrganizer`) in the Inbox.

- How to get it: find the EntryID of `IPM.RuleOrganizer` with `GetTable("", olHiddenItems)` on the Inbox, then read it with
  `PropertyAccessor.GetProperty` of `Namespace.GetItemFromID(EntryID, StoreID)` (which returns a StorageItem).
  535,714 bytes for 342 rules were read in about 1 second.
  - `Folder.GetStorage("IPM.RuleOrganizer", olIdentifyByMessageClass)` failed with "not found".
  - `GetStorage` with the EntryID did not return the existing item but a new, empty StorageItem (not kept unless saved).
  - Adding this property as a table column made Outlook return "Out of memory".
- The format is the same as an exported rules file (.rwz). Compared with an .rwz exported at the same time, the size was identical and
  the only difference was 4 bytes per rule (the separator between rules and the rule signature: `0x00140000` in .rwz, `0x060F4240` in the stream).
- The format is not documented, but it is implemented in C# (`RulesStream.cs`) based on the analysis of
  [asklar/rwzreader](https://github.com/asklar/rwzreader) and [hughbe/OutlookRulesReader](https://github.com/hughbe/OutlookRulesReader) (both MIT).
  Parsing 342 rules takes 16 ms.
- Elements carry no length, so an unknown element makes the rest unreadable. In that case the app falls back to reading rule by rule.

What the rules data alone cannot tell is asked from Outlook (a few ms per rule):

| Item | Reason |
|---|---|
| Name and enabled flag of each rule | Checks that the rules data matches Outlook. If any rule differs, the app falls back to reading rule by rule |
| `Rule.IsLocalRule` | True even without "on this computer only" when the rule contains client-only actions such as a sound or a new item alert |
| Whether "on this computer only" means this computer | The rules data only has a computer identifier. Checked with `Conditions.OnLocalMachine.Enabled` |
| Path of the move-to / copy-to folder | The rules data only has the EntryID and the folder name. Resolved with `GetFolderFromID`; not found means an error rule (each folder is looked up once) |

Conditions, exceptions and actions are sorted into the order the object model returns them (fixed per type).

Comparison on a real mailbox (342 rules): fast load 2.3 s, rule-by-rule load 99.4 s.
Name, enabled flag, type, client-only flag, conditions, exceptions, actions (recipient addresses, words, move-to folder EntryID and path),
recipient resolution and diagnostics all matched.

## Detecting errors

- A move or copy action is enabled but `MoveToFolder.Folder` (`CopyToFolder.Folder`) is null or throws → the folder no longer exists.
  This is the typical cause of rules shown as "error" in Outlook.
- Exchange server-side rules have an `ST_ERROR` flag in `PR_RULE_MSG_STATE`, but client-side rules in a .pst do not, so it is not used.

## Detecting rules that never run

If an earlier enabled rule A stops processing, has no exceptions, and all of A's conditions cover the conditions of a later rule B,
every message that matches B stops at A, so B never runs. "Covers" is decided as follows:

| Condition type | A covers B when |
|---|---|
| From / sent to (matches any of the addresses) | B's addresses ⊆ A's addresses |
| Words in subject, body, ... (contains any of the words) | each of B's words contains one of A's words |
| Category | B's categories ⊆ A's categories |
| Conditions without values (has attachment, ...) | B has the same type |
| Others | the sets of values are equal |

"On this computer only" only decides where the rule runs, so it is left out of the comparison. When A has a single From condition,
senders that only partly overlap with B are also reported individually.

## Store types

- Only stores that support rules (here, a POP/IMAP .pst) can be used. For stores such as a calendar-only store,
  `GetRules()` fails with "Rules are not supported for this store".
