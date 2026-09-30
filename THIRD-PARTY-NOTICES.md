# Third-Party Notices

Created: 2026-09-29
AI agent: Claude Code
Model: Claude Opus 5.5 / claude-opus-5-5

**English** | [日本語](THIRD-PARTY-NOTICES.ja.md)

This project uses the following third-party components.
The application itself (`MailRuleManager.exe`) does not include any third-party library binary.
It calls classic Outlook through COM late binding and does not reference or redistribute
the Outlook primary interop assembly (`Microsoft.Office.Interop.Outlook`).

## Outlook rules format (reimplemented in C#)

`src/OutlookRuleManager.Core/RulesStream.cs` reads the Outlook rules stream (`PR_RW_RULES_STREAM`, the same
format as `.rwz` files). The format is not documented by Microsoft; the C# implementation is based on the
format analysis and parsers in the following projects, both licensed under the MIT License.

- [asklar/rwzreader](https://github.com/asklar/rwzreader) — Copyright (c) 2021 Alexander Sklar
- [hughbe/OutlookRulesReader](https://github.com/hughbe/OutlookRulesReader) — Copyright (c) 2021 Hugh Bellamy

```
MIT License

Copyright (c) 2021 Alexander Sklar
Copyright (c) 2021 Hugh Bellamy

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Test-only packages

These packages are used only by `tests/OutlookRuleManager.Core.Tests` and are not included in the application.

| Package | Version | License |
|---|---|---|
| [xunit](https://github.com/xunit/xunit) | 2.9.3 | Apache-2.0 |
| [xunit.runner.visualstudio](https://github.com/xunit/visualstudio.xunit) | 2.8.2 | Apache-2.0 |
| [Microsoft.NET.Test.Sdk](https://github.com/microsoft/vstest) | 17.11.1 | MIT |

## Runtime

The application runs on the .NET 8 Desktop Runtime (Windows Forms), which is installed separately by the user
and is not redistributed with the framework-dependent release. .NET and Windows Forms are licensed under the MIT License.
If a self-contained build is distributed, the .NET runtime's own license and third-party notices apply to that package.

## Microsoft Outlook

Microsoft Outlook is required at runtime and is not distributed with this project.
Microsoft and Outlook are trademarks of the Microsoft group of companies.
