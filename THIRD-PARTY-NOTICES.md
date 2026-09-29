# Third-Party Notices

This project uses the following third-party components.
The application itself (`Outlook仕訳ルール管理.exe`) does not include any third-party library.
It calls Outlook Classic through COM late binding and does not reference or redistribute
the Outlook primary interop assembly (`Microsoft.Office.Interop.Outlook`).

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
