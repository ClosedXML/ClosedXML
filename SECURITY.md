# Security Policy

## Supported Versions

Only the latest version without a label (e.g. rc, beta) released through [NuGet.org](https://www.nuget.org/packages/ClosedXML) is supported.

## NETStandard dependencies

ClosedXML directly or indirectly depends on [NETStandard.Library](https://www.nuget.org/packages/NETStandard.Library).
The netStandard package references some other packages with minimal versions that might be marked as vulnerable.

One of them is `System.Text.RegularExpressions@4.3.0` or `System.Net.Http@4.3.0`, generally through `Irony.NetCore` package.

* ClosedXML@0.100.3 > XLParser@1.5.2 > Irony.NetCore@1.0.11 > NETStandard.Library@1.6.1 > System.Xml.XDocument@4.3.0 > System.Xml.ReaderWriter@4.3.0 > System.Text.RegularExpressions@4.3.0
* ClosedXML@0.100.3 > XLParser@1.5.2 > Irony.NetCore@1.0.11 > NETStandard.Library@1.6.1 > System.Net.Http@4.3.0

**These reports are false positives.**

Netstandard is only a specification of API, not an implementation and it defers to actual implementation that is being maintained.
* For .NET Framework, NETStandard uses a facade dll to forward types of the installed .NET Framework.
* For .NET Core, NETStandard uses the implementation from the selected .NET Core version of the application (e.g. 7 for net7).

The vulnerability through NETStandard reference is a security issue only if the application is running an obsolete framework/core (e.g. 4.5 or core2).

That is also position of the team that maintained .NET Standard: https://github.com/dotnet/standard/issues/1786

> System.Text.RegularExpressions never applies a vulnerable binary on .NETFramework. It applies a facade dll that typeforwards to System.dll where all this code lives. The facade dll is not vulnerable as it does not contain the code. System.Text.RegularExpressions also does not apply its binary on .NETCore2.0 and later. There the implementation is provided by the shared framework. This package only exists for delivering the implementation to older frameworks (.netcore1.x), which are now out of support.
> In general we don't churn the entire package ecosystem when a single package is updated. If you'd like to update your package reference to suppress this false positive from a validation tool you may. This wouldn't be much different than if we shipped a new version of NETStandard.Library, you'd still need all the packages that referenced the old version to update to a new one.

## Reporting a Vulnerability

Please file an issue with a *SECURITY* as the first word in the title of the issue.