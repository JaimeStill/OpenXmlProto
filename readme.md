# Excel Generation Prototype

This repository captures research into generating excel documents with the OpenXML SDK.

## Links

* [nuget - DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)
* [GitHub - OfficeDev/Open-XML-SDK](https://github.com/OfficeDev/Open-XML-SDK)
* [MSDN - Open XML SDK Docs](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk)
* [Open XML SDK Productivity Tool - Install BOTH](https://www.microsoft.com/en-us/download/details.aspx?id=30425)

## Running

```bash
dotnet run -- {manifest-title}
```

This will generate `{manifest-title}.xlsx` in the directory that the  command is executed from.