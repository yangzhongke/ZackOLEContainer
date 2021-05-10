# ZackOLEContainer

Embed documents in Windows form applications, like Powerpoint(ppt, pptx), Word(doc,docx), Excel(xls,xlsx).

This library supports .NET Core and .NET Framework.

Step one:
```
Install-Package ZackOLEContainer.WinFormCore
```

Step two:
Put an OLEContainer control on a form.

Step three:
```
this.oleContainer.OpenFile(@"E:\a.pptx");
```