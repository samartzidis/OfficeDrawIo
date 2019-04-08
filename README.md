# OfficeDrawIo

OfficeDrawIo is a Microsoft Office add-in that allows embedding and inline editing of [Draw.io](https://about.draw.io/) diagrams into Word documents. Users without the add-in will still be able to see the diagrams but not edit them.

Known issues
------------
- Only Word is currently supported due to VSTO custom control restrictions.
- Copy-paste (effectively cloning) of complete Word Draw.io diagram controls does not work across different Word documents, only within the same document. Workaround is to create a new diagram in the target document and then copy the diagram contents from the first document to it.

Runtime Requirements
--------------------
- Windows 10 ([.NET Framework](https://dotnet.microsoft.com/download/dotnet-framework) 4.6 or better).
- Office 2017 (may work in earlier versions but not tested).
- Drawio [Desktop](https://about.draw.io/integrations/).

Screenshots
-----------
[![raspikey-diagram](screen1_tn.png)](screen1.png)

