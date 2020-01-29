# OfficeDrawIo

OfficeDrawIo is a Microsoft Office add-in that allows embedding and inline editing of [Draw.io](https://about.draw.io/) diagrams into **Word** and **PowerPoint** documents. Users without the add-in are able to view the diagrams but not edit them.

Check your installed version (bitness) of Microsoft Office before picking the right installer. x86 is for 32-bit Office and x64 is for 64-bit office.

**Note:** *Version 2 is NOT backwards compatible with version 1. If you have any documents using version 1 of OfficeDrawIo, please extract the document images first (extracted as DrawIo editable vector png files) using the *Export* add-in option in Word before installing the new version.*

Runtime Requirements
-
- Windows OS with [.NET Framework](https://dotnet.microsoft.com/download/dotnet-framework) 4.6 or better.
- Office 2017 or better (may work in earlier versions but not tested).
- Draw.io [Desktop](https://about.draw.io/integrations/).

Troubleshooting
-
1. The add-in does not show up in the Word or PowerPoint ribbon.

    - Make sure you installed the correct installer for your Office Version. If in doubt you can install both. 
    
    - Right click onto the PowerPoint ribbon area and select **Customize the Ribbon...** -> **Add-ins** -> **Manage: COM Add-ins** -> **Go...**, make sure that the *OfficeDrawIoPpt* or *OfficeDrawIoWord* add-in in the **Add-ins available** list is checked.

Screenshots
-
[![raspikey-diagram](screen1_tn.png)](screen1.png)

