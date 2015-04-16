Open XML Power Tools
============================================

New Installation Procedures
---------------------------
1)  Make sure you are running PowerShell 3.0 or later
2)  If necessary, run PowerShell as administrator, Set-ExecutionPolicy Unrestricted
3)  Install the Open XML SDK 2.5  http://bit.ly/1qMaf6i
4)  Download and unzip PowerTools for Open XMl 3.1.0 or later:  http://bit.ly/1ss8hfV
    Unzip into C:\Users\<user>\Documents\WindowsPowerShell\Modules\Oxpt
	(Create WindowsPowerShell and Modules directories as needed)
5)  Import-Module OxPt
6)  Visual Studio not needed!!!

Version 3.1.07 : February 9, 2015
- Added Merge-Pptx Cmdlet
- Added New-Pptx Cmdlet
- Added New-PmlDocument
- Fixed help for Merge-Docx
- Don't throw duplicate attribute exception when running FormattingAssembler.AssembleFormatting
  twice on same document.

Version 3.1.06 : February 7, 2015
- Added Expand-DocxFormatting Cmdlet
- Cmdlets do not keep a handle to the current directory, preventing deletion of the directory.
- Added additional tests to Test-OxPtCmdlets

Version 3.1.05 : January 29, 2015
- Added GetListItemText_zh_CN.cs
- Fixed GetListItemText_fr_FR.cs
- Partially fixed GetListItemText_ru_RU.cs
- Fixed GetListItemText_Default.cs
- Added better support in ListItemRetriever.cs
- Added FileUtils class in PtUtil.cs

Version 3.1.04 : December 17, 2014
- Added Get-DocxMetrics Cmdlet
- Added New-WmlDocument Cmdlet
- Added MetricsGetter.cs module
- Added MettricsGetter01.cs module, along with sample documents
- Reworked Add-DocxText, new style of using it with New-WmlDocument

Version 3.1.03 : December 9, 2014
- Added ChartUpdater.cs module
- Added ChartUpdater01.cs module, along with sample documents
- Added Test-OxPtCmdlets Cmdlet

Version 3.1.02 : December 1, 2014
- Added Add-DocxText Cmdlet

Version 3.1.01 : November 23, 2014
- Added Convert-DocxToHtml Cmdlet
- Added Chinese and Hebrew sample documents
- Cmdlets in this release
	Clear-DocxTrackedRevision
	Convert-DocxToHtml
	ConvertFrom-Base64
	ConvertFrom-FlatOpc
	ConvertTo-Base64
	ConvertTo-FlatOpc
	Get-OpenXmlValidationErrors
	Merge-Docx
	New-Docx
	Test-OpenXmlValid

Version 3.1.00 : November 13, 2014
- Changed installation process - no longer requires compilation using Visual Studio
- Added ConvertTo-FlatOpc Cmdlet
- Added ConvertFrom-FlatOpc Cmdlet
- Changed parameters for Test-OpenXmlValid, Get-OpenXmlValidationErrors
- Removed the unnecessary 1/2 second sleep when doing Word automation in the New-Docx Cmdlet

Version 3.0.00 : October 29, 2014
- New release of cmdlets that are written as 'Advanced Functions' instead of in C#.

Procedures for enhancing OxPt
-----------------------------
There are a variety of things to do when adding a new CmdLet to OxPt:
- Write the new CmdLet.  Put it in OxPtCmdlets
- Modify OxPt.psm1
    Call the new Cmdlet script to make the function available
    Modify Export-ModuleMember function to export the Cmdlet and any aliases
- Update Readme.txt, describing the enhancement
- Add a new test to Test-OxPtCmdlets.ps1
- Update Downloads page on CodePlex