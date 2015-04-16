﻿<#***************************************************************************

Copyright (c) Microsoft Corporation 2014.
 
This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:
 
http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx
 
Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

***************************************************************************#>

function Add-DocxText {

    <#
    .SYNOPSIS
    Appends given text to a specified DOCX document.
    .DESCRIPTION
    Add-DocxText cmdlet appends given text to a specified DOCX document.  Supports adding text with
    bold, italic, underline, forecolor, backcolor, paragraph style.
    .EXAMPLE
    # Simple use
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor"
    .EXAMPLE
    # Demonstrates adding bold text
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -Bold
    .EXAMPLE
    # Demonstrates adding text with forecolor
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -ForeColor Red
    .EXAMPLE
    # Demonstrates adding text with backcolor
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -BackColor Green
    .EXAMPLE
    # Demonstrates adding text with bold, italic, underline
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -Bold -Italic -Underline
    .EXAMPLE
    # Demonstrates adding text with forecolor, backcolor
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -ForeColor White -BackColor Red
    .EXAMPLE
    # Demonstrates adding text with paragraph style
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -HeadingStyle Title
    .EXAMPLE
    # Demonstrates adding text with bold, italic, underline, forcolor, backcolor, style
    Add-DocxText .\Doc1.docx "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -Bold -Italic -Underline White Red Heading1
    #>     
  
    [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Medium')]
    param
    (
        [Parameter(Mandatory=$True, Position=0)]
        [OpenXmlPowerTools.WmlDocument]$WmlDocument,

        [Parameter(Mandatory=$True, Position=1)]
        [string]$Content,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Bold,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Italic,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Underline,

        [Parameter(Mandatory=$False)]
        [string]$ForeColor,

        [Parameter(Mandatory=$False)]
        [string]$BackColor,

        [Parameter(Mandatory=$False)]
        [string]$Style
    )

    write-verbose "Appending Text to a document"     
  
    if ($WmlDocument -ne $null)
    {
        $newWmlDoc = [OpenXmlPowerTools.AddDocxTextHelper]::AppendParagraphToDocument(
			$WmlDocument, 
			$Content,
			$Bold,
			$Italic,
			$Underline,
			$ForeColor,
			$BackColor,
			$Style)
        $WmlDocument.DocumentByteArray = $newWmlDoc.DocumentByteArray
        #return $newWmlDoc
    }
}
