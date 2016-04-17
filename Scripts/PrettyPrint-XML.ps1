function PrettyPrint-XML
{
<#
.Synopsis
   Pretty-print an XML string
.DESCRIPTION
   Pretty-print an XML string with nice indenting
.EXAMPLE
   Get-AppLockerPolicy -Effective -XML | PrettyPrint-XML
#>
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # XML String
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [string]$xml
    )

    Begin
    {
        $Doc=New-Object system.xml.xmlDataDocument 
    }
    Process
    {
        $doc.LoadXml(($xml)) 
        $sw=New-Object system.io.stringwriter 
        $writer=New-Object system.xml.xmltextwriter($sw) 
        $writer.Formatting = [System.xml.formatting]::Indented 
        $doc.WriteContentTo($writer) 
        Write-OutPut $sw.ToString()
    }
    End
    {
    }
}