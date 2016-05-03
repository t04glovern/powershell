<#
.Synopsis
   Gets the bitness of the office 365/2013 install
.DESCRIPTION
   Gets the bitness of the office 365/2013 install
.EXAMPLE
   Get-O365Bitness -ComputerName MyPC
.EXAMPLE
   Get-Content -File ComputerNames.txt | Get-O365Bitness
#>
[CmdletBinding()]
[Alias()]
Param
(
    # Param1 help description
    [Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                Position=0)]
    [string[]]$ComputerName
)

Begin
{
    $ErrorActionPreference = "SilentlyContinue"
}
Process
{
    foreach ($comp in $ComputerName)
    {
        if (Test-Connection $comp -Quiet -Count 1)
        {
            $out = Write-Output "$comp`t"
            $out += Invoke-Command -ComputerName $comp -ScriptBlock { (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Office\15.0\Outlook).Bitness }
            $out += Invoke-Command -ComputerName $comp -ScriptBlock { (Get-ItemProperty HKLM:\Software\Microsoft\Office\15.0\Outlook).Bitness }
            Write-Output $out
        }
    }
}
End
{
}
