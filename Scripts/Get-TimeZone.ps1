function Get-TimeZone($Name)
{
<#
.Synopsis
Retrieves Timezone data
.DESCRIPTION
Retrieves Timezone information for a given locale
#>
    system.timezoneinfo]::GetSystemTimeZones() | 
    Where-Object { $_.ID -like "*$Name*" -or $_.DisplayName -like "*$Name*" } | 
    Select-Object -ExpandProperty ID
}