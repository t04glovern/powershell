function Get-TimeZone($Name)
{
    system.timezoneinfo]::GetSystemTimeZones() | 
    Where-Object { $_.ID -like "*$Name*" -or $_.DisplayName -like "*$Name*" } | 
    Select-Object -ExpandProperty ID
}