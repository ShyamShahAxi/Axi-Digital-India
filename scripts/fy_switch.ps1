# Returns the financial-year switch script for a report page, from data/mappings/report_years.json.
# To add a year, add its page to that file and rebuild every year's dashboard.
param([Parameter(Mandatory)][string]$Current)
$y = Get-Content "$PSScriptRoot\..\data\mappings\report_years.json" -Raw -Encoding utf8 | ConvertFrom-Json
$json = [ordered]@{ pagesBase = $y.pagesBase; current = $Current; years = @($y.years) } | ConvertTo-Json -Depth 4 -Compress
(Get-Content "$PSScriptRoot\..\reports\templates\fy_switch.partial.html" -Raw -Encoding utf8).Replace('/*__YEARS__*/null', $json)
