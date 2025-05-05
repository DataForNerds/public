$msdata = "https://learn.microsoft.com/en-us/defender-endpoint/attack-surface-reduction-rules-reference"
$d4nData = Invoke-WebRequest "https://raw.datafornerds.io/ms/msother/msasrguid.json" | Select-Object -ExpandProperty Content | ConvertFrom-Json

$pageData = Invoke-WebRequest $msdata -UseBasicParsing

If($pageData.StatusCode -ne 200) {
    Throw "Error $($pageData.StatusCode) Getting Page Data"
}

$tables = [regex]::New('(?msi)<table>(?:.*?)<tbody>(.*?)<\/tbody>').Matches($pageData.Content)

$guidRows = [regex]::New('(?msi)<tr>(.*?)<\/tr>').Matches($tables[6].Groups[1].Value)

$guids = New-Object System.Collections.ArrayList

$guidRows.ForEach{
    
    $cellData = [regex]::New('(?msi)<td(?:[^>]*)>(.*?)<\/td>').Matches($_.Groups[1].Value)
    
    $asrName = $cellData[0].Groups[1].value.Trim() -replace "(<(?:.*?)>)",""
    $asrGuid = $cellData[1].Groups[1].Value.Trim() -replace "(<(?:.*?)>)",""
    
    $guids.Add(
        [PSCustomObject]@{
            AsrName = $asrName
            AsrGuid = $asrGuid
        }
    ) | Out-Null
}

$guids = $guids | Sort-Object AsrGuid | Select-Object AsrName,AsrGuid -Unique

$outputData = [PSCustomObject]@{
    "DataForNerds"=[PSCustomObject]@{
        "LastUpdatedUTC" = (Get-Date).ToUniversalTime()
        "SourceList" = @($msdata)
    }
    "Data" = $guids
}

$allProperties = $guids[0].psobject.Properties.Name

If(Compare-Object $d4nData.Data $outputData.Data -Property $allProperties) {
    $outputFolder = Resolve-Path (Join-Path $PSScriptRoot -ChildPath "../../../content/ms/msother")
    $outputFile = Join-Path $outputFolder -ChildPath "msasrguid.json"

    $jsonData = $outputData | ConvertTo-Json
    [System.IO.File]::WriteAllLines($outputFile, $jsonData)
} else {
    Write-Host "The data has not changed."
}
