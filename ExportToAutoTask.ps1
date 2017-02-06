. ./Data/AutoTaskPropertiesOverrides.ps1

function ExportToAutoTask ($InputExportFilename, $InputAutoTaskImportFilename)
{
    $TimeStamp = $(Get-Date -Format "s").Replace(":", "").Replace("-", "")

    if (-Not (Get-Module -ListAvailable ImportExcel)) {
        Write-Host "ImportExcel module is not available"
        Write-Host "Use 'Install-Module ImportExcel' to install for all users"
        Write-Host "Use 'Install-Module ImportExcel -Scope CurrentUser' to install for current user"
        Write-Error -Message "ImportExcel module is not available"
        return;
    }

    $ExportFilename = $InputExportFilename.Replace(".xlsx", "-" +$TimeStamp + ".xlsx")
    
    if ($ExportFilename -eq $InputExportFilename)
    {
        Write-Error -Message "Input SalesDouble filename must end in .xlsx. Invalid filename: $InputExportFilename"
        return;
    }

    $AutoTaskImportFilename = $InputAutoTaskImportFilename.Replace(".csv", "-" +$TimeStamp + ".csv")
    
    if ($AutoTaskImportFilename -eq $InputAutoTaskImportFilename)
    {
        Write-Error -Message "Output autotask filename must end in .csv. Invalid filename: $InputAutoTaskImportFilename"
        return;
    }

    Copy-Item $InputExportFilename $ExportFilename

    Write-Host "Processing $ExportFilename..."

    Import-Excel -Path $ExportFilename |
        ForEach-Object {

            $_.PSObject.Properties | ForEach-Object {
                if (!$_.Value) {
                    $_.Value = "";
                }
            }

            $AutoTaskRow = Get-AutoTaskRow

            $AutoTaskRow
        } |
        Export-Csv -Path $AutoTaskImportFilename -NoTypeInformation
    
    Write-Host "$AutoTaskImportFilename Saved."    
}

function Get-AutoTaskRow
{
    $AutoTaskProperties = Get-AutoTaskProperties
    $AutoTaskPropertiesOverrides = Get-AutoTaskPropertiesOverrides

    $AutoTaskProperties = Merge-AutoTaskProperties $AutoTaskProperties $AutoTaskPropertiesOverrides

    $AutoTaskRow = New-Object -TypeName PSObject -Property $AutoTaskProperties

    return $AutoTaskRow
}

function Merge-AutoTaskProperties ($AutoTaskProperties, $AutoTaskPropertiesOverrides)
{
    foreach ($key in $AutoTaskPropertiesOverrides.Keys) {
        $AutoTaskProperties[$key] = $AutoTaskPropertiesOverrides[$key];
    }

    return $AutoTaskProperties;
} 