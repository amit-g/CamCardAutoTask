function Main
{
    param (
        [Parameter(Mandatory = $false)] [string] $InputCamCardExpotFilename = ".\Data\Personal_Contacts.xlsx",
        [Parameter(Mandatory = $false)] [string] $InputAutoTaskImportFilename = ".\Data\AutoTask.csv"
    )
    
    $TimeStamp = $(Get-Date -Format "s").Replace(":", "").Replace("-", "")

    $CamCardExpotFilename = $InputCamCardExpotFilename.Replace(".xlsx", "-" +$TimeStamp + ".xlsx")
    
    if ($CamCardExpotFilename -eq $InputCamCardExpotFilename)
    {
        Write-Error -Message "Input camcard filename must end in .xlsx. Invalid filename: $InputCamCardExpotFilename"
        return;
    }

    $AutoTaskImportFilename = $InputAutoTaskImportFilename.Replace(".csv", "-" +$TimeStamp + ".csv")
    
    if ($AutoTaskImportFilename -eq $InputAutoTaskImportFilename)
    {
        Write-Error -Message "Output autotask filename must end in .csv. Invalid filename: $InputAutoTaskImportFilename"
        return;
    }

    Copy-Item $InputCamCardExpotFilename $CamCardExpotFilename

    Write-Host "Processing $CamCardExpotFilename..."

    Import-Excel -Path $CamCardExpotFilename |
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
    $AutoTaskProperties = @{
        "[required] Account: Name" = $_."Company1";
        "Account: Number" = "";
        "Account: Address 1" = $_."Street1";
        "Account: Address 2" = "";
        "Account: City" = $_."City1";
        "Account: State" = $_."State1";
        "Account: Zip Code" = $_."Zip1";
        "Account: Country" = "";
        "Account: Additional Address Information" = "";
        "[required] Account: Phone" = $_."Telephone1";
        "Account: Alternate Phone 1" = "";
        "Account: Alternate Phone 2" = "";
        "Account: Fax" = $_."Fax1";
        "Account: Web" = $_."External Link".Replace("Personal Homepage:", "");
        "Account: Round-Trip Distance" = "";
        "Account: Account Type" = "";
        "Account: Classification" = "";
        "Account: Account Manager" = "";
        "Account: Territory Name" = "";
        "Account: Market Segment" = "";
        "Account: Competitor" = "";
        "Account: Parent Account" = "";
        "Account: Facebook URL" = "";
        "Account: Twitter URL" = "";
        "Account: LinkedIn URL" = "";
        "Account: Stock Symbol" = "";
        "Account: Stock Market" = "";
        "Account: SIC Code" = "";
        "Account: Account Detail Alert" = "";
        "Account: New Ticket Alert" = "";
        "Account: Ticket Detail Alert" = "";
        "Account: Tax Region" = "";
        "Account: Tax Exempt" = "";
        "Account: Tax ID" = "";
        "Account: Invoice Template" = "";
        "Account: Quote Template" = "";
        "Account: Quote Email Message" = "";
        "Account: Active/Inactive" = "";
        "Account UDF:29682812 Number of Users" = "";
        "Account UDF:29682815 Number of Servers" = "";
        "Account UDF:29682817 Competitive Contract Expiration Date" = "";
        "Account UDF:29682814 Lead Category" = "";
        "Account UDF:29682816 Lead Source" = "";
        "Account UDF:29682811 Sales Volume" = "";
        "Account UDF:29682805 Kaseya Customer ID" = "";
        "Site Configuration UDF:29682819 Server Password (s) [protected]" = "";
        "Contact: External ID" = "";
        "Contact: Prefix" = "";
        "[required] Contact: First Name" = $_."First Name";
        "Contact: Middle Initial" = "";
        "[required] Contact: Last Name" = $_."Last Name";
        "Contact: Suffix" = "";
        "Contact: Title" = "";
        "[required] Contact: Email Address" = $_."Email1";
        "Contact: Email Address 2" = "";
        "Contact: Email Address 3" = "";
        "Contact: Address 1" = $_."Street1";
        "Contact: Address 2" = "";
        "Contact: City" = $_."City1";
        "Contact: State" = $_."State1";
        "Contact: Zip Code" = $_."Zip1";
        "Contact: Country" = $_."Country1";
        "Contact: Additional Address Information" = "";
        "Contact: Phone" = $_."Telephone1";
        "Contact: Extension" = "";
        "Contact: Alternate Phone" = "";
        "Contact: Mobile Phone" = "";
        "Contact: Fax" = "";
        "Contact: Facebook URL" = "";
        "Contact: Twitter URL" = "";
        "Contact: LinkedIn URL" = "";
        "Contact: Client Portal User Name" = "";
        "Contact: Client Portal Password" = "";
        "Contact: Client Portal Security Level" = "";
        "Contact: Contact Group Name" = "";
        "Contact: New Email Address" = "";
        "Contact: Active/Inactive" = "";
        "Contact: Primary Contact" = "";
        "Contact UDF:29682818 Email List" = "";
    };

    New-Object -TypeName PSObject -Property $AutoTaskProperties
}

main @args