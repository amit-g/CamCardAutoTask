. ./ExportToAutoTask.ps1

function Main
{
    param (
        [Parameter(Mandatory = $false)] [string] $InputSalesDoubleExportFilename = ".\Data\SalesDoubleList.xlsx",
        [Parameter(Mandatory = $false)] [string] $InputAutoTaskImportFilename = ".\Data\AutoTaskFromSalesDoubleList.csv"
    )

    ExportToAutoTask $InputSalesDoubleExportFilename $InputAutoTaskImportFilename
}

function Get-AutoTaskProperties
{
    $AutoTaskProperties = [ordered]@{
        "[required] Account: Name" = $_."Company Name";
        "Account: Number" = "";
        "Account: Address 1" = $_."Primary Address 1";
        "Account: Address 2" = "";
        "Account: City" = $_."Primary City";
        "Account: State" = $_."Mailing St";
        "Account: Zip Code" = $_."Zip";
        "Account: Country" = "";
        "Account: Additional Address Information" = "";
        "[required] Account: Phone" = $_."Phone Number";
        "Account: Alternate Phone 1" = "";
        "Account: Alternate Phone 2" = "";
        "Account: Fax" = "";
        "Account: Web" = If ($_."Web Address ".Length -ge 10) { $_."Web Address " } else { "" } ;
        "Account: Round-Trip Distance" = "";
        "Account: Account Type" = "";
        "Account: Classification" = "C";
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
        "Account: SIC Code" = $_."SIC";
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
        "Account UDF:29682812 Number of Users" = $_."Emp";
        "Account UDF:29682815 Number of Servers" = "";
        "Account UDF:29682817 Competitive Contract Expiration Date" = "";
        "Account UDF:29682814 Lead Category" = "";
        "Account UDF:29682816 Lead Source" = "Telemarketing";
        "Account UDF:29682811 Sales Volume" = If ($_."Revenue (US Dollars, million)" -le 500) { $_."Revenue (US Dollars, million)" * 1000000 } else { $_."Revenue (US Dollars, million)" };
        "Account UDF:29682805 Kaseya Customer ID" = "";
        "Site Configuration UDF:29682819 Server Password (s) [protected]" = "";
        "Contact: External ID" = "";
        "Contact: Prefix" = "";
        "[required] Contact: First Name" = $_."First";
        "Contact: Middle Initial" = "";
        "[required] Contact: Last Name" = $_."Last";
        "Contact: Suffix" = "";
        "Contact: Title" = $_."Contact Title";
        "[required] Contact: Email Address" = "TestEmail@example.com";
        "Contact: Email Address 2" = "";
        "Contact: Email Address 3" = "";
        "Contact: Address 1" = "";
        "Contact: Address 2" = "";
        "Contact: City" = "";
        "Contact: State" = "";
        "Contact: Zip Code" = "";
        "Contact: Country" = "";
        "Contact: Additional Address Information" = "";
        "Contact: Phone" = "";
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

    return $AutoTaskProperties
}

main @args