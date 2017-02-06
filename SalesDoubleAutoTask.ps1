. ./ExportToAutoTask.ps1

function Main
{
    param (
        [Parameter(Mandatory = $false)] [string] $InputSalesDoubleExportFilename = ".\Data\SalesDoubleData.xlsx",
        [Parameter(Mandatory = $false)] [string] $InputAutoTaskImportFilename = ".\Data\AutoTaskFromSalesDouble.csv"
    )

    ExportToAutoTask $InputSalesDoubleExportFilename $InputAutoTaskImportFilename
}

function Get-AutoTaskProperties
{
    $AutoTaskProperties = [ordered]@{
        "[required] Account: Name" = $_."Company";
        "Account: Number" = "";
        "Account: Address 1" = $_."Mailing Address1";
        "Account: Address 2" = $_."Mailing Address2";
        "Account: City" = $_."Mailing City";
        "Account: State" = $_."Mailing State";
        "Account: Zip Code" = $_."Mailing Zip";
        "Account: Country" = "";
        "Account: Additional Address Information" = "";
        "[required] Account: Phone" = $_."Phone";
        "Account: Alternate Phone 1" = "";
        "Account: Alternate Phone 2" = "";
        "Account: Fax" = $_."Fax";
        "Account: Web" = $_."Website";
        "Account: Round-Trip Distance" = "";
        "Account: Account Type" = "";
        "Account: Classification" = "B";
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
        "Account UDF:29682816 Lead Source" = $_."Group";
        "Account UDF:29682811 Sales Volume" = "";
        "Account UDF:29682805 Kaseya Customer ID" = "";
        "Site Configuration UDF:29682819 Server Password (s) [protected]" = "";
        "Contact: External ID" = "";
        "Contact: Prefix" = "";
        "[required] Contact: First Name" = $_."Contact First";
        "Contact: Middle Initial" = "";
        "[required] Contact: Last Name" = $_."Contact Last";
        "Contact: Suffix" = "";
        "Contact: Title" = "";
        "[required] Contact: Email Address" = $_."Email";
        "Contact: Email Address 2" = $_."Alt Email";
        "Contact: Email Address 3" = "";
        "Contact: Address 1" = $_."Mailing Address1";
        "Contact: Address 2" = $_."Mailing Address2";
        "Contact: City" = $_."Mailing City";
        "Contact: State" = $_."Mailing State";
        "Contact: Zip Code" = $_."Mailing Zip";
        "Contact: Country" = "";
        "Contact: Additional Address Information" = "";
        "Contact: Phone" = $_."Phone";
        "Contact: Extension" = $_."Phone Ext";
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