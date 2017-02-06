### Powershell scripts to transform exported data (Excel format) from CamCard and SalesDouble to AutoTask import format (CSV).

**Prerequisite:**

ImportExcel module from https://github.com/dfinke/ImportExcel

```Powershell
Install-Module ImportExcel -scope CurrentUser
Install-Module ImportExcel # Requires Elevation (Installs for everyone)
```

**Usage:**

```Powershell
CamCardAutoTask {.\Data\Personal_Contacts.xlsx} {.\Data\AutoTaskFromCamCard.csv}
SalesDoubleAutoTask {.\Data\SalesDoubleData.xlsx} {.\Data\AutoTaskFromSalesDouble.csv}
```