### Powershell scripts to transform exported data (Excel format) from CamCard and SalesDouble to AutoTask import format (CSV).

**Prerequisite:**

ImportExcel module from https://github.com/dfinke/ImportExcel

```Powershell
Install-Module ImportExcel -scope CurrentUser
Install-Module ImportExcel # Requires Elevation (Installs for everyone)
```

**Usage:**

```Powershell
CamCardToAutoTask {.\Data\Personal_Contacts.xlsx} {.\Data\AutoTaskFromCamCard.csv}
SalesDoubleToAutoTask {.\Data\SalesDoubleData.xlsx} {.\Data\AutoTaskFromSalesDouble.csv}
SalesDoubleListToAutoTask {.\Data\SalesDoubleList.xlsx} {.\Data\AutoTaskFromSalesDouble.csv}
```