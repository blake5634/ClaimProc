Process and track Regence Blue Shield Claims.

## New workflow:

    1. Goto regence.com   (user blake5634, pwd RE_tr......)

    2. Click "Claims" -> "Download CSV"

    (system downloads a .xlsx file(!))


    3. Using command line:

``` >p3 procClaims.py  <download file>

    >p3 reformatAsHealthBills.py  EOB_Claims_data.ods


    >lo NewSheetforHealthBills.ods
```

    4. Paste each patient's rows into
         appropriate tab of HealthBills20xx.ods

