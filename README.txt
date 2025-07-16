Process and track Regence Blue Shield Claims.

Goto regence.com   (user blake5634, pwd RE_tr......)

    Click "Claims" , Download CSV

    >p3 procnewClaims.py  <download file>

    >p3 reformatAsHealthBills.py  EOB_Claims_data.ods

   >lo NewSheetforHealthBills.ods
         * sort by patient and date
         * paste into HealthBills20xx.ods

