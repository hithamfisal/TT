# Excel Opening Month Dropdown Guide

`=OpeningMonthList` in the Data Validation source box is not a normal worksheet formula. It is a **named range**. Excel is being told to use the cells assigned to the name `OpeningMonthList` as the dropdown list source. You can check or edit it from **Formulas → Name Manager**.

The reliable setup is to create helper columns in the source ticket table first. In `Tickets_Data`, add an `Opening Month Key` column based on the TT opening date or Observation Date. If your date is in `[@[Observation Date]]`, use `=IF([@[Observation Date]]="","",TEXT([@[Observation Date]],"yyyy-mm"))`. Add a display column named `Opening Month` using `=IF([@[Observation Date]]="","",TEXT([@[Observation Date]],"mmm yyyy"))`. The `yyyy-mm` key is best for filtering and sorting because it stays chronological.

Create the dropdown list source on a helper area, ideally a hidden sheet such as `Lists`. In `Lists!A1`, write `All`. In `Lists!A2`, use `=SORT(UNIQUE(FILTER(Tickets_Data[Opening Month Key],Tickets_Data[Opening Month Key]<>"")))`. Then define the name `OpeningMonthList` as the spilled range, for example `=Lists!$A$1#` if `All` and the dynamic list spill from A1, or use a separate combined formula if needed. The dashboard dropdown cell can then use **Data Validation → List → Source: `=OpeningMonthList`**.

To make the dropdown affect the Dashboard, choose one Dashboard cell as the selected month, for example `Dashboard!B2`. Every KPI, chart source, and report formula must include a condition that says: if `Dashboard!B2="All"`, include every row; otherwise include only rows where `Tickets_Data[Opening Month Key]=Dashboard!B2`.

For a simple count of rows, use `=COUNTIFS(Tickets_Data[Opening Month Key],IF($B$2="All","*",$B$2))`. For unique TT count, use `=ROWS(UNIQUE(FILTER(Tickets_Data[TT],(Tickets_Data[TT]<>"")*IF($B$2="All",1,Tickets_Data[Opening Month Key]=$B$2))))`. For a filtered distinct TT report, use `=UNIQUE(FILTER(Tickets_Data[[TT]:[Comments-Feedback]],IF($B$2="All",Tickets_Data[TT]<>"",Tickets_Data[Opening Month Key]=$B$2)))`, adjusting the selected columns to match your report order.

If you are using PivotTables and PivotCharts instead of formulas, the better approach is to add `Opening Month` or `Opening Month Key` to the source table, refresh the PivotTable, then insert a real PivotTable slicer from **PivotTable Analyze → Insert Slicer**. A Data Validation dropdown does not automatically control PivotTables unless you add formulas, VBA, or rebuild the dashboard around formula-driven summary ranges.
