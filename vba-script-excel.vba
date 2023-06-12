Sub oracleGLAccountAnalysisParser()

' This VBA for Excel script parses Oracle's standard
' General Ledger Account Analysis Report for a 360 view
' what's hitting your GLs and a deeper dive into financial
' analysis.
'
' Read more about General Ledger Account Analysis Report at
' https://docs.oracle.com/en/cloud/saas/financials/23b/ocuar/general-ledger-account-analysis-reports.html#s20048874
'
' Alternatively, you may ask your support team to develop a custom report.
'
' GL account analysis standard report parser by voidbydefault
' https://github.com/voidbydefault/
'
' Raise issue on my github page for support



' Add Excel functions to calculate Net (debits-credits),
' parsing GL account number, and GL description

    Range("I21").Select
    ActiveCell.Formula = "=G21-H21"
         
    Range("J21").Select
    ActiveCell.FormulaR1C1 = "=IF(MID(RC[-8],1,2)=""02"",RC[-8],R[-1]C)"
    
    Range("K21").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-8]=""Description"",RC[-7],R[-1]C)"
    
    Range("L21").Select
    ActiveCell.Formula = "=MID(J21,13,5)"

' Freez output of functions

    Range("I21:L21").Select
    Selection.Copy
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.Paste
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
         :=False, Transpose:=False

' Delete Oracle's default header rows

    Rows("1:24").Select
    Selection.Delete Shift:=xlUp

' Add heading/title to the functions added

    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Net"
     
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "GL"
    
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Description"
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "GL Account"

' Convert data range to a table and apply filters to
' work on data cleansing

    ActiveSheet.ListObjects.Add(xlSrcRange, , , xlYes).Name _
         = "tbl_GLAccountData"
    Range("tbl_GLAccountData[#All]").Select
    ActiveSheet.ListObjects("tbl_GLAccountData").Range.AutoFilter Field:=1, Criteria1:= _
         Array("  ", "=", "Source", "Ending Balance for Period", "Beginning Balance for Period", _
     "End of Report", "Account"), Operator:=xlFilterValues
 
' Removes garbage

    Range("tbl_GLAccountData").Select
    Selection.ClearContents
     
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("tbl_GLAccountData").Range.AutoFilter Field:=1
    Range("F12").Select
 
' Formats table

    Cells.Select
    Selection.ClearFormats
    Cells.EntireRow.AutoFit
    ActiveSheet.ListObjects("tbl_GLAccountData").TableStyle = "TableStyleLight1"
     
    Columns("G:I").Select
    Selection.Style = "Comma"
    Columns("C:C").Select
    Selection.NumberFormat = "[$-en-US]dd-mmm-yy;@"
    Cells.Select
    Selection.ColumnWidth = 18.88

' Add new columns to separate GL account number, project and department code
' This may slightly vary depending on your taxomony/coding. Raise a request
' in issues section for support

    Range("L1").Select
    Selection.ListObject.ListColumns.Add Position:=13
    Selection.ListObject.ListColumns.Add Position:=13

' Rename new columns and add formulas

    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Project"
    Range("M2").Select
    ActiveCell.Formula = "=MID([@GL],4,3)"
         
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Department"
    Range("N2").Select
    ActiveCell.Formula = "=MID([@GL],8,4)"

' Complete and final tip

    ' Sum on [Net] column, if the total is zero (i.e. debits = credits),
    ' you're good to go. Thank you for using my script.
    '
    ' My Github: https://github.com/voidbydefault/
    ' My YouTubes:
        ' https://youtube.com/@LivingLinux101
        ' https://youtube.com/@taimurSM

End Sub
