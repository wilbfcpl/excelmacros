'------------------------------------------------------------------------------
' <file>        circ_stats_report14.bas
' <author>      Your Name
' <created>     YYYY-MM-DD
' <lastupdate>  YYYY-MM-DD
' <version>     1.0
' <description> Brief description of what this module/class does.
'
' <usage>       Optional: Example usage or notes for developers.
'
' <notes>       Optional: Special considerations, dependencies, or warnings.
'
' <copyright>   © YYYY Your Company/Organization. All rights reserved.
'---------------

Attribute VB_Name = "Module2"
Option Explicit

'--------------------------------------------
' Subroutine Name : PP_Text_Bounce()
' Purpose         : 
' Parameters      :
'   
'   
' Returns         : None (Subroutines do not return values)


Sub PP_Text_Bounce()
Attribute PP_Text_Bounce.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' PP_Text_Bounce Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long

    ' Set the worksheet where duplicates need to be removed
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

    ' Find the last row in column A (or adjust to your data range)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Define the range to remove duplicates from (e.g., A1 to the last row in column A)
    Set rng = ws.Range("A1:C" & lastRow) ' Adjust columns (A:C) as needed

    ' Remove duplicates based on all columns in the range
    On Error Resume Next ' Handle errors gracefully
    rng.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes ' Adjust columns and header option
    On Error GoTo 0

    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=INT(RIGHT(RC[-2], 10))"
    Range("E3").Select
    Columns("E:E").EntireColumn.AutoFit
    Columns("E:E").ColumnWidth = 22.3
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=INT(RIGHT(RC[-1], 10))"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E23")
    Range("E2:E23").Select
    Selection.NumberFormat = "###-###-####"
End Sub
