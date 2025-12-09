Sub OriginalRecordedMacro()
'
' Originally Recorded Macro
'
' Keyboard Shortcut: Ctrl+Shift+N
'
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""##############"")"
    Range("B1").Select
    Selection.AutoFill Destination:=Range("B1:B3202")
    Range("B1:B3202").Select
    Columns("B:B").ColumnWidth = 12.09
    Columns("B:B").ColumnWidth = 15.91
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub BarcodeSciToText()
'
' Convert 14 Character Barcode from Scientific Notation to Text
' Expects Barcode in the first column.
'
'
' Keyboard Shortcut: Ctrl+Shift+N
'
'[TODO] Prompt the user for the column with the first as default
'
TextPattern = """##############"""
FirstRow = 1
SecondColumn = 2
BColumnStart = "B:B"
StartCell = "B1"
AColumnStart = "A:A"



    Columns(BColumnStart).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(StartCell).Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1]," & TextPattern & ")"
    Range(StartCell).Select
    lastRow = Range(AColumnStart).SpecialCells(xlCellTypeLastCell).Row
    Selection.AutoFill Destination:=Range(Cells(FirstRow, SecondColumn), Cells(lastRow, SecondColumn))
    Range(Cells(FirstRow, SecondColumn), Cells(lastRow, SecondColumn)).Select
    Columns(BColumnStart).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

