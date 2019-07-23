Attribute VB_Name = "Miniature_Macs"
Option Explicit

Sub Tighten_Worksheet()

    Dim ws As Worksheet
    
    Set ws = Application.ActiveWorkbook.ActiveSheet
    
    ws.Cells.Select
    
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

End Sub
