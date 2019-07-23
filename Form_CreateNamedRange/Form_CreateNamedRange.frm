VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_CreateNamedRange 
   Caption         =   "Create Named Range"
   ClientHeight    =   2760
   ClientLeft      =   2040
   ClientTop       =   2385
   ClientWidth     =   8505
   OleObjectBlob   =   "Form_CreateNamedRange.frx":0000
End
Attribute VB_Name = "Form_CreateNamedRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim n_rng As Range
    
    Set wb = Application.ActiveWorkbook
    Set ws = wb.ActiveSheet
    
    Dim rng_name As String
    rng_name = Me.TextBox1.Value
    
    If rng_name <> "" Then
        If OptionButton1 = True Then
            wb.Names.Add Name:=rng_name, RefersTo:=Selection
            Unload Me
        ElseIf OptionButton2 = True Then
            ws.Names.Add Name:=rng_name, RefersTo:=Selection
            Unload Me
        Else
            MsgBox ("Choose a scope for your Named Range.")
        End If
    Else
        MsgBox ("Choose a name for your Named Range.")
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim n_rng As Range

    Set wb = Application.ActiveWorkbook
    Set ws = wb.ActiveSheet
    Set n_rng = Selection

    Me.Label8.Caption = wb.Name
    Me.Label7.Caption = ws.Name
    
    Dim rc As Integer
    rc = n_rng.Rows.Count
    Dim cc As Integer
    cc = n_rng.Columns.Count
    
    Dim topLeft As String
    topLeft = Cells(Selection.Row, Selection.Column).Address
    
    Dim botRight As String
    botRight = Cells(Selection.Rows.Count + Selection.Row - 1, Selection.Columns.Count + Selection.Column - 1).Address
    
    Dim result As String
    result = topLeft + ":" + botRight
    
    Me.Label4.Caption = result
End Sub
