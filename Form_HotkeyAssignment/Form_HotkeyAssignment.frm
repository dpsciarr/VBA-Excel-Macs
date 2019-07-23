VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_HotkeyAssignment 
   Caption         =   "Macro Hotkey Assignment"
   ClientHeight    =   3855
   ClientLeft      =   2040
   ClientTop       =   2385
   ClientWidth     =   10020
   OleObjectBlob   =   "Form_HotkeyAssignment.frx":0000
End
Attribute VB_Name = "Form_HotkeyAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Assign Button
Private Sub CommandButton1_Click()
    Dim ctrl As Boolean
    Dim shift As Boolean
    Dim letter As String
    Dim hotkey As String
    Dim appendage As String
    
    If CheckBox1.Value <> True And CheckBox2.Value <> True And CheckBox3.Value <> True Then
        MsgBox "Choose (CTRL, SHIFT, ALT) for your hotkey configuration."
        GoTo jump
    End If
    
    appendage = ""
    
    If CheckBox1.Value Then
        appendage = appendage & "^"
    End If
    
    If CheckBox2.Value Then
        appendage = appendage & "+"
    End If
    
    If CheckBox3.Value Then
        appendage = appendage & "%"
    End If
    
    hotkey = appendage & "{" & Me.ComboBox1.Value & "}"
    
    Dim listCount As Integer
    listCount = Me.ListBox1.listCount
    
    Dim macro As String
    
    For i = 0 To listCount - 1
        If Me.ListBox1.Selected(i) = True Then
            macro = Me.ListBox1.List(i)
        End If
    Next i
    
    Application.OnKey hotkey, macro
    
    Me.Label3.Caption = macro & " was assigned hotkey " & hotkey
jump:
    
End Sub

'Remove Button
Private Sub CommandButton2_Click()
    Dim appendage As String
    Dim hotkey As String
    
    appendage = ""
    If CheckBox1.Value <> True And CheckBox2.Value <> True And CheckBox3.Value <> True Then
        MsgBox "Choose (CTRL, SHIFT, ALT) for removing your hotkey configuration."
        GoTo jump
    End If
    
    If CheckBox1.Value Then
        appendage = appendage & "^"
    End If
    
    If CheckBox2.Value Then
        appendage = appendage & "+"
    End If
    
    If CheckBox3.Value Then
        appendage = appendage & "%"
    End If
    
    hotkey = appendage & "{" & Me.ComboBox1.Value & "}"
    
    Application.OnKey hotkey
    
    Me.Label3.Caption = hotkey & " has been cleared."
    
jump:
    
End Sub

'Close Button
Private Sub CommandButton3_Click()
    Unload Me
End Sub


Private Sub ListBox1_Change()
    Dim listCount As Integer
    listCount = Me.ListBox1.listCount
    
    For i = 0 To listCount - 1
        If Me.ListBox1.Selected(i) = True Then
            Debug.Print Me.ListBox1.List(i)
        End If
    Next i
    
End Sub


Private Sub UserForm_Initialize()
    Dim macroStr As String
    Dim macroArray() As String
        
    macroStr = Functions.getMiniMacs
    macroArray() = Split(macroStr, " ")
    
    For i = 0 To UBound(macroArray)
        temp = macroArray(i)
        If temp <> "" Then
            ListBox1.AddItem temp
        End If
    Next i
    
    
    
    
    Dim startLetter As String
    startLetter = "A"
    
    Dim ascChar As Integer
    ascChar = Asc(startLetter)
    
    Dim letter As String * 1
    
    For i = 1 To 26
        If ascChar > 90 Then ascChar = 64 + ascChar - 90
        letter = Chr(ascChar)
        Me.ComboBox1.AddItem letter
        
        ascChar = ascChar + 1
    Next i
    
    Me.ComboBox1.AddItem "BACKSPACE"
    Me.ComboBox1.AddItem "CAPSLOCK"
    Me.ComboBox1.AddItem "DEL"
    Me.ComboBox1.AddItem "DOWN"
    Me.ComboBox1.AddItem "END"
    Me.ComboBox1.AddItem "ENTER"
    Me.ComboBox1.AddItem "ESC"
    Me.ComboBox1.AddItem "HOME"
    Me.ComboBox1.AddItem "INSERT"
    Me.ComboBox1.AddItem "LEFT"
    Me.ComboBox1.AddItem "NUMLOCK"
    Me.ComboBox1.AddItem "PGDN"
    Me.ComboBox1.AddItem "PGUP"
    Me.ComboBox1.AddItem "RETURN"
    Me.ComboBox1.AddItem "RIGHT"
    Me.ComboBox1.AddItem "SCROLLLOCK"
    Me.ComboBox1.AddItem "TAB"
    Me.ComboBox1.AddItem "UP"
    
    For i = 1 To 15
        If i <> 8 Then Me.ComboBox1.AddItem "F" + CStr(i)
    Next i
    
    Me.ComboBox1.ListIndex = 0
    
End Sub

