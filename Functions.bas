Attribute VB_Name = "Functions"
Function getMiniMacs()

    Dim project As VBProject
    Dim vbComp As VBComponent
    Dim currentMacro As String
    Dim newMacro As String
    Dim x As String
    Dim y As String
    Dim macros As String
    
    On Error Resume Next
    currentMacro = ""
    Documents.Add
    
    For Each project In Application.VBE.VBProjects
    
        For Each vbComp In project.VBComponents
            If Not vbComp Is Nothing Then
                If vbComp.CodeModule = "Miniature_Macs" Then
                    For i = 1 To vbComp.CodeModule.CountOfLines
                        newMacro = vbComp.CodeModule.ProcOfLine(Line:=i, _
                            prockind:=vbext_pk_Proc)
                        If currentMacro <> newMacro Then
                            currentMacro = newMacro
                        
                            If currentMacro <> "" And currentMacro <> "app_NewDocument" Then
                                macros = currentMacro + " " + macros
                            End If
                        End If
                    Next
                End If
            End If
        Next
    
    Next
    
getMiniMacs = macros

End Function

Sub execute()

    Dim AppArray() As String
    macs = getMiniMacs
    Debug.Print macs
    AppArray() = Split(macs, " ")
    
    
    For i = 0 To UBound(AppArray)
        
        temp = AppArray(i)
        If temp <> "" Then
            If temp <> "execute" And temp <> "getMacros" Then
                Application.Run (AppArray(i))
                Sheet1.Range("D" & i + 1).Value = temp
            End If
        End If
    
    Next i

End Sub
