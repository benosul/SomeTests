' AUTHOR: Christian

Option Explicit

Sub exe()
' LAST CHANGE: 07/04/2016
' VARIABLES
    Dim r As Integer
    Dim c_exe As Integer
    Dim c_link As Integer
    Dim c_path As Integer
    Dim c_name As Integer
    Dim c_order As Integer
    Dim slog As String
    Dim s As String
    
' ALGORITHM
    With Me
        With .Range("structure_nexe")
            r = .Row + 1
            c_exe = .Column
        End With
        c_link = .Range("structure_noriginal").Column
        c_path = .Range("structure_npath").Column
        c_name = .Range("structure_nfile").Column
        c_order = .Range("structure_norder").Column
        
        While Not IsEmpty(.Cells(r, c_order))
            If .Cells(r, c_exe).Value = 1 Then
                If .Cells(r, c_order).Value = "copy" Then
                    s = exe_copy(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "copy_all" Then
                    s = exe_copy_all(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "copy_ask" Then
                    s = exe_copy_ask(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "copy_newest" Then
                    s = exe_copy_newest(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "create_file" Then
                    s = exe_create_file(r, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "create_folder" Then
                    s = exe_create_folder(r, c_path)
                ElseIf .Cells(r, c_order).Value = "delete_file" Then
                    s = exe_delete_file(r, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "delete_folder" Then
                    s = exe_delete_folder(r, c_path)
                ElseIf .Cells(r, c_order).Value = "delete_folder_ask" Then
                    s = exe_delete_folder_ask(r, c_path)
                ElseIf .Cells(r, c_order).Value = "lnk" Then
                    s = exe_lnk(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "move" Then
                    s = exe_move(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "move_all" Then
                    s = exe_move_all(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "move_ask" Then
                    s = exe_move_ask(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "move_newest" Then
                    s = exe_move_newest(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "overwrite" Then
                    s = exe_overwrite(r, c_link, c_path, c_name)
                ElseIf .Cells(r, c_order).Value = "pause" Then
                    Application.Wait (Now + TimeSerial(0, 0, 0.5))
                ElseIf .Cells(r, c_order).Value = "url" Then
                    s = exe_url(r, c_link, c_path, c_name)
                End If
                If Not s = vbNullString Then _
                        slog = slog & Chr(10) & s
            End If
            r = r + 1
        Wend
    End With
            
    If Not slog = vbNullString Then
        If InStr(1, slog, "Error:") > 0 Then
            slog = Right(slog, Len(slog) - 1)
            Call tab_GUI.exe_help(slog)
        ElseIf tab_GUI.box_log.Value Then
            slog = "Modus: Log-Datei anzeigen" & Chr(10) & slog
            Call tab_GUI.exe_help(slog)
        End If
    Else
        
    End If
End Sub

Function exe_copy(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' VARIABLES
    Dim sP As String
    Dim sF As String
    Dim sPnF_lnk As String
    
' ALGORITHM
    On Error GoTo Error
    With Me
        sF = .Cells(r, c_name).Value
        sP = .Cells(r, c_path).Value
        sPnF_lnk = .Cells(r, c_link).Value
        
        If InStr(1, sPnF_lnk, "*") Then
            If basPath.info_exists(basFile.info_P(sPnF_lnk)) Then
                sPnF_lnk = basFile.info_P(sPnF_lnk) & "\" & basFile.info_newest(sPnF_lnk)
                If InStr(1, sF, "*") Then _
                        sF = basFile.info_F(sPnF_lnk)
            Else
                exe_copy = "Error: '" & sPnF_lnk & "' exe_copy '" & sP & "\" & sF & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            End If
        End If
        
        If basFile.info_exists(sP & "\" & sF) Then
            exe_copy = "Warning: exe_copy '" & sP & "\" & sF & "' (Datei existiert bereits; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            If basFile.info_size(sPnF_lnk) < 3 Then
                exe_copy = "Warning: exe_copy '" & sPnF_lnk & "' (DUMMY-Datei; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            Else
                If Not basFile.exe_copy(sPnF_lnk, sP & "\" & sF, False) Then _
                        exe_copy = "Error: '" & sPnF_lnk & "' exe_copy '" & sP & "\" & sF & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            End If
        End If
    End With
    Exit Function
Error:
    exe_copy = "Error: '" & Me.Cells(r, c_link).Value & "' exe_copy '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_copy_ask(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' ALGORITHM
    On Error GoTo Error
    With Me
        If basFile.info_size(.Cells(r, c_link).Value) < 3 Then
            exe_copy_ask = "Warning: exe_copy_ask '" & .Cells(r, c_link).Value & "' (DUMMY-Datei; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            If Not basFile.exe_copy(.Cells(r, c_link).Value, .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value, False) Then _
                    exe_copy_ask = "Error: '" & .Cells(r, c_link).Value & "' exe_copy_ask '" & .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        End If
    End With
    Exit Function
Error:
    exe_copy_ask = "Error: '" & Me.Cells(r, c_link).Value & "' exe_copy_ask '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_copy_newest(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' VARIABLES
    Dim spath As String
    Dim sfile As String
    
' ALGORITHM
    On Error GoTo Error
    With Me
        spath = basFile.info_P(.Cells(r, c_link).Value)
        sfile = basFile.info_newest(.Cells(r, c_link).Value)
        If sfile = vbNullString Then _
                exe_copy_newest = "Error: exe_copy_newest '" & .Cells(r, c_link).Value & "' (Datei nicht gefunden; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        If Not basFile.exe_copy(spath & "\" & sfile, .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value, False) Then _
                exe_copy_newest = "Error: '" & spath & "\" & sfile & "' exe_copy_newest '" & .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
    End With
    Exit Function
Error:
    exe_copy_newest = "Error: '" & Me.Cells(r, c_link).Value & "' exe_copy_newest '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_create_file(r As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' ALGORITHM
    On Error GoTo Error
    With Me
        If basFile.info_exists(.Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value) Then
            exe_create_file = "Warning: exe_create_file ' " & .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value & "' (Datei existiert bereits; " & r & ", " & c_path & ", " & c_name & ")"
        Else
            If Not basTXT.exe_generate(.Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value, False) Then _
                    exe_create_file = "Error: exe_create_file '" & .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value & "' (" & r & ", " & c_path & ", " & c_name & ")"
        End If
    End With
    Exit Function
Error:
    exe_create_file = "Error: exe_create_file '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_create_folder(r As Integer, c_path As Integer) As String
' LAST CHANGE: 11/04/2016
' ALGORITHM
    On Error GoTo Error
    With Me
        If Not basPath.exe_create(.Cells(r, c_path).Value) Then _
                exe_create_folder = "Error: exe_create_folder '" & .Cells(r, c_path).Value & "' (" & r & ", " & c_path & ")"
    End With
    Exit Function
Error:
    exe_create_folder = "Error: exe_create_folder '" & Me.Cells(r, c_path).Value & "' (VBA-Fehler: " & r & ", " & c_path & ")"
End Function

Function exe_delete_file(r As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' ALGORITHM
    On Error GoTo Error
    With Me
        If Not basFile.exe_delete(.Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value) Then _
                exe_delete_file = "Warning: exe_delete_file '" & .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value & "' (" & r & ", " & c_path & ", " & c_name & ")"
    End With
    Exit Function
Error:
    exe_delete_file = "Error: exe_delete_file '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_delete_folder(r As Integer, c_path As Integer) As String
' LAST CHANGE: 11/04/2016
' ALGORITHM
    On Error GoTo Error
    With Me
        If Not basPath.exe_delete(.Cells(r, c_path).Value) Then _
                exe_delete_folder = "Warning: exe_delete_folder '" & .Cells(r, c_path).Value & "' (" & r & ", " & c_path & ")"
    End With
    Exit Function
Error:
    exe_delete_folder = "Error: exe_delete_folder '" & Me.Cells(r, c_path).Value & "' (VBA-Fehler: " & r & ", " & c_path & ")"
    MsgBox "'" & Me.Cells(r, c_path).Value & "' konnte nicht gelöscht werden. Es benutzt jemand den Ordner." & Chr(10) & _
            "Aufgrund dieser Fehlermeldung muss kein Mail versendet werden."
End Function

Function exe_delete_folder_ask(r As Integer, c_path As Integer) As String
' LAST CHANGE: 11/04/2016
' VARIBLES
    Dim sP As String
    Dim b As Byte
    
' ALGORITHM
    On Error GoTo Error
    With Me
        sP = .Cells(r, c_path).Value
    End With
    If basPath.info_exists(sP) Then
        b = MsgBox("Soll der Ordner '" & sP & "' wirklich gelöscht werden?", vbYesNo)
        If b = vbYes Then
            If Not basPath.exe_delete(sP) Then _
                    exe_delete_folder_ask = "Error: exe_delete_folder_ask'" & sP & "' (" & r & ", " & c_path & ")"
        End If
    End If
    Exit Function
Error:
    exe_delete_folder_ask = "Error: exe_delete_folder_ask '" & Me.Cells(r, c_path).Value & "' (VBA-Fehler: " & r & ", " & c_path & ")"
    MsgBox "'" & Me.Cells(r, c_path).Value & "' konnte nicht gelöscht werden. Es benutzt jemand den Ordner." & Chr(10) & _
            "Aufgrund dieser Fehlermeldung muss kein Mail versendet werden."
End Function

Function exe_lnk(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 18/07/2016
' VARIABLES
    Dim sPnF_old As String
    Dim spath As String
    Dim sfilename As String
    
' ALGORITHM
    On Error GoTo Error
    With Me
        sPnF_old = .Cells(r, c_link).Value
        spath = .Cells(r, c_path).Value
        sfilename = .Cells(r, c_name).Value
        
        If Not basFile.info_exists(sPnF_old) Then _
                exe_lnk = "Warning: '" & sPnF_old & "' exe_lnk '" & spath & "', '" & sfilename & "' (Verlinkte Datei existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        
        If Not basPath.info_exists(spath) Then
            exe_lnk = "Error: '" & sPnF_old & "' exe_lnk '" & spath & "', '" & sfilename & "' (Zielordner '" & spath & "' existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            If Not basFile.exe_link(sPnF_old, spath, sfilename) Then _
                    exe_lnk = "Error: '" & sPnF_old & "' exe_lnk '" & spath & "', '" & sfilename & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        End If
    End With
    Exit Function
Error:
    exe_lnk = "Error: '" & Me.Cells(r, c_link).Value & "' exe_lnk '" & Me.Cells(r, c_path).Value & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_move(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' VARIABLES
    Dim soldPnF As String
    Dim spath As String
    Dim sfilename As String
    
' ALGORITHM
    On Error GoTo Error
    With Me
        soldPnF = .Cells(r, c_link).Value
        spath = .Cells(r, c_path).Value
        sfilename = .Cells(r, c_name).Value
        
        If InStr(1, soldPnF, "*") Then
            If basPath.info_exists(basFile.info_P(soldPnF)) Then
                soldPnF = basFile.info_P(soldPnF) & "\" & basFile.info_newest(soldPnF)
                If InStr(1, sfilename, "*") Then _
                        sfilename = basFile.info_F(soldPnF)
            Else
                exe_move = "Error: '" & soldPnF & "' exe_move '" & spath & "\" & sfilename & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            End If
        End If
        
        If basFile.info_size(soldPnF) = 0 Then
            exe_move = "Warning: exe_move '" & .Cells(r, c_link).Value & "' (Ursprungsdatei existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        ElseIf basFile.info_size(soldPnF) < 3 Then
            exe_move = "Warning: exe_move '" & .Cells(r, c_link).Value & "' (DUMMY-Datei; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            If basPath.info_exists(spath) And basFile.info_exists(soldPnF) Then
                If basFile.exe_copy(soldPnF, spath & "\" & sfilename, True) Then
                    If basFile.info_exists(spath & "\" & sfilename) Then
                        If Not basFile.exe_delete(soldPnF) Then _
                                exe_move = "Warning: exe_move '" & soldPnF & "' (Ursprungsdatei löschen fehlgeschlagen; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                    Else
                        exe_move = "Error: exe_move '" & spath & "\" & sfilename & "' (Zieldatei kopieren fehlgeschlagen; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                    End If
                Else
                    exe_move = "Error: exe_move '" & spath & "\" & sfilename & "' (Zieldatei kopieren fehlgeschlagen; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                End If
            Else
                If Not basPath.info_exists(spath) Then
                    exe_move = "Error: exe_move '" & spath & "' (Zielpfad existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                Else
                    exe_move = "Warning: exe_move '" & soldPnF & "' (Ursprungsdatei existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                End If
            End If
        End If
    End With
    Exit Function
Error:
    exe_move = "Error: '" & Me.Cells(r, c_link).Value & "' exe_move '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
    MsgBox "'" & Me.Cells(r, c_link).Value & "' konnte nicht gelöscht werden. Es benutzt jemand die Datei." & Chr(10) & _
            "Aufgrund dieser Fehlermeldung muss kein Mail versendet werden."
End Function

Function exe_move_all(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 01/04/2020
' VARIABLES
    Dim soldPnF As String
    Dim spath As String
    Dim sfilename As String
    Dim s As String
    Dim Check As Boolean

' ALGORITHM
    On Error GoTo Error
    With Me
        
        Check = True
        Do
            soldPnF = .Cells(r, c_link).Value
            spath = .Cells(r, c_path).Value
            sfilename = .Cells(r, c_name).Value
            
            If InStr(1, soldPnF, "*") Then
                If basPath.info_exists(basFile.info_P(soldPnF)) Then
                    s = exe_move(r, c_link, c_path, c_name)
                Else
                    Check = False
                End If
            Else
                Check = False
            End If
        Loop Until Check = False
    End With
    Exit Function
Error:
    exe_move_all = "Error: '" & Me.Cells(r, c_link).Value & "' exe_move '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
    MsgBox "'" & Me.Cells(r, c_link).Value & "' konnte nicht gelöscht werden. Es benutzt jemand die Datei." & Chr(10) & _
            "Aufgrund dieser Fehlermeldung muss kein Mail versendet werden."
End Function

Function getPath(pf) As String: getPath = Left(pf, InStrRev(pf, "\")): End Function

' **** Funktioniert (noch) nicht ***********************************************
Function exe_copy_all(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 01/04/2020
' VARIABLES
    Dim soldPnF As String
    Dim spath As String
    Dim sfilename As String
    Dim sP As String
    Dim sF As String
    Dim sPnF_lnk As String
    Dim sourcePath
    Dim s As String
    Dim Check As Boolean

    Check = True

' ALGORITHM
    On Error GoTo Error
    With Me
    
    Dim copyFiles As Variant
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    Dim oFile As Object
    Dim i As Integer
    
    sPnF_lnk = getPath(.Cells(r, c_link).Value)
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPnF_lnk)
    Set oFiles = oFolder.Files
    
    'ReDim copyFiles(1 To oFiles.Count)
    For Each oFile In oFolder.Files
        ' copyFiles(i) = oFile.Name
        sF = oFile.Name
        'sF = .Cells(r, c_name).Value
        sP = .Cells(r, c_path).Value
        'sPnF_lnk = .Cells(r, c_link).Value
        
        If InStr(1, sPnF_lnk, "*") Then
            If basPath.info_exists(basFile.info_P(sPnF_lnk)) Then
                sPnF_lnk = basFile.info_P(sPnF_lnk) & "\" & basFile.info_newest(sPnF_lnk)
                If InStr(1, sF, "*") Then _
                        sF = basFile.info_F(sPnF_lnk)
            Else
                exe_copy_all = "Error: '" & sPnF_lnk & "' exe_copy '" & sP & "\" & sF & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            End If
        End If
        
        If basFile.info_exists(sP & "\" & sF) Then
            exe_copy_all = "Warning: exe_copy '" & sP & "\" & sF & "' (Datei existiert bereits; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            If basFile.info_size(sPnF_lnk) < 3 Then
                exe_copy_all = "Warning: exe_copy '" & sPnF_lnk & "' (DUMMY-Datei; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            Else
                If Not basFile.exe_copy(sPnF_lnk, sP & "\" & sF, False) Then _
                        exe_copy_all = "Error: '" & sPnF_lnk & "' exe_copy '" & sP & "\" & sF & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            End If
        End If

    Next oFile
    
    End With
    Exit Function
Error:
    exe_copy_all = "Error: '" & Me.Cells(r, c_link).Value & "' exe_copy '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_move_ask(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' VARIABLES
    Dim soldPnF As String
    Dim spath As String
    Dim sfilename As String
    Dim b As Byte
    
' ALGORITHM
    On Error GoTo Error
    With Me
        soldPnF = .Cells(r, c_link).Value
        If basFile.info_size(soldPnF) = 0 Then
            exe_move = "Warning: exe_move '" & .Cells(r, c_link).Value & "' (Ursprungsdatei existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        ElseIf basFile.info_size(soldPnF) < 3 Then
            exe_move_ask = "Warning: exe_move_ask '" & .Cells(r, c_link).Value & "' (DUMMY-Datei; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            spath = .Cells(r, c_path).Value
            sfilename = .Cells(r, c_name).Value
            If basPath.info_exists(spath) And basFile.info_exists(soldPnF) Then
                If basFile.info_exists(spath & "\" & sfilename) Then
                    b = MsgBox("Soll die Datei '" & spath & "\" & sfilename & "' wirklich überschrieben werden?", vbYesNo)
                    If b = vbNo Then
                        exe_move_ask = "Warning: exe_move_ask '" & spath & "\" & sfilename & "' (Zieldatei nicht überschrieben; Benutzerentscheidung; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                        Exit Function
                    End If
                End If
                If basFile.exe_copy(soldPnF, spath & "\" & sfilename, True) Then
                    If basFile.info_exists(spath & "\" & sfilename) Then
                        If Not basFile.exe_delete(soldPnF) Then _
                                exe_move_ask = "Warning: exe_move_ask '" & soldPnF & "' (Ursprungsdatei löschen fehlgeschlagen; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                    Else
                        exe_move_ask = "Error: exe_move_ask '" & spath & "\" & sfilename & "' (Zieldatei kopieren fehlgeschlagen; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                    End If
                Else
                    exe_move_ask = "Error: exe_move_ask '" & spath & "\" & sfilename & "' (Zieldatei kopieren fehlgeschlagen; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                End If
            Else
                If Not basPath.info_exists(spath) Then
                    exe_move = "Error: exe_move_ask '" & spath & "' (Zielpfad existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                Else
                    exe_move = "Warning: exe_move_ask '" & soldPnF & "' (Ursprungsdatei existiert nicht; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
                End If
            End If
        End If
    End With
    Exit Function
Error:
    exe_move_ask = "Error: '" & Me.Cells(r, c_link).Value & "' exe_move_ask '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
    MsgBox "'" & Me.Cells(r, c_link).Value & "' konnte nicht gelöscht werden. Es benutzt jemand die Datei." & Chr(10) & _
            "Aufgrund dieser Fehlermeldung muss kein Mail versendet werden."
End Function

Function exe_move_newest(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' VARIABLES
    Dim spath As String
    Dim sfile As String
    
' ALGORITHM
    On Error GoTo Error
    With Me
        spath = basFile.info_P(.Cells(r, c_link).Value)
        sfile = basFile.info_newest(.Cells(r, c_link).Value)
        If sfile = vbNullString Then _
                exe_move_newest = "Error: exe_move_newest '" & .Cells(r, c_link).Value & "' (Datei nicht gefunden; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        If Not basFile.exe_copy(spath & "\" & sfile, .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value, False) Then
            exe_move_newest = "Error: '" & spath & "\" & sfile & "' exe_move_newest '" & .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            If Not basFile.exe_delete(spath & "\" & sfile) Then _
                    exe_move_newest = "Warning: exe_move_newest '" & spath & "\" & sfile & "' (Ursprungsdatei löschen fehlgeschlagen; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        End If
    End With
    Exit Function
Error:
    exe_move_newest = "Error: '" & Me.Cells(r, c_link).Value & "' exe_move_newest '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
End Function

Function exe_overwrite(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 11/04/2016
' VARIABLES
    Dim sP As String
    Dim sF As String
    Dim sPnF_lnk As String
    
' ALGORITHM
    On Error GoTo Error
    With Me
        sF = .Cells(r, c_name).Value
        sP = .Cells(r, c_path).Value
        sPnF_lnk = .Cells(r, c_link).Value
        
        If InStr(1, sPnF_lnk, "*") Then
            If basPath.info_exists(basFile.info_P(sPnF_lnk)) Then
                sPnF_lnk = basFile.info_P(sPnF_lnk) & "\" & basFile.info_newest(sPnF_lnk)
                If InStr(1, sF, "*") Then _
                        sF = basFile.info_F(sPnF_lnk)
            Else
                exe_overwrite = "Error: '" & sPnF_lnk & "' exe_overwrite '" & sP & "\" & sF & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
            End If
        End If
        
        If basFile.info_size(sPnF_lnk) < 3 Then
            exe_overwrite = "Warning: exe_overwrite '" & sPnF_lnk & "' (DUMMY-Datei; " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        Else
            If Not basFile.exe_copy(sPnF_lnk, sP & "\" & sF, True) Then _
                    exe_overwrite = "Error: '" & sPnF_lnk & "' exe_overwrite '" & sP & "\" & sF & "' (" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
        End If
    End With
    Exit Function
Error:
    exe_overwrite = "Error: '" & Me.Cells(r, c_link).Value & "' exe_overwrite '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "' (VBA-Fehler: " & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
    MsgBox "'" & Me.Cells(r, c_link).Value & "' konnte nicht gelöscht werden. Es benutzt jemand die Datei." & Chr(10) & _
            "Aufgrund dieser Fehlermeldung muss kein Mail versendet werden."
End Function

Function exe_url(r As Integer, c_link As Integer, c_path As Integer, c_name As Integer) As String
' LAST CHANGE: 27/02/2016
' VARIABLES
    Dim ifile As Integer
    
' ALGORITHM
    On Error GoTo Error
    ifile = FreeFile
    With Me
        Open .Cells(r, c_path).Value & "\" & .Cells(r, c_name).Value & ".url" For Output As #ifile
        Print #ifile, "[InternetShortcut]"
        Print #ifile, "URL=" & .Cells(r, c_link).Value
        Close #ifile
    End With
    Exit Function
Error:
    exe_url = "Error: '" & Me.Cells(r, c_link).Value & "' exe_url '" & Me.Cells(r, c_path).Value & "\" & Me.Cells(r, c_name).Value & "(" & r & ", " & c_link & ", " & c_path & ", " & c_name & ")"
End Function

Sub Worksheet_Activate()
' LAST CHANGE: 27/01/2016
' ALGORITHM
    Me.Protect Password:="MachKeiScheiss"
End Sub
