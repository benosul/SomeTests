' AUTHOR: Christian Plattner

Option Explicit

Private Sub cb_exe_Click()
' LAST CHANGE: 29/04/2016
' ALGORITHM
    With Me
        If IsEmpty(.Range("nUser")) Or IsEmpty(.Range("nDate")) Then
            MsgBox "Wichtige Felder wurden beim Öffnen nicht ausgefüllt. Die Werte werden neu initialisiert."
            Call DieseArbeitsmappe.Workbook_Open
            Exit Sub
        End If
        
        Call basTXT.exe_generate(.Range("nConfig").Value, True, .Range("nProcess").Value)
'        Call .exe_create_backup
    End With
        
    Call tabSTRUCTURE.exe
    With Me
        .Range("nBoolean_DUMMY").Value = False
    End With
                
End Sub

Private Sub cb_exe_delete_dummies_Click()
' LAST CHANGE: 05/03/2016
' VARIABLES
    Dim spath As String
    Dim sfile As String
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim b As Byte
    
' ALGORITHM
    On Error Resume Next
    
    b = MsgBox("Sollen die Dummy-Dateien wirklich gelöscht werden?" & Chr(10) & "In einem laufenden Prozess bedeutet dies Nachgenerieren diverser Daten.", vbYesNo)
    
    If b = vbYes Then
        spath = Me.Range("nDummies").Value
        
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(spath)
        
        For Each oFile In oFolder.Files
            sfile = oFile.Name
            If InStr(1, sfile, " - DUMMY.") Then
                basFile.exe_delete (spath & "\" & sfile)
            ElseIf oFile.Size < 3 Then
                basFile.exe_delete (spath & "\" & sfile)
            End If
        Next oFile
    End If
    
    Set oFSO = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
End Sub

Private Sub cb_exe_delete_links_Click()
' LAST CHANGE: 18/03/2016
' VARIABLES
    Dim sP As String
    Dim sF As String
    
' ALGORITHM
    sP = Me.Range("nDesktop").Value
    sF = Dir(sP & "\*.lnk", vbNormal)
    While Not sF = vbNullString
        If Len(sF) > 4 Then _
                Kill sP & "\" & sF
        sF = Dir
    Wend
    sF = Dir(sP & "\*.url", vbNormal)
    While Not sF = vbNullString
        If Len(sF) > 4 Then _
                Kill sP & "\" & sF
        sF = Dir
    Wend
End Sub

Private Sub cb_exe_help_Click()
' LAST CHANGE: 27/02/2016
' ALGORITHM
    Call exe_help("Ich habe vergessen die Bemerkung im PS zu löschen...")
End Sub

Sub exe_help(slog As String)
' LAST CHANGE: 21/12/2016
' VARIABLES
    Dim oOutlook As Object
    
' ALGORITHM
    Set oOutlook = CreateObject("Outlook.Application")
    With oOutlook.CreateItem(0)
        .Subject = ThisWorkbook.FullName & ": " & Me.Range("nUser").Value & "/" & Me.Range("nProcess").Value & "/" & Format(Me.Range("nDate").Value, "YYYY-MM-DD")
        .To = "valentin.buergler@swissgrid.ch"
        .Body = "Lieber Valentin" & Chr(10) & Chr(10) & _
                "Mit '" & ThisWorkbook.Name & "' gibt es ein Problem:" & Chr(10) & _
                "..." & Chr(10) & Chr(10) & _
                "Gruss" & Chr(10) & _
                "Namenslose/r OPler/in" & Chr(10) & Chr(10) & _
                "PS: " & slog
        .Importance = 2
        .Display
    End With
    Set oOutlook = Nothing
End Sub

Private Sub cb_Notprozess_Click()
    Dim PathNotprozess As String

    PathNotprozess = "G:\Aplan\Notprozess\WVP"
    Call Shell(PathNotprozess & "\NotprozessWVP.bat", vbNormalFocus)
End Sub

Private Sub cmd_exit_Click()
' LAST CHANGE: 13.08.2019
' VARIABLES
    Dim oWB As Workbook
    Dim i As Integer

' ALGORITHM
    For Each oWB In Workbooks
        i = i + 1
        If i > 1 Then Exit For
    Next oWB
    
    If i = 1 Then
        Application.DisplayAlerts = False
        Application.Quit
    Else
        ThisWorkbook.Close Savechanges:=False
        Application.DisplayAlerts = True                    'Fragefenster wieder einschalten
        ActiveWorkbook.UpdateLinks = xlUpdateLinksAlways    'Akutalisiere Verknuepfungen
    End If
End Sub

Sub Worksheet_Activate()
' LAST CHANGE: 27/01/2016
' ALGORITHM
    Me.Protect Password:="MachKeiScheiss"
End Sub

' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' BACKUP PROCEDURES
' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Sub exe_create_backup()
' LAST CHANGE: 15/04/2016
' VARIABLES
    Dim spath As String
    Dim sname As String
    Dim sname_old As String
    Dim b As Boolean
    Dim r As Integer
    Dim r_max As Integer
    Dim c As Integer
    Dim c_max As Integer
    Dim srow As String
    Dim scsv As String
    
' ALGORITHM
    spath = ThisWorkbook.Path & "\_Planung_Vorlagen\3_WVP\BACKUP\Struktur-Excel"
    sname = Format(Me.Range("nDate").Value, "YYYYMMDD") & "_Struktur.csv"
    sname_old = Format(Me.Range("nDate").Value - 1, "YYYYMMDD") & "_Struktur.csv"
    
    If Not basFile.info_exists(spath & "\" & sname) Then
        Call basPath.exe_create(spath, 4)
        
        b = Application.ScreenUpdating
        Application.ScreenUpdating = False
        With tabSTRUCTURE
            .Unprotect "MachKeiScheiss"
            r_max = .UsedRange.SpecialCells(xlCellTypeLastCell).Row
            c_max = .UsedRange.SpecialCells(xlCellTypeLastCell).Column
            scsv = Now
            For r = 1 To r_max
                srow = .Cells(r, 1).Value & ";"
                For c = 2 To c_max
                    srow = srow & .Cells(r, c).Value & ";"
                Next c
                scsv = scsv & Chr(10) & srow
            Next r
            .Protect "MachKeiScheiss"
        End With
        Call basTXT.exe_generate(spath & "\" & sname, False, scsv)
        Call basFile.exe_delete(spath & "\" & sname_old)
        Application.ScreenUpdating = b
    End If
End Sub
