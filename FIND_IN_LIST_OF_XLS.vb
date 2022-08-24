Option Explicit
Option Compare Text

Sub loop_books()
    Dim wbfiles As Workbook
    On Error Resume Next
    Set wbfiles = Workbooks("W:\2022\all_excels.xlsx")
    On Error GoTo 0
    If wbfiles Is Nothing Then Set wbfiles = Workbooks.Open("W:\2022\all_excels.xlsx", , True)
    
    Dim r As Range, rbooks As Range
    Set rbooks = wbfiles.Sheets(1).Range("B2")
    Set rbooks = Range(rbooks, rbooks.End(xlDown))
    Dim i As Long, j As Long
    i = 0
    j = 0
    For Each r In rbooks
        r.Parent.Activate
        r.Select
        j = j + 1
        If j Mod 10 = 0 Then
            Application.ScreenUpdating = True
            r.Select
            DoEvents
            Application.ScreenUpdating = False
        End If
        If r.Offset(0, 1) = "" Then
            r.Offset(0, 1) = look_for_vba(r.Value)
            i = i + 1
            If i Mod 50 = 0 Then
                Application.ScreenUpdating = True
                r.Select
                wbfiles.Save
                DoEvents
                Application.ScreenUpdating = False
            End If
            
        End If
    Next r
End Sub


Function look_for_vba(sBookname As String)
    Dim wb As Workbook
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error Resume Next
    Set wb = Workbooks.Open(sBookname, False, False, True)
    If Err.Number <> 0 Then
        look_for_vba = Err.Description
        Exit Function
    End If
    On Error GoTo 0
    Debug.Print sBookname
    wb.Application.Calculation = xlCalculationManual
    wb.Application.EnableEvents = False
    Dim sRes As String
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim FindWhat As String
    Dim SL As Long ' start line
    Dim EL As Long ' end line
    Dim SC As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean
    On Error Resume Next
    Set VBProj = wb.VBProject
    On Error GoTo 0
    If VBProj Is Nothing Then
        look_for_vba = "No VBA_PROJECT"
        wb.Close False
        Exit Function
    End If
    FindWhat = ThisWorkbook.Sheets(1).Range("C3").Value
    
    For Each VBComp In VBProj.VBComponents
        Set CodeMod = VBComp.CodeModule
        With CodeMod
            SL = 1
            EL = .CountOfLines
            SC = 1
            EC = 255
            Found = .Find(target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
                EndLine:=EL, EndColumn:=EC, _
                wholeword:=False, MatchCase:=False, patternsearch:=False)
            Do Until Found = False
                sRes = sRes & "Found " & FindWhat & " in module " & CodeMod & " at  Line: " & CStr(SL) & " Column: " & CStr(SC) & vbNewLine
                EL = .CountOfLines
                SC = EC + 1
                EC = 2550
                Found = .Find(target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
                    EndLine:=EL, EndColumn:=EC, _
                    wholeword:=True, MatchCase:=False, patternsearch:=False)
            Loop
        End With
    Next
   
    wb.Close False
    Application.EnableEvents = True
    If sRes = "" Then sRes = "Search term not found"
    look_for_vba = sRes

End Function


Sub SearchCodeModule()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim FindWhat As String
        Dim SL As Long ' start line
        Dim EL As Long ' end line
        Dim SC As Long ' start column
        Dim EC As Long ' end column
        Dim Found As Boolean
        
        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents("Module1")
        Set CodeMod = VBComp.CodeModule
        
        FindWhat = "sun"
        
        With CodeMod
            SL = 1
            EL = .CountOfLines
            SC = 1
            EC = 2550
            Found = .Find(target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
                EndLine:=EL, EndColumn:=EC, _
                wholeword:=True, MatchCase:=False, patternsearch:=False)
            Do Until Found = False
                Debug.Print "Found at: Line: " & CStr(SL) & " Column: " & CStr(SC)
                EL = .CountOfLines
                SC = EC + 1
                EC = 255
                Found = .Find(target:=FindWhat, StartLine:=SL, StartColumn:=SC, _
                    EndLine:=EL, EndColumn:=EC, _
                    wholeword:=True, MatchCase:=False, patternsearch:=False)
            Loop
        End With
    End Sub



