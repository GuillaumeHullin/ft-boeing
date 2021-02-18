Option Explicit

Const R_SheetName As String = "Results"

Sub Updatesheets()
    Dim rowEnd As Long, row As Long, match As Long
    Dim A_WB As Workbook, B_WB As Workbook, R_WB As Workbook
    Dim A_WS As Worksheet, B_WS As Worksheet, R_WS As Worksheet
    Dim A_colRef As Integer, A_colLnk As Integer, B_colRef As Integer, B_colLnk As Integer, R_Row As Integer
    Dim A_Path As String, B_Path As String, A_SheetName As String, B_SheetName As String
    Dim A_arr As Variant, B_arr As Variant
    
    
    Set R_Row = 2

    Set A_Path = "FULL PATH TO THE XLSX A"
    Set B_Path = "FULL PATH TO THE XLSX B"

    Set A_colRef = 3 
    Set A_colLnk = 6   
    Set B_colRef = 3 
    Set B_colLnk = 6
    Set A_SheetName = "Feuil1"
    Set B_SheetName = "Feuil1"

    Set R_WB = ThisWorkbook
    Set R_WS = C_WB.Sheets(R_SheetName)
    

    Workbooks.Open Filename:=A_Path
    Set A_WB = ActiveWorkbook
    Set A_WS = A_WB.Sheets(A_SheetName)

    Workbooks.Open Filename:=B_Path
    Set B_WB = ActiveWorkbook
    Set B_WS = B_WB.Sheets(B_SheetName)

    '-- Store Version A items into array --
    rowEnd = A_WS.Cells(Rows.Count, A_colRef).End(xlUp).Row
    A_arr = A_WS.Range(Cells(1, A_colRef).Address, Cells(rowEnd, A_colRef).Address).Value
                        
    '-- Store Version B items into array --
    rowEnd = B_WS.Cells(Rows.Count, B_colRef).End(xlUp).Row
    B_arr = B_WS.Range(Cells(1, B_colRef).Address, Cells(rowEnd, B_colRef).Address).Value

    R_WS.Cells.ClearContents 'EEEEERRRRRRAAAAAAASSSSSEEEEEEE EVERYTHING!!!!!! MOUHAHAHAHAHAHA

    ' Pause screen update  
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Lets put some headers
    R_WS.Cells(1,1) = "Workbook_A"
    R_WS.Cells(1,2) = "Ref_A"
    R_WS.Cells(1,3) = "Lnk_A"
    R_WS.Cells(1,4) = "Lnk_B"
    R_WS.Cells(1,5) = "Ref_B"
    R_WS.Cells(1,6) = "Workbook_B"

    For row = 2 To UBound(A_arr, 1)
        match = 0
        On Error Resume Next
        match = WorksheetFunction.Match(B_arr(row, 1), A_arr, 0)
        On Error GoTo 0
        
        R_WS.Cells(R_Row,1) = "Prout A"
        R_WS.Cells(R_Row,2) = A_WS.Cells(match, A_colRef).Value
        R_WS.Cells(R_Row,3) = A_WS.Cells(match, A_colLnk).Value

        If match <> 0 Then
            ' B_WS.Cells(row, B_colLnk).Value = A_WS.Cells(match, A_colLnk).Value
            ' MsgBox "Ref A " & A_WS.Cells(match, A_colRef).Value & "Ref B" &  & "Lnk A " & A_WS.Cells(match, A_colLnk).Value & "Lnk B" & B_WS.Cells(row, B_colLnk).Value
            R_WS.Cells(R_Row,4) = B_WS.Cells(row, B_colLnk).Value
            R_WS.Cells(R_Row,5) = B_WS.Cells(row, B_colRef).Value

        End If

        R_WS.Cells(R_Row,6) = "Prout B"

        R_Row = R_Row + 1

    Next row

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With


End Sub

Private Function OpenWB(ByVal Title As String) As Workbook
    Dim vWBName As Variant

    vWBName = Application.GetOpenFilename(filefilter:="Excel files (*.xlsx),*.xlsx", Title:=Title)
    If vWBName = False Then
        Set OpenWB = Nothing
        Exit Function
    End If

    Workbooks.Open Filename:=vWBName
    Set OpenWB = ActiveWorkbook
End Function

'                       /^--^\     /^--^\     /^--^\
'                       \____/     \____/     \____/
'                      /      \   /      \   /      \
'                     |        | |        | |        |
'                      \__  __/   \__  __/   \__  __/
' |^|^|^|^|^|^|^|^|^|^|^|^\ \^|^|^|^/ /^|^|^|^|^\ \^|^|^|^|^|^|^|^|^|^|^|^|
' | | | | | | | | | | | | |\ \| | |/ /| | | | | | \ \ | | | | | | | | | | |
' | | | | | | | | | | | | / / | | |\ \| | | | | |/ /| | | | | | | | | | | |
' | | | | | | | | | | | | \/| | | | \/| | | | | |\/ | | | | | | | | | | | |
' #########################################################################
' | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | |
' | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | |