Option Explicit

Const Result_SheetName As String = "Results"

Sub Updatesheets()
    Dim rowEnd As Long, row As Long, match As Long
    Dim A_WB As Workbook, B_WB As Workbook, Result_WB As Workbook
    Dim A_WS As Worksheet, B_WS As Worksheet, Result_WS As Worksheet
    
    Dim A_arr As Variant, B_arr As Variant

    Dim Result_Row As Long = 2 ' Variable for the number of row lines in the worksheet result


    Dim A_Path As String = "FULL PATH TO THE XLSX A"
    Dim A_colRef As Integer = 3 
    Dim A_colLnk As Integer = 6   
    Dim B_Path As String = "FULL PATH TO THE XLSX B"
    Dim B_colRef As Integer = 3 
    Dim B_colLnk As Integer = 6
    Dim A_SheetName As String = "Feuil1"   'sheet name of 'Our' workbook
    Dim B_SheetName As String = "Feuil1" 

    Set Result_WB = ThisWorkbook
    Set Result_WS = C_WB.Sheets(Result_SheetName)
    

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

    Result_WS.Cells.ClearContents 'EEEEERRRRRRAAAAAAASSSSSEEEEEEE EVERYTHING!!!!!! MOUHAHAHAHAHAHA

    ' Pause screen update  
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Lets put some headers
    Result_WS.Cells(1,1) = "Workbook_A"
    Result_WS.Cells(1,2) = "Ref_A"
    Result_WS.Cells(1,3) = "Lnk_A"
    Result_WS.Cells(1,4) = "Lnk_B"
    Result_WS.Cells(1,5) = "Ref_B"
    Result_WS.Cells(1,6) = "Workbook_B"

    For row = 2 To UBound(A_arr, 1)
        match = 0
        On Error Resume Next
        match = WorksheetFunction.Match(B_arr(row, 1), A_arr, 0)
        On Error GoTo 0
        
        Result_WS.Cells(Result_Row,1) = "Prout A"
        Result_WS.Cells(Result_Row,2) = A_WS.Cells(match, A_colRef).Value
        Result_WS.Cells(Result_Row,3) = A_WS.Cells(match, A_colLnk).Value

        If match <> 0 Then
            ' B_WS.Cells(row, B_colLnk).Value = A_WS.Cells(match, A_colLnk).Value
            ' MsgBox "Ref A " & A_WS.Cells(match, A_colRef).Value & "Ref B" &  & "Lnk A " & A_WS.Cells(match, A_colLnk).Value & "Lnk B" & B_WS.Cells(row, B_colLnk).Value
            Result_WS.Cells(Result_Row,4) = B_WS.Cells(row, B_colLnk).Value
            Result_WS.Cells(Result_Row,5) = B_WS.Cells(row, B_colRef).Value

        End If

        Result_WS.Cells(Result_Row,6) = "Prout B"

        Result_Row = Result_Row + 1

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