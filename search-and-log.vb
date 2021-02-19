Option Explicit

Const R_SheetName As String = "Results"

Sub Updatesheets()
    Dim rowEnd As Long, row As Long, match As Long
    Dim A_WB As Workbook, B_WB As Workbook, R_WB As Workbook
    Dim A_WS As Worksheet, B_WS As Worksheet, R_WS As Worksheet
    Dim A_colRef As Integer, A_colLnk As Integer, B_colRef As Integer, B_colLnk As Integer, R_Row As Integer
    Dim A_Path As String, B_Path As String, A_SheetName As String, B_SheetName As String
    Dim A_arr As Variant, B_arr As Variant

    R_Row = 2


    ' Need to create a loop that would loop through a config table which would have 6 columns:
    ' A_path, A_SheetName, A_colRef, A_colLnk, B_path, B_SheetName, B_colRef, B_colLnk

    A_Path = "FULL PATH TO THE XLSX A"
    B_Path = "FULL PATH TO THE XLSX B"
    A_colRef = 3 
    A_colLnk = 6   
    B_colRef = 3 
    B_colLnk = 6
    A_SheetName = "Feuil1"
    B_SheetName = "Feuil1"

    Set R_WB = ThisWorkbook
    Set R_WS = R_WB.Sheets(R_SheetName)
    
    Set A_WB = Workbooks.Open(A_Path)
    Set A_WS = A_WB.Sheets(A_SheetName)

    Set B_WB = Workbooks.Open(B_Path)
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
        match = WorksheetFunction.Match(A_arr(row, 1), B_arr, 0)
        On Error GoTo 0

        R_WS.Cells(R_Row,1) = A_WB.Name
        R_WS.Cells(R_Row,2) = A_WS.Cells(row, A_colRef).Value
        R_WS.Cells(R_Row,3) = A_WS.Cells(row, A_colLnk).Value

        If match <> 0 Then            
            R_WS.Cells(R_Row,4) = B_WS.Cells(match, B_colLnk).Value
            R_WS.Cells(R_Row,5) = B_WS.Cells(match, B_colRef).Value
        Else
            R_WS.Cells(R_Row,4) = "Plic!!!"
            R_WS.Cells(R_Row,5) = "Ploc!!!"
        End If

        R_WS.Cells(R_Row,6) = B_WB.Name
        R_Row = R_Row + 1

    Next row

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With


End Sub

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