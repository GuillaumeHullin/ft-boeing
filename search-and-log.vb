Option Explicit

Const R_SheetName As String = "Results"
Const C_SheetName As String = "Config"

Sub Updatesheets()
    Dim rowEnd As Long, row As Long, conf_line As Long, match As Long, R_Row As Integer
    Dim A_WB As Workbook, B_WB As Workbook, WB As Workbook
    Dim A_WS As Worksheet, B_WS As Worksheet, R_WS As Worksheet, C_WS As Worksheet
    Dim A_colRef As Variant, A_colLnk As Variant, B_colRef As Variant, B_colLnk As Variant
    Dim A_Path As Variant, B_Path As Variant, A_SheetName As Variant, B_SheetName As Variant
    Dim A_arr As Variant, B_arr As Variant, C_arr As Variant
    Dim regEx As New RegExp
    
    R_Row = 2
    
    With regEx
            .Pattern = "$.*\["
    End With

    Set WB = ThisWorkbook
    Set R_WS = WB.Sheets(R_SheetName)
    Set C_WS = WB.Sheets(C_SheetName)
        
    R_WS.Cells.ClearContents 'EEEEERRRRRRAAAAAAASSSSSEEEEEEE EVERYTHING!!!!!! MOUHAHAHAHAHAHA
        
    '-- Store Config items into array --
    rowEnd = C_WS.Cells(Rows.Count, 1).End(xlUp).Row
    C_arr = C_WS.Range(Cells(1, 1).Address, Cells(rowEnd, 1).Address).Value
                        
    For conf_line = 2 To UBound(C_arr, 1)

        A_Path = C_WS.Cells(conf_line, 1).Value
        A_SheetName = C_WS.Cells(conf_line, 2).Value
        A_colRef = C_WS.Cells(conf_line, 3).Value 
        A_colLnk = C_WS.Cells(conf_line, 4).Value   
        B_colLnk = C_WS.Cells(conf_line, 5).Value
        B_colRef = C_WS.Cells(conf_line, 6).Value
        B_SheetName = C_WS.Cells(conf_line, 7).Value
        B_Path = C_WS.Cells(conf_line, 8).Value

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

        '-- Clean up Arr A
        For row = 2 To UBound(A_arr, 1)
            A_arr(row, 1) = regEx.Replace(A_arr(row, 1), "")
        Next row

        '-- Clean up Arr A
        For row = 2 To UBound(B_arr, 1)
            B_arr(row, 1) = regEx.Replace(B_arr(row, 1), "")
        Next row

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
    
    Next conf_line

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
