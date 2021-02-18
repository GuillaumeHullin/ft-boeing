Option Explicit
Const miWB1ItemNumberCol As Integer = 3 
Const miWB1NumItemscol As Integer = 6   
Const miWB2ItemNumberCol As Integer = 2 
Const miWB2NumItemsCol As Integer = 6

Const msWB1SheetName As String = "Sheet1"   'sheet name of 'Our' workbook
Const msWB2SheetName As String = "Sheet1"   'Sheet name of 'Their' workbook

Dim mvaWB1Data As Variant, mvaWB2Data As Variant

Sub Updatesheets()
    Dim lRowEnd As Long, lRow As Long, lFound As Long
    Dim vWB1 As Variant, vWB2 As Variant
    Dim WB1 As Workbook, WB2 As Workbook
    Dim WS1 As Worksheet, WS2 As Worksheet

    Set WB1 = OpenWB("Select 'Our' excel file")

    If (WB1 Is Nothing) Then
        MsgBox "Macro abandoned"
        Exit Sub
    End If
    Set WS1 = Sheets(msWB1SheetName)

    Set WB2 = OpenWB("Select 'Their' excel file")
    If WB2 Is Nothing Then
        WB1.Close savechanges:=False
        MsgBox "Macro Abandoned"
        Exit Sub
    End If
    Set WS2 = Sheets(msWB2SheetName)

    '-- Store WB1 item numbers into array --
    lRowEnd = WS1.Cells(Rows.Count, miWB1ItemNumberCol).End(xlUp).Row
    mvaWB1Data = WS1.Range(Cells(1, miWB1ItemNumberCol).Address, _
                        Cells(lRowEnd, miWB1ItemNumberCol).Address).Value
                        
    '-- Store WB2 item numbers into array --
    lRowEnd = WS2.Cells(Rows.Count, miWB2ItemNumberCol).End(xlUp).Row
    mvaWB2Data = WS2.Range(Cells(1, miWB2ItemNumberCol).Address, _
                        Cells(lRowEnd, miWB2ItemNumberCol).Address).Value

    For lRow = 2 To UBound(mvaWB2Data, 1)
        lFound = 0
        On Error Resume Next
        lFound = WorksheetFunction.Match(mvaWB2Data(lRow, 1), mvaWB1Data, 0)
        On Error GoTo 0
        If lFound <> 0 Then
            ' WS2.Cells(lRow, miWB2NumItemsCol).Value = WS1.Cells(lFound, miWB1NumItemscol).Value
            MsgBox WS2.Cells(lRow, miWB2NumItemsCol).Value
        End If
    Next lRow
End Sub

Private Function OpenWB(ByVal Title As String) As Workbook
    Dim vWBName As Variant

    vWBName = Application.GetOpenFilename(filefilter:="Excel files (*.xls),*.xls", _
                                        Title:=Title)
    If vWBName = False Then
        Set OpenWB = Nothing
        Exit Function
    End If

    Workbooks.Open Filename:=vWBName
    Set OpenWB = ActiveWorkbook
End Function

