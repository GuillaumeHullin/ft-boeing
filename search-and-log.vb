Option Explicit
Const A_colRef As Integer = 3 
Const A_colLnk As Integer = 6   
Const B_colRef As Integer = 3 
Const B_colLnk As Integer = 6

Const A_SheetName As String = "Feuil1"   'sheet name of 'Our' workbook
Const B_SheetName As String = "Feuil1"   'Sheet name of 'Their' workbook

Dim A_arr As Variant, B_arr As Variant

Sub Updatesheets()
    Dim lRowEnd As Long, lRow As Long, lFound As Long
    Dim vA_WB As Variant, vB_WB As Variant
    Dim A_WB As Workbook, B_WB As Workbook
    Dim A_WS As Worksheet, B_WS As Worksheet

    Set A_WB = OpenWB("Select 'Our' excel file")

    If (A_WB Is Nothing) Then
        MsgBox "Macro abandoned"
        Exit Sub
    End If
    Set A_WS = A_WB.Sheets(A_SheetName)

    Set B_WB = OpenWB("Select 'Their' excel file")
    If B_WB Is Nothing Then
        A_WB.Close savechanges:=False
        MsgBox "Macro Abandoned"
        Exit Sub
    End If
    Set B_WS = B_WB.Sheets(B_SheetName)

    '-- Store A_WB item numbers into array --
    lRowEnd = A_WS.Cells(Rows.Count, A_colRef).End(xlUp).Row
    A_arr = A_WS.Range(Cells(1, A_colRef).Address, _
        Cells(lRowEnd, A_colRef).Address).Value
                        
    '-- Store B_WB item numbers into array --
    lRowEnd = B_WS.Cells(Rows.Count, B_colRef).End(xlUp).Row
    B_arr = B_WS.Range(Cells(1, B_colRef).Address, _
        Cells(lRowEnd, B_colRef).Address).Value

    For lRow = 2 To UBound(B_arr, 1)
        lFound = 0
        On Error Resume Next
        lFound = WorksheetFunction.Match(B_arr(lRow, 1), A_arr, 0)
        On Error GoTo 0
        
        If lFound <> 0 Then
            ' B_WS.Cells(lRow, B_colLnk).Value = A_WS.Cells(lFound, A_colLnk).Value
            MsgBox "Ref A " & A_WS.Cells(lFound, A_colRef).Value & "Ref B" & B_WS.Cells(lRow, B_colRef).Value & "Lnk A " & A_WS.Cells(lFound, A_colLnk).Value & "Lnk B" & B_WS.Cells(lRow, B_colLnk).Value
        End If

    Next lRow
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

                    '  /\_/\
                    ' ( o.o )
                    '  > ^ <