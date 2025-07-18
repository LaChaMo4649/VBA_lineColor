Attribute VB_Name = "lineColor"
Public Sub lineColorMake()
    Dim tableRangeS As String
    Dim rangeCollection As New Collection
    Dim rng As Range
    Dim st As Worksheet
    Dim clearRng As Range
    
    Set st = ThisWorkbook.Sheets("table")
    tableRangeS = GetBorderedRange(st)
    Set clearRng = st.UsedRange
    clearRng.Interior.ColorIndex = 0
    
    Set rangeCollection = lineRanges(tableRangeS)
    For Each rng In rangeCollection
        rng.Interior.ColorIndex = 34
        'colorIndex
        '0:色なし
        '1:黒
        '2:白
        '3:赤
        '4:緑
        '5:青
        '6:黄
        '7:桃
        '8:水
        '灰色   15
        '薄い赤 38
        '薄い青 34
        '薄い黄 36
    Next
    MsgBox "終了しました"
End Sub

Private Function lineRanges(tableRangeS As String) As Collection
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim firstR As String
    Dim lastR As String
    Dim temp As Long
    Dim colorLine As Long
    Set lineRanges = New Collection
    
    firstR = Left(tableRangeS, InStr(tableRangeS, ":") - 1)
    temp = InStr(firstR, "C")
    startRow = CLng(Mid(firstR, 2, temp - 2))
    startCol = CLng(Right(firstR, Len(firstR) - temp))
    lastR = Right(tableRangeS, Len(tableRangeS) - InStr(tableRangeS, ":"))
    temp = InStr(lastR, "C")
    endRow = CLng(Mid(lastR, 2, temp - 2))
    endCol = CLng(Right(lastR, Len(lastR) - temp))
    colorLine = 5 '着色開始行の初期値
    While colorLine <= endRow
        lineRanges.Add Range(Cells(colorLine, startCol), Cells(colorLine, endCol))
        colorLine = colorLine + 2
    Wend
End Function

Private Function GetBorderedRange(st As Worksheet) As String
    'シート全体の罫線が設定されている範囲をRangeで取得し、R1C1アドレス文字列で返す
    Dim ws As Worksheet
    Dim cell As Range
    Dim borderedRange As Range
    Dim checkRange As Range
    
    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("table")
    
    ' チェックする範囲を指定（例: シート全体）
    Set checkRange = ws.UsedRange
    ' 罫線があるセルを確認
    For Each cell In checkRange
        If HasBorders(cell) Then
            If borderedRange Is Nothing Then
                Set borderedRange = cell
            Else
                Set borderedRange = Union(borderedRange, cell)
            End If
        End If
    Next cell
    GetBorderedRange = borderedRange.Address(, , xlR1C1)
' 結果を表示
'    If Not borderedRange Is Nothing Then
'        Debug.Print "罫線が設定されている範囲: " & borderedRange.Address(, , xlR1C1)
'    Else
'        MsgBox "罫線が設定されているセルはありません。"
'    End If
End Function

Function HasBorders(targetCell As Range) As Boolean
' セルに罫線があるか確認する関数
    Dim i As Integer
    For i = 8 To 9
        'xlEdgeLeft (7) 範囲内の左端の罫線
        'xlEdgeTop(8)   範囲内の上側の罫線
        'xlEdgeBottom(9)範囲内の下側の罫線
        'xlEdgeRight(10)範囲内の右端の罫線
        'xlInsideVertical(11)範囲内のすべてのセルの垂直罫線
        'xlInsideHorizontal (12)範囲内のすべてのセルの水平罫線
        If targetCell.Borders(i).LineStyle <> xlNone Then
            HasBorders = True
            Debug.Print i
            Exit Function
        End If
    Next i
    HasBorders = False
End Function
