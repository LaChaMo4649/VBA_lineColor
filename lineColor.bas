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
        '0:�F�Ȃ�
        '1:��
        '2:��
        '3:��
        '4:��
        '5:��
        '6:��
        '7:��
        '8:��
        '�D�F   15
        '������ 38
        '������ 34
        '������ 36
    Next
    MsgBox "�I�����܂���"
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
    colorLine = 5 '���F�J�n�s�̏����l
    While colorLine <= endRow
        lineRanges.Add Range(Cells(colorLine, startCol), Cells(colorLine, endCol))
        colorLine = colorLine + 2
    Wend
End Function

Private Function GetBorderedRange(st As Worksheet) As String
    '�V�[�g�S�̂̌r�����ݒ肳��Ă���͈͂�Range�Ŏ擾���AR1C1�A�h���X������ŕԂ�
    Dim ws As Worksheet
    Dim cell As Range
    Dim borderedRange As Range
    Dim checkRange As Range
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets("table")
    
    ' �`�F�b�N����͈͂��w��i��: �V�[�g�S�́j
    Set checkRange = ws.UsedRange
    ' �r��������Z�����m�F
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
' ���ʂ�\��
'    If Not borderedRange Is Nothing Then
'        Debug.Print "�r�����ݒ肳��Ă���͈�: " & borderedRange.Address(, , xlR1C1)
'    Else
'        MsgBox "�r�����ݒ肳��Ă���Z���͂���܂���B"
'    End If
End Function

Function HasBorders(targetCell As Range) As Boolean
' �Z���Ɍr�������邩�m�F����֐�
    Dim i As Integer
    For i = 8 To 9
        'xlEdgeLeft (7) �͈͓��̍��[�̌r��
        'xlEdgeTop(8)   �͈͓��̏㑤�̌r��
        'xlEdgeBottom(9)�͈͓��̉����̌r��
        'xlEdgeRight(10)�͈͓��̉E�[�̌r��
        'xlInsideVertical(11)�͈͓��̂��ׂẴZ���̐����r��
        'xlInsideHorizontal (12)�͈͓��̂��ׂẴZ���̐����r��
        If targetCell.Borders(i).LineStyle <> xlNone Then
            HasBorders = True
            Debug.Print i
            Exit Function
        End If
    Next i
    HasBorders = False
End Function
