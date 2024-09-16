Attribute VB_Name = "basFlexgrid"
'
'   Flexgrid ���W���[��
'
Option Explicit


Public lngClassCnt As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �O���b�h�ɃO���b�h�A�C�e�����Z�b�g����
'
'   ���l: �Ȃ�
'
Public Sub SetItem(ByRef itemA As clsGridItem, ByRef wflexgrid As ctlWrappedGrid, row As Long, col As Long)
On Error GoTo errH
    Set itemA = New clsGridItem
    
    wflexgrid.Grid.col = col
    wflexgrid.Grid.row = row
    
    itemA.Alignment = wflexgrid.Grid.CellAlignment
    itemA.BGColor = wflexgrid.Grid.CellBackColor
    itemA.FRColor = wflexgrid.Grid.CellForeColor
    
    itemA.Strikethru = wflexgrid.Grid.CellFontStrikeThrough
    itemA.Text = wflexgrid.Grid.Text
    
    Set itemA = wflexgrid.mMSFlexData.ItemMatrix(row, col)
    
    Exit Sub
errH:
    gApp.ErrLog
    gApp.Log "SetItem Error"
    Resume Next
End Sub


'
'   �@�\: �O���b�h�A�C�e�����R�s�[����
'
'   ���l: �Ȃ�
'
Public Sub SetItemFrmFlex(ByRef itemA As clsGridItem, ByRef itemB As clsGridItem)
On Error GoTo errH
    Set itemA = New clsGridItem
        
    Set itemA = itemB
    
    Exit Sub
errH:
    gApp.ErrLog
    gApp.Log "SetItemFrmFlex Error"
    Resume Next
End Sub


'
'   �@�\: �O���b�h�Ƀf�[�^���Z�b�g����
'
'   ���l: �Ȃ�
'
Public Sub SetFlexGrid(ByRef itemA As clsGridItem, ByRef mMSFlexData As clsMSFlexData, _
    ByRef mflexgrid As MSFlexGrid, row As Long, col As Long)
On Error GoTo errH
    
    Dim tmpcol As Long
    Dim tmprow As Long
    
    tmpcol = mflexgrid.col
    tmprow = mflexgrid.row
    
    mflexgrid.col = col
    mflexgrid.row = row
    mflexgrid.CellAlignment = itemA.Alignment
    mflexgrid.CellBackColor = itemA.BGColor
    mflexgrid.CellForeColor = itemA.FRColor
        
    With itemA
        mMSFlexData.SetItemMatrix row, col, .Text, .ToolTip, .Alignment, _
            .Link, .Key, .BGColor, .FRColor, .Strikethru, .SortString
    End With

    mflexgrid.CellFontStrikeThrough = itemA.Strikethru
    mflexgrid.Text = itemA.Text
    
    mflexgrid.col = tmpcol
    mflexgrid.row = tmprow
    
    Exit Sub
errH:
    gApp.ErrLog
    gApp.Log "SetFlexGrid Error"
    Resume Next
End Sub


'
'   �@�\: �O���b�h�𒲐�����
'
'   ���l: �Ȃ�
'
Public Sub AutoFitFlexGrid(ByVal flx As MSFlexGrid, lngStart As Long, lngEnd As Long, _
    Optional blnCollapse As Boolean = False, Optional blnConvertCR As Boolean = False, _
    Optional intBeginningRow As Integer = 0)
On Error GoTo errH
    Dim r As Long
    Dim c As Long
    Dim cell_wid As Single
    Dim col_wid As Single
    Dim row_hei As Single
    
    Dim row_mult As Single

    Dim fontAdj As Double
    Dim intLines As Integer

    Dim strTemp  As String

    fontAdj = flx.Font.Size * Screen.TwipsPerPixelX
    row_hei = 198
    
    row_mult = row_hei

    ' �Z�����ƍ����̏����l
    For c = lngStart To lngEnd
        flx.ColWidth(c) = 900
    Next c
    For r = 0 To flx.Rows - 1
        flx.RowHeight(r) = 225
    Next r

    For c = lngStart To lngEnd
        col_wid = 0
        For r = intBeginningRow To flx.Rows - 1
            cell_wid = GetStringSize(flx.TextMatrix(r, c), intLines, blnCollapse, _
                blnConvertCR)

            If col_wid < cell_wid Then col_wid = cell_wid
            row_hei = row_mult * intLines
            If row_hei > flx.RowHeight(r) Then flx.RowHeight(r) = row_hei
        Next r

        flx.ColWidth(c) = col_wid
    Next c

    Exit Sub
errH:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: ������̃T�C�Y�𓾂�
'
'   ���l: �Ȃ�
'
Private Function GetStringSize(ByVal strInput As String, ByRef intLines As Integer, _
    blnCollapse As Boolean, blnConvertCR As Boolean) As Long
On Error GoTo errH
    Dim arrByte() As Byte, lngUB As Long, lngLoop As Long
    Dim lngSize As Long     '' ������̒���
    Dim lngLSize As Long    '' �����s�̕�����̒���
    Dim blnCrLf As Boolean  '' vbCrLf�t���O
    
    ' �o�C�g�z��ɃZ�b�g����
    strInput = StrConv(strInput, vbFromUnicode)
    lngUB = LenB(strInput)
    
    If lngUB = 0 Then Exit Function
    
    ReDim arrByte(lngUB - 1) As Byte
    arrByte = strInput
    
    ' ������
    lngSize = 0
    lngLSize = 0
    intLines = 1
    blnCrLf = False
        
    If blnCollapse Then
        If UBound(arrByte) = 0 Then
            If arrByte(0) = 255 Then
                GetStringSize = 0
                Exit Function
            End If
        End If
    End If
    
    For lngLoop = 0 To lngUB - 1
        If arrByte(lngLoop) = 13 Then
            If lngLoop <> lngUB - 1 Then
                If arrByte(lngLoop + 1) = 10 Then
                    intLines = intLines + 1
            
                    If lngSize > lngLSize Then
                        lngLSize = lngSize
                    End If
                    lngSize = 0
                
                ElseIf blnConvertCR Then
                    intLines = intLines + 1
                End If
            End If
        Else
            lngSize = lngSize + 100
        End If
    Next
    
    If lngSize > lngLSize Then lngLSize = lngSize
    lngLSize = lngLSize + 45
    
    ' �Q�o�C�g�����ɑΉ�
    If lngLSize < 250 Then
        GetStringSize = 320
    ElseIf lngLSize < 450 Then
        GetStringSize = 480
    Else
        GetStringSize = lngLSize
    End If
    
    Exit Function
errH:
    gApp.ErrLog
    Resume Next
End Function


'
'   �@�\: �O���b�h���\�[�g����
'
'   ���l: �Ȃ�
'
Public Sub SortFlexGrid(ByVal ctlFlx As ctlWrappedGrid, lngCol As Long)
On Error GoTo errH
    Dim r As Long, c As Long
    Dim inR As Long, inC As Long
    Dim strSearch As String
    
    Dim tmpAlignment() As Long
    Dim tmpBGColor() As Long
    Dim tmpFRColor() As Long
    Dim tmpKey() As String
    Dim tmpLink() As String
    Dim tmpStrikethru() As Boolean
    Dim tmpText() As String
    Dim tmpToolTip() As String
    
    With ctlFlx
        For r = 1 To .Grid.Rows - 1
            .Grid.col = lngCol
            .Grid.row = r
            strSearch = .Grid.Text
            
            ReDim tmpAlignment(.Grid.Cols)
            ReDim tmpBGColor(.Grid.Cols)
            ReDim tmpFRColor(.Grid.Cols)
            ReDim tmpKey(.Grid.Cols)
            ReDim tmpLink(.Grid.Cols)
            ReDim tmpStrikethru(.Grid.Cols)
            ReDim tmpText(.Grid.Cols)
            ReDim tmpToolTip(.Grid.Cols)
            
            ' �l���ꎞ�ۑ�
            For c = 0 To .Grid.Cols - 1
                tmpAlignment(c) = .mMSFlexData.ItemMatrix(r, c).Alignment
                tmpBGColor(c) = .mMSFlexData.ItemMatrix(r, c).BGColor
                tmpFRColor(c) = .mMSFlexData.ItemMatrix(r, c).FRColor
                tmpKey(c) = .mMSFlexData.ItemMatrix(r, c).Key
                tmpLink(c) = .mMSFlexData.ItemMatrix(r, c).Link
                tmpStrikethru(c) = .mMSFlexData.ItemMatrix(r, c).Strikethru
                tmpText(c) = .mMSFlexData.ItemMatrix(r, c).Text
                tmpToolTip(c) = .mMSFlexData.ItemMatrix(r, c).ToolTip
            Next c
            
            For inR = 1 To .Grid.Rows - 1
                If strSearch = .mMSFlexData.ItemMatrix(inR, lngCol).Text Then
                    For c = 0 To .Grid.Cols - 1
                        ' ��s�ɒl��߂�
                        .mMSFlexData.ItemMatrix(r, c).Alignment = .mMSFlexData.ItemMatrix(inR, c).Alignment
                        .mMSFlexData.ItemMatrix(r, c).BGColor = .mMSFlexData.ItemMatrix(inR, c).BGColor
                        .mMSFlexData.ItemMatrix(r, c).FRColor = .mMSFlexData.ItemMatrix(inR, c).FRColor
                        .mMSFlexData.ItemMatrix(r, c).Key = .mMSFlexData.ItemMatrix(inR, c).Key
                        .mMSFlexData.ItemMatrix(r, c).Link = .mMSFlexData.ItemMatrix(inR, c).Link
                        .mMSFlexData.ItemMatrix(r, c).Strikethru = .mMSFlexData.ItemMatrix(inR, c).Strikethru
                        .mMSFlexData.ItemMatrix(r, c).Text = .mMSFlexData.ItemMatrix(inR, c).Text
                        .mMSFlexData.ItemMatrix(r, c).ToolTip = .mMSFlexData.ItemMatrix(inR, c).ToolTip
                        
                        ' �ꎞ�ۑ������l�����ɖ߂�
                        .mMSFlexData.ItemMatrix(inR, c).Alignment = tmpAlignment(c)
                        .mMSFlexData.ItemMatrix(inR, c).BGColor = tmpBGColor(c)
                        .mMSFlexData.ItemMatrix(inR, c).FRColor = tmpFRColor(c)
                        .mMSFlexData.ItemMatrix(inR, c).Key = tmpKey(c)
                        .mMSFlexData.ItemMatrix(inR, c).Link = tmpLink(c)
                        .mMSFlexData.ItemMatrix(inR, c).Strikethru = tmpStrikethru(c)
                        .mMSFlexData.ItemMatrix(inR, c).Text = tmpText(c)
                        .mMSFlexData.ItemMatrix(inR, c).ToolTip = tmpToolTip(c)
                    Next c
                    
                    Exit For
                End If
            Next inR
        Next r
    End With
    
    Exit Sub
errH:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �O���b�h��j������
'
'   ���l: �Ȃ�
'
Public Sub DestroyFlexGrid(ByRef ctlFlx As ctlWrappedGrid)
On Error GoTo errH
    Set ctlFlx.mMSFlexData = Nothing
    
    Exit Sub
errH:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �ăA���P�[�g
'
'   ���l: �Ȃ�
'
Public Sub ReallocateFlexGrid(ByRef ctlFlx As ctlWrappedGrid)
On Error GoTo errH
    ctlFlx.mMSFlexData.Reallocate
    
    Exit Sub
errH:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: ���O��j������
'
'   ���l: �Ȃ�
'
Public Sub ClearFlexLog()
On Error GoTo errH
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    
    Dim filename As String
    filename = "c:\umakichiDebug.txt"
    
    If Not fs.FileExists(filename) Then
        Set ts = fs.CreateTextFile(filename, True)
    Else
        Set ts = fs.OpenTextFile(filename, ForWriting, False)
    End If
    
    ts.Close
    
    Set ts = Nothing
    Set fs = Nothing

    Exit Sub
errH:
    gApp.ErrLog
    gApp.Log "Clear write Log error"
End Sub


'
'   �@�\: ���O���o��
'
'   ���l: �Ȃ�
'
Public Sub WriteFlexLog(s As String)
On Error GoTo errH
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    
    Dim filename As String
    filename = "c:\umakichiDebug.txt"
    
    If Not fs.FileExists(filename) Then
        Set ts = fs.CreateTextFile(filename, True)
    Else
        Set ts = fs.OpenTextFile(filename, ForAppending, False)
    End If
    
    ts.WriteLine s
    ts.Close
    
    Set ts = Nothing
    Set fs = Nothing

    Exit Sub
errH:
    gApp.ErrLog
    gApp.Log "Write Log error"
End Sub

