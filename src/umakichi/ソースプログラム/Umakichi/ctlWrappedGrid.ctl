VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ctlWrappedGrid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1605
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2831
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlWrappedGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   MSFlexGrid�����b�v���郆�[�U�[�R���g���[��
'
'   MSFlexGrid������R���e�i(ctlPane)�ɒ��ڔz�u�����
'   ToolTipText���\������Ȃ����ւ̑Ή��̈׃��b�v����
'

Option Explicit


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngMouseCol As Long
Private mlngMouseRow As Long
Private mlngPrevCol As Long
Private mlngPrevRow As Long

Private mlngHalt As Long


' �\�[�g�����z��i�J�������j
Private mblnSortOrder() As Boolean


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public mMSFlexData As clsMSFlexData


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
Event AfterSort(ByVal col As Long, Order As Integer)
Event BeforeSort(ByVal col As Long, Order As Integer)
Event ToolTipChange(ToolTipText As String)
Event Click()


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �O���b�h�擾�v���p�e�B
'
'   ���l: �Ȃ�
'
Public Property Get Grid() As MSFlexGrid
    Set Grid = MSFlexGrid1
End Property


'
'   �@�\: �\�[�g�I�[�_�擾�v���p�e�B
'
'   ���l: �Ȃ�
'
Public Property Get SortOrder(Index As Integer) As Boolean
    SortOrder = mblnSortOrder(Index)
End Property


'
'   �@�\: �\�[�g�I�[�_�ݒ�v���p�e�B
'
'   ���l: �Ȃ�
'
Public Property Let SortOrder(Index As Integer, ByVal RHS As Boolean)
    mblnSortOrder(Index) = RHS
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �X�N���[����C�x���g
'
'   ���l: �Ȃ�
'
Private Sub MSflexGrid1_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrorHandler
    RaiseEvent AfterScroll(OldTopRow, OldLeftCol, NewTopRow, NewLeftCol)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\�[�g��C�x���g
'
'   ���l: �Ȃ�
'
Private Sub MSFlexGrid1_AfterSort(ByVal col As Long, Order As Integer)
On Error GoTo ErrorHandler
    RaiseEvent AfterSort(col, Order)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\�[�g�O�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub MSFlexGrid1_BeforeSort(ByVal col As Long, Order As Integer)
On Error GoTo ErrorHandler
    RaiseEvent BeforeSort(col, Order)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ��r�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub MSFlexGrid1_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrorHandler
    Dim col As Long
    Dim a As String
    Dim b As String
    
    col = MSFlexGrid1.col
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �}�E�X���[�u�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    With MSFlexGrid1
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    
    If msrow < 0 Or mscol < 0 Then
        MSFlexGrid1.MousePointer = vbDefault
        MSFlexGrid1.ToolTipText = ""
        Exit Sub
    End If
    
    If msrow <> mlngMouseRow Or mscol <> mlngMouseCol Then
        mlngPrevRow = mlngMouseRow
        mlngPrevCol = mlngMouseCol
        mlngMouseRow = msrow
        mlngMouseCol = mscol
        
        RaiseEvent MouseMove(Button, Shift, X, Y)
        
        If msrow < 0 Or mscol < 0 Then
            MSFlexGrid1.MousePointer = vbDefault
        End If
    
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub MSFlexGrid1_Click()
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    With MSFlexGrid1
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    If msrow < 0 Or mscol < 0 Then
        Exit Sub
    End If
    RaiseEvent Click
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �}�E�X�_�E���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    If MSFlexGrid1.MouseCol < 0 Or MSFlexGrid1.MouseRow < 0 Then
        Exit Sub
    End If
    If Button = vbRightButton Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���E�E�J�����ύX�C�x���g
'
'   ���l: �J���������ύX���ꂽ��A�\�[�g�����z�񐔂����킹��
'
Private Sub MSFlexGrid1_RowColChange()
On Error GoTo ErrorHandler
    ReDim Preserve mblnSortOrder(0 To MSFlexGrid1.Cols - 1)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[��������
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler

    Set mMSFlexData = New clsMSFlexData

    MSFlexGrid1.Move 0, 0, UserControl.width, UserControl.Height
    ReDim mblnSortOrder(0 To MSFlexGrid1.Cols - 1)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[���̃��T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    ' �ő���Ƀt�B�b�g
    MSFlexGrid1.Move 0, 0, UserControl.width, UserControl.Height
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �Z���̏�Ԃ�߂�
'
'   ���l: �Ȃ�
'
Private Sub SetCellOriginally(row As Long, col As Long)
    Dim item As clsGridItem
    
    gApp.Log "Set Cell Originally S - " & row & " - " & col
    
    Dim tmprow As Long
    Dim tmpcol As Long
    
    mlngHalt = mlngHalt + 1
    gApp.Log "halt - " & mlngHalt
    If mlngHalt > 1 Then
        gApp.Log "W A R N I N G ! ! ! - " & mlngHalt
    End If
            
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    
    If item.HasAKey And item.FRColor = 0 Then
        ' �L�[�������Ă���Ȃ�(����FRColor�����w��Ȃ�)
        With MSFlexGrid1
            
            tmprow = .row
            tmpcol = .col
            
            .row = row
            .col = col
            .CellForeColor = RGB(0, 0, 255)
            .CellFontUnderline = True
            
            .row = tmprow
            .col = tmpcol
        End With
    Else
        ' �L�[�������Ă��Ȃ��Ȃ�
        With MSFlexGrid1
            
            tmprow = .row
            tmpcol = .col
            
            .row = row
            .col = col
            .CellForeColor = item.FRColor
            .CellFontUnderline = False
            
            .row = tmprow
            .col = tmpcol
        End With
    End If
    
    gApp.Log "Set Cell Originally E - " & row & " - " & col
End Sub


'
'   �@�\: �Z���������N����������
'
'   ���l: �Ȃ�
'
Private Sub SetCellClickable(row As Long, col As Long)
    Dim item As clsGridItem
    
    gApp.Log "Set Cell Clickable S - " & row & " - " & col
    
    Dim tmprow As Long
    Dim tmpcol As Long
        
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    If item.HasAKey Then
        With MSFlexGrid1
            
            tmprow = .row
            tmpcol = .col
            
            .row = row
            .col = col
            .CellForeColor = ColorLinked
            .MousePointer = vbCustom
            Set .MouseIcon = LoadResPicture(101, vbResCursor)
            
            .row = tmprow
            .col = tmpcol
        End With
    Else
        MSFlexGrid1.MousePointer = vbDefault
    End If
    
    gApp.Log "Set Cell Clickable E - " & row & " - " & col
End Sub


'
'   �@�\: �אڂ����Ԃ�߂��ׂ��Z�����ׂĂ̏�Ԃ�߂�
'
'   ���l: �Ȃ�
'
Private Sub ReflexiveSetCellOriginally(row As Long, col As Long, Optional HorizontalFlag As Boolean = False)
    Dim item As clsGridItem
    
    gApp.Log "Reflexive Set Cell Originally S - " & row & " - " & col
    
    
    If row < 0 Or col < 0 Or row >= MSFlexGrid1.Rows Or col >= MSFlexGrid1.Cols Then
        Exit Sub
    End If
    
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    Call SetCellOriginally(row, col)
    
    If col >= 1 Then ' ��
        Call CallReflexiveSetCellOriginally(item, row, col - 1, HorizontalFlag)
    End If
    If row >= 1 And Not HorizontalFlag Then ' ��
        Call CallReflexiveSetCellOriginally(item, row - 1, col, HorizontalFlag)
    End If
    If col <= MSFlexGrid1.Cols - 2 Then ' �E
        Call CallReflexiveSetCellOriginally(item, row, col + 1, HorizontalFlag)
    End If
    If row <= MSFlexGrid1.Rows - 2 And Not HorizontalFlag Then ' ��
        Call CallReflexiveSetCellOriginally(item, row + 1, col, HorizontalFlag)
    End If

    gApp.Log "Reflexive Set Cell Originally E - " & row & " - " & col
End Sub


'
'   �@�\: �Z�����r���ď�Ԃ�߂��ׂ��Z���ł���΁A�ċA�Ăяo������
'
'   ���l: �Ȃ�
'
Private Sub CallReflexiveSetCellOriginally(ByRef itemA As clsGridItem, row As Long, col As Long, Optional HorizontalFlag As Boolean = False)
    Dim itemB As clsGridItem
    
    Dim tmprow As Long
    Dim tmpcol As Long
    
    gApp.Log "Call Reflexive Set Cell Originally S - " & row & " - " & col
    
    
    Call SetItemFrmFlex(itemB, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    tmprow = MSFlexGrid1.row
    tmpcol = MSFlexGrid1.col

    MSFlexGrid1.col = col
    MSFlexGrid1.row = row
    If itemB.Key = itemA.Key And MSFlexGrid1.CellForeColor = ColorLinked Then
        Call ReflexiveSetCellOriginally(row, col, HorizontalFlag)
    End If
    
    MSFlexGrid1.col = tmpcol
    MSFlexGrid1.row = tmprow
    
    gApp.Log "Call Reflexive Set Cell Originally E - " & row & " - " & col
End Sub


'
'   �@�\: �אڂ��郊���N��������ׂ��Z�����ׂĂ𔽉�������
'
'   ���l: �Ȃ�
'
Private Sub ReflexiveSetCellClickable(row As Long, col As Long, Optional HorizontalFlag As Boolean = False)
    Dim item As clsGridItem
    
    gApp.Log "Reflexive Set Cell Clickable S - " & row & " - " & col
    
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    If Not item.HasAKey Then
        Exit Sub
    End If
    
    Call SetCellClickable(row, col)
    
    If col >= 1 Then ' ��
        Call CallReflexiveSetCellClickable(item, row, col - 1, HorizontalFlag)
    End If
    If row >= 1 And Not HorizontalFlag Then ' ��
        Call CallReflexiveSetCellClickable(item, row - 1, col, HorizontalFlag)
    End If
    If col <= MSFlexGrid1.Cols - 2 Then ' �E
        Call CallReflexiveSetCellClickable(item, row, col + 1, HorizontalFlag)
    End If
    If row <= MSFlexGrid1.Rows - 2 And Not HorizontalFlag Then ' ��
        Call CallReflexiveSetCellClickable(item, row + 1, col, HorizontalFlag)
    End If

    gApp.Log "Reflexive Set Cell Clickable E - " & row & " - " & col
End Sub


'
'   �@�\: �Z�����r���ă����N��������ׂ��Z���ł���΁A�ċA�Ăяo������
'
'   ���l: �Ȃ�
'
Private Sub CallReflexiveSetCellClickable(ByRef itemA As clsGridItem, row As Long, col As Long, Optional HorizontalFlag As Boolean = False)
    Dim itemB As clsGridItem
    
    Dim tmprow As Long
    Dim tmpcol As Long
    
    gApp.Log "Call Reflexive Set Cell Clickable S - " & row & " - " & col
    
    Call SetItemFrmFlex(itemB, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    tmpcol = MSFlexGrid1.col
    tmprow = MSFlexGrid1.row

    MSFlexGrid1.col = col
    MSFlexGrid1.row = row
        
    If itemB.Key = itemA.Key And MSFlexGrid1.CellForeColor <> ColorLinked Then
        Call ReflexiveSetCellClickable(row, col, HorizontalFlag)
    End If
    
    MSFlexGrid1.col = tmpcol
    MSFlexGrid1.row = tmprow
    
    gApp.Log "Call Reflexive Set Cell Clickable E - " & row & " - " & col
End Sub


'
'   �@�\: ���O�������o��
'
'   ���l: �Ȃ�
'
Private Sub WriteLog(s As String)
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
    
    ts.Write s
    ts.Close
    
    Set ts = Nothing
    Set fs = Nothing

    Exit Sub
errH:
    gApp.ErrLog
    gApp.Log "Write Log error"
End Sub


'
'   �@�\: ���O���N���A����
'
'   ���l: �Ȃ�
'
Private Sub ClearWriteLog()
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
    
    ts.WriteLine "FlexGrid contents :"
    ts.WriteLine "row,col: .CellForeColor-item.FRColor/.CellAlignment-item.Alignment/" & _
        ".CellBackColor-item.Alignment/.Text-item.Text"
    ts.Close
    
    Set ts = Nothing
    Set fs = Nothing

    Exit Sub
errH:
    gApp.ErrLog
    gApp.Log "Clear write Log error"
End Sub


'
'   �@�\: ���[�U�R���g���[���I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Terminate()
    Set mMSFlexData = Nothing
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: clsGridData �������Ƃ�A�O���b�h�ɑ}������
'
'   ���l: .Rows, .Cols ���v���p�e�B���ݒ肷��
'
Public Sub InsertGrid(ByRef GridData As clsGridData)
On Error GoTo ErrorHandler
    Dim i     As Long ' ArrayIndex
    Dim row   As Long ' MatrixIndex Row
    Dim col   As Long ' MatrixIndex Col
    Dim item  As clsGridItem
    Dim sngST As Single
    Dim sngET As Single
    
    sngST = Timer
   
    mlngHalt = 0
    
    mlngPrevCol = -1
    mlngPrevRow = -1
    
    With MSFlexGrid1
        .Redraw = False
        .Clear
        
        ' �J�������A���E����ݒ�
        .Cols = GridData.Cols
        .Rows = GridData.Rows
                        
        ' �S�Z���̐ݒ胋�[�v
        For i = 0 To (.Cols * .Rows) - 1

            ' ���W�l
            col = i Mod .Cols
            row = Int(i / .Cols)

            ' �A�C�e��
            Set item = GridData.ItemArray(i)

            ' �\��������
            .TextArray(i) = item.Text
            
            .row = row
            .col = col
            
            ' �O���b�h�A�C�e����ݒ�
            mMSFlexData.Cols = .Cols
            mMSFlexData.Rows = .Rows
            Call SetFlexGrid(item, mMSFlexData, MSFlexGrid1, row, col)
            
            
            If item.HasAKey And item.FRColor = 0 Then
                ' �L�[������A�F�����w��Ȃ�@�A�A���_�[���C��
                .CellForeColor = RGB(0, 0, 255)
                .CellFontUnderline = True
            Else
                ' �L�[��������΍��A�A���_�[���C������
                .CellForeColor = item.FRColor
                .CellFontUnderline = False
            End If
            
            ' �\���ʒu
            .CellAlignment = item.Alignment
            
            ' �w�i�F
            .CellBackColor = item.BGColor

            ' ��������
            .CellFontStrikeThrough = item.Strikethru
            
            ' �󔒃Z���̃\�[�g�����Ō�ɂ����
            If Trim(.Text) = "" Then
                .Text = Chr(&HFFFF)
                .CellFontSize = 1
                item.FRColor = .CellForeColor
                .CellForeColor = item.FRColor
            End If
        Next i
                
        .Redraw = True
    End With
    
    sngET = Timer
    
    If sngET - sngST > 1 Then
        gApp.Log "InsertGridTile: " & i & " " & Format$(sngET - sngST, "0.0000")
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    gApp.Log "InsertGrid Error"
    Resume Next
End Sub


'
'   �@�\: �O���b�h��MouseMove�C�x���g���ʃ��[�`�� �P�Ɣ�
'
'   ���l: Grid_MouseMove�����ԁBmFlexGrid�𗘗p����B�אڂ����������N�͓������Ȃ�
'
Public Sub MouseMoveDriven()
    Dim fg As MSFlexGrid
    Dim gd As clsGridData
    Dim item As clsGridItem
    
    Set fg = MSFlexGrid1 ' �����̂�
    
    ' �`��̒�~
    fg.Redraw = False
    
    ' �O�̏ꏊ�����ɖ߂�
    If mlngPrevRow >= 0 And mlngPrevCol >= 0 Then
        Call SetCellOriginally(mlngPrevRow, mlngPrevCol)
    End If
        
    ' �V�����}�E�X���Z���������N�\�Ȃ甽������
    If mlngMouseRow >= 0 And mlngMouseCol >= 0 Then
        Call SetCellClickable(mlngMouseRow, mlngMouseCol)
    End If
        
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * mlngMouseRow + mlngMouseCol))
    
    ' �c�[���`�b�v�e�L�X�g�̕ύX
    fg.ToolTipText = item.ToolTip
    
    ' �|�C���^�̕ύX
    If item.HasAKey Then
        MSFlexGrid1.MousePointer = vbCustom
        Set MSFlexGrid1.MouseIcon = LoadResPicture(101, vbResCursor)
    Else
        MSFlexGrid1.MousePointer = vbDefault
    End If

    RaiseEvent ToolTipChange(fg.ToolTipText)
    
    ' �`��̍ĊJ
    fg.Redraw = True

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �O���b�h��MouseMove�C�x���g���ʃ��[�`��
'
'   ���l: Grid_MouseMove�����ԁBmFlexGrid�𗘗p����B�אڂ����������N����������
'
Public Sub ReflexiveMouseMoveDriven(Optional HorizontalFlag As Boolean = False)
    Dim fg As MSFlexGrid
    Dim gd As clsGridData
    Dim item As clsGridItem
    
    Set fg = MSFlexGrid1
    
    ' �`��̒�~
    fg.Redraw = False
    
    ' �O�̏ꏊ�����ɖ߂�
    If mlngPrevRow >= 0 And mlngPrevCol >= 0 Then
        Call ReflexiveSetCellOriginally(mlngPrevRow, mlngPrevCol, HorizontalFlag)
    End If
        
    ' �V�����}�E�X���Z���������N�\�Ȃ甽������
    If mlngMouseRow >= 0 And mlngMouseCol >= 0 Then
        Call ReflexiveSetCellClickable(mlngMouseRow, mlngMouseCol, HorizontalFlag)
    End If
        
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * mlngMouseRow + mlngMouseCol))
    
    ' �c�[���`�b�v�e�L�X�g�̕ύX
    fg.ToolTipText = item.ToolTip
    
    ' �|�C���^�̕ύX
    If item.HasAKey Then
        MSFlexGrid1.MousePointer = vbCustom
        Set MSFlexGrid1.MouseIcon = LoadResPicture(101, vbResCursor)
    Else
        MSFlexGrid1.MousePointer = vbDefault
    End If

    RaiseEvent ToolTipChange(fg.ToolTipText)
    
    ' �`��̍ĊJ
    fg.Redraw = True

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J�����̍ő啝
'
'   ���l: �Ȃ�
'
Public Sub LimitedAutosize(Col1 As Long, Optional Col2 As Long, Optional Equal As Boolean, Optional ExtraSpace As Long)
    Dim i As Long
    
    If Col2 < Col1 Then
        Col2 = Col1
    End If
    
    For i = Col1 To Col2
        If MSFlexGrid1.ColWidth(i) > 3000 Then
            MSFlexGrid1.ColWidth(i) = 3000
        End If
    Next i
End Sub


'
'   �@�\: �Z�����L�[�������Ă��邩
'
'   ���l: �Ȃ�
'
Public Function HasKey(row As Long, col As Long) As Boolean
On Error GoTo ErrorHandler
    HasKey = mMSFlexData.HasAKey(row, col)
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   �@�\: ��������
'
'   ���l: �Ȃ�
'
Public Sub AutoSize(lngStart As Long, lngEnd As Long, Optional blnCollapse As Boolean = False, _
    Optional blnConvertCR As Boolean = False, Optional intBeginningRow As Integer = 0)
On Error GoTo ErrorHandler
    Call AutoFitFlexGrid(MSFlexGrid1, lngStart, lngEnd, blnCollapse, blnConvertCR, _
        intBeginningRow)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �O���b�h�𖳌�(�L��)�ɂ���
'
'   ���l: �Ȃ�
'
Public Sub FlexDisable(Optional blnDisable As Boolean = True)
On Error GoTo ErrorHandler
    MSFlexGrid1.Enabled = Not blnDisable
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub

