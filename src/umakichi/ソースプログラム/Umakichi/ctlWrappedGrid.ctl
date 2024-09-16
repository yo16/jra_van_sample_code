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
'   MSFlexGridをラップするユーザーコントロール
'
'   MSFlexGridを自作コンテナ(ctlPane)に直接配置すると
'   ToolTipTextが表示されない問題への対応の為ラップする
'

Option Explicit


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngMouseCol As Long
Private mlngMouseRow As Long
Private mlngPrevCol As Long
Private mlngPrevRow As Long

Private mlngHalt As Long


' ソート方向配列（カラム分）
Private mblnSortOrder() As Boolean


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public mMSFlexData As clsMSFlexData


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
Event AfterSort(ByVal col As Long, Order As Integer)
Event BeforeSort(ByVal col As Long, Order As Integer)
Event ToolTipChange(ToolTipText As String)
Event Click()


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: グリッド取得プロパティ
'
'   備考: なし
'
Public Property Get Grid() As MSFlexGrid
    Set Grid = MSFlexGrid1
End Property


'
'   機能: ソートオーダ取得プロパティ
'
'   備考: なし
'
Public Property Get SortOrder(Index As Integer) As Boolean
    SortOrder = mblnSortOrder(Index)
End Property


'
'   機能: ソートオーダ設定プロパティ
'
'   備考: なし
'
Public Property Let SortOrder(Index As Integer, ByVal RHS As Boolean)
    mblnSortOrder(Index) = RHS
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: スクロール後イベント
'
'   備考: なし
'
Private Sub MSflexGrid1_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrorHandler
    RaiseEvent AfterScroll(OldTopRow, OldLeftCol, NewTopRow, NewLeftCol)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ソート後イベント
'
'   備考: なし
'
Private Sub MSFlexGrid1_AfterSort(ByVal col As Long, Order As Integer)
On Error GoTo ErrorHandler
    RaiseEvent AfterSort(col, Order)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ソート前イベント
'
'   備考: なし
'
Private Sub MSFlexGrid1_BeforeSort(ByVal col As Long, Order As Integer)
On Error GoTo ErrorHandler
    RaiseEvent BeforeSort(col, Order)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 比較イベント
'
'   備考: なし
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
'   機能: マウスムーブイベント
'
'   備考: なし
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
'   機能: クリックイベント
'
'   備考: なし
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
'   機能: マウスダウンイベント
'
'   備考: なし
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
'   機能: ロウ・カラム変更イベント
'
'   備考: カラム数が変更されたら、ソート方向配列数も合わせる
'
Private Sub MSFlexGrid1_RowColChange()
On Error GoTo ErrorHandler
    ReDim Preserve mblnSortOrder(0 To MSFlexGrid1.Cols - 1)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロール初期化
'
'   備考: なし
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
'   機能: ユーザコントロールのリサイズイベント
'
'   備考: なし
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    ' 最大限にフィット
    MSFlexGrid1.Move 0, 0, UserControl.width, UserControl.Height
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: セルの状態を戻す
'
'   備考: なし
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
        ' キーを持っているなら(かつFRColorが未指定なら)
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
        ' キーを持っていないなら
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
'   機能: セルをリンク反応させる
'
'   備考: なし
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
'   機能: 隣接する状態を戻すべきセルすべての状態を戻す
'
'   備考: なし
'
Private Sub ReflexiveSetCellOriginally(row As Long, col As Long, Optional HorizontalFlag As Boolean = False)
    Dim item As clsGridItem
    
    gApp.Log "Reflexive Set Cell Originally S - " & row & " - " & col
    
    
    If row < 0 Or col < 0 Or row >= MSFlexGrid1.Rows Or col >= MSFlexGrid1.Cols Then
        Exit Sub
    End If
    
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    Call SetCellOriginally(row, col)
    
    If col >= 1 Then ' 左
        Call CallReflexiveSetCellOriginally(item, row, col - 1, HorizontalFlag)
    End If
    If row >= 1 And Not HorizontalFlag Then ' 上
        Call CallReflexiveSetCellOriginally(item, row - 1, col, HorizontalFlag)
    End If
    If col <= MSFlexGrid1.Cols - 2 Then ' 右
        Call CallReflexiveSetCellOriginally(item, row, col + 1, HorizontalFlag)
    End If
    If row <= MSFlexGrid1.Rows - 2 And Not HorizontalFlag Then ' 下
        Call CallReflexiveSetCellOriginally(item, row + 1, col, HorizontalFlag)
    End If

    gApp.Log "Reflexive Set Cell Originally E - " & row & " - " & col
End Sub


'
'   機能: セルを比較して状態を戻すべきセルであれば、再帰呼び出しする
'
'   備考: なし
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
'   機能: 隣接するリンク反応するべきセルすべてを反応させる
'
'   備考: なし
'
Private Sub ReflexiveSetCellClickable(row As Long, col As Long, Optional HorizontalFlag As Boolean = False)
    Dim item As clsGridItem
    
    gApp.Log "Reflexive Set Cell Clickable S - " & row & " - " & col
    
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * row + col))
    
    If Not item.HasAKey Then
        Exit Sub
    End If
    
    Call SetCellClickable(row, col)
    
    If col >= 1 Then ' 左
        Call CallReflexiveSetCellClickable(item, row, col - 1, HorizontalFlag)
    End If
    If row >= 1 And Not HorizontalFlag Then ' 上
        Call CallReflexiveSetCellClickable(item, row - 1, col, HorizontalFlag)
    End If
    If col <= MSFlexGrid1.Cols - 2 Then ' 右
        Call CallReflexiveSetCellClickable(item, row, col + 1, HorizontalFlag)
    End If
    If row <= MSFlexGrid1.Rows - 2 And Not HorizontalFlag Then ' 下
        Call CallReflexiveSetCellClickable(item, row + 1, col, HorizontalFlag)
    End If

    gApp.Log "Reflexive Set Cell Clickable E - " & row & " - " & col
End Sub


'
'   機能: セルを比較してリンク反応するべきセルであれば、再帰呼び出しする
'
'   備考: なし
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
'   機能: ログを書き出す
'
'   備考: なし
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
'   機能: ログをクリアする
'
'   備考: なし
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
'   機能: ユーザコントロール終了イベント
'
'   備考: なし
'
Private Sub UserControl_Terminate()
    Set mMSFlexData = Nothing
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: clsGridData をうけとり、グリッドに挿入する
'
'   備考: .Rows, .Cols 両プロパティも設定する
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
        
        ' カラム数、ロウ数を設定
        .Cols = GridData.Cols
        .Rows = GridData.Rows
                        
        ' 全セルの設定ループ
        For i = 0 To (.Cols * .Rows) - 1

            ' 座標値
            col = i Mod .Cols
            row = Int(i / .Cols)

            ' アイテム
            Set item = GridData.ItemArray(i)

            ' 表示文字列
            .TextArray(i) = item.Text
            
            .row = row
            .col = col
            
            ' グリッドアイテムを設定
            mMSFlexData.Cols = .Cols
            mMSFlexData.Rows = .Rows
            Call SetFlexGrid(item, mMSFlexData, MSFlexGrid1, row, col)
            
            
            If item.HasAKey And item.FRColor = 0 Then
                ' キーがあり、色が無指定なら　青、アンダーライン
                .CellForeColor = RGB(0, 0, 255)
                .CellFontUnderline = True
            Else
                ' キーが無ければ黒、アンダーライン無し
                .CellForeColor = item.FRColor
                .CellFontUnderline = False
            End If
            
            ' 表示位置
            .CellAlignment = item.Alignment
            
            ' 背景色
            .CellBackColor = item.BGColor

            ' 取り消し線
            .CellFontStrikeThrough = item.Strikethru
            
            ' 空白セルのソート順を最後にする為
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
'   機能: グリッドのMouseMoveイベント共通ルーチン 単独版
'
'   備考: Grid_MouseMoveからよぶ。mFlexGridを利用する。隣接した同リンクは同期しない
'
Public Sub MouseMoveDriven()
    Dim fg As MSFlexGrid
    Dim gd As clsGridData
    Dim item As clsGridItem
    
    Set fg = MSFlexGrid1 ' 長いので
    
    ' 描画の停止
    fg.Redraw = False
    
    ' 前の場所を元に戻す
    If mlngPrevRow >= 0 And mlngPrevCol >= 0 Then
        Call SetCellOriginally(mlngPrevRow, mlngPrevCol)
    End If
        
    ' 新しいマウス下セルがリンク可能なら反応する
    If mlngMouseRow >= 0 And mlngMouseCol >= 0 Then
        Call SetCellClickable(mlngMouseRow, mlngMouseCol)
    End If
        
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * mlngMouseRow + mlngMouseCol))
    
    ' ツールチップテキストの変更
    fg.ToolTipText = item.ToolTip
    
    ' ポインタの変更
    If item.HasAKey Then
        MSFlexGrid1.MousePointer = vbCustom
        Set MSFlexGrid1.MouseIcon = LoadResPicture(101, vbResCursor)
    Else
        MSFlexGrid1.MousePointer = vbDefault
    End If

    RaiseEvent ToolTipChange(fg.ToolTipText)
    
    ' 描画の再開
    fg.Redraw = True

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: グリッドのMouseMoveイベント共通ルーチン
'
'   備考: Grid_MouseMoveからよぶ。mFlexGridを利用する。隣接した同リンクが同期する
'
Public Sub ReflexiveMouseMoveDriven(Optional HorizontalFlag As Boolean = False)
    Dim fg As MSFlexGrid
    Dim gd As clsGridData
    Dim item As clsGridItem
    
    Set fg = MSFlexGrid1
    
    ' 描画の停止
    fg.Redraw = False
    
    ' 前の場所を元に戻す
    If mlngPrevRow >= 0 And mlngPrevCol >= 0 Then
        Call ReflexiveSetCellOriginally(mlngPrevRow, mlngPrevCol, HorizontalFlag)
    End If
        
    ' 新しいマウス下セルがリンク可能なら反応する
    If mlngMouseRow >= 0 And mlngMouseCol >= 0 Then
        Call ReflexiveSetCellClickable(mlngMouseRow, mlngMouseCol, HorizontalFlag)
    End If
        
    Call SetItemFrmFlex(item, mMSFlexData.ItemArray(MSFlexGrid1.Cols * mlngMouseRow + mlngMouseCol))
    
    ' ツールチップテキストの変更
    fg.ToolTipText = item.ToolTip
    
    ' ポインタの変更
    If item.HasAKey Then
        MSFlexGrid1.MousePointer = vbCustom
        Set MSFlexGrid1.MouseIcon = LoadResPicture(101, vbResCursor)
    Else
        MSFlexGrid1.MousePointer = vbDefault
    End If

    RaiseEvent ToolTipChange(fg.ToolTipText)
    
    ' 描画の再開
    fg.Redraw = True

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: カラムの最大幅
'
'   備考: なし
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
'   機能: セルがキーを持っているか
'
'   備考: なし
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
'   機能: 自動調整
'
'   備考: なし
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
'   機能: グリッドを無効(有効)にする
'
'   備考: なし
'
Public Sub FlexDisable(Optional blnDisable As Boolean = True)
On Error GoTo ErrorHandler
    MSFlexGrid1.Enabled = Not blnDisable
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub

