VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   グリッドデータクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ToolTipChange(ToolTipText As String)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mItems() As clsGridItem                '' グリッドアイテム
Private mCols As Long                          '' カラム数
Private mRows As Long                          '' ロウ数


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: カラム数プロパティ
'
'   備考: 変更と同時にグリッドアイテムの要素数を調整する
'
Public Property Let Cols(RHS As Long)
Attribute Cols.VB_Description = "カラム数"
    Dim i       As Long
    Dim oldCols As Long
    oldCols = mCols
    mCols = RHS
    ReDim Preserve mItems(0 To (mCols * mRows) - 1)
    If oldCols < mCols Then
        For i = oldCols * mRows To (mRows * mCols) - 1
            Set mItems(i) = New clsGridItem
        Next i
    End If
End Property


'
'   機能: カラム数取得プロパティ
'
'   備考: なし
'
Public Property Get Cols() As Long
    Cols = mCols
End Property


'
'   機能: ロウ数プロパティ
'
'   備考: 変更と同時にグリッドアイテムの要素数を調整する
'
Public Property Let Rows(RHS As Long)
Attribute Rows.VB_Description = "ロウ数"
    Dim i As Long
    Dim oldRows As Long
    oldRows = mRows
    mRows = RHS
    ReDim Preserve mItems(0 To (mCols * mRows) - 1)
    If oldRows < mRows Then
        For i = oldRows * mCols To (mRows * mCols) - 1
            Set mItems(i) = New clsGridItem
        Next i
    End If
End Property


'
'   機能: ロウ数取得プロパティ
'
'   備考: 変更と同時にグリッドアイテムの要素数を調整する
'
Public Property Get Rows() As Long
    Rows = mRows
End Property


'
'   機能: グリッドアイテムを２次元的に参照する
'
'   備考: アイテムにアクセスする為
'
Public Property Get ItemMatrix(row As Long, col As Long) As clsGridItem
On Error GoTo EH
    Set ItemMatrix = mItems(row * mCols + col)
    Exit Property
EH:
    gApp.ErrLog
End Property


'
'   機能: グリッドアイテムを１次元的に参照する
'
'   備考: アイテムにアクセスする為
'
Public Property Get ItemArray(Index As Long) As clsGridItem
Attribute ItemArray.VB_Description = "グリッドアイテムを１次元的に参照する\r\nclsGridItemを返す"
    Set ItemArray = mItems(Index)
End Property


'
'   機能: グリッドアイテムを１次元的に参照する
'
'   備考: アイテムにアクセスする為
'
Public Property Set ItemArray(Index As Long, ByRef RHS As clsGridItem)
    Set mItems(Index) = RHS
End Property



'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: グリッドアイテムを複製する
'
'   備考: なし
'
Public Function Clone() As clsGridData
    Dim newGD   As clsGridData
    Dim i       As Long
    Set newGD = New clsGridData
    
    With newGD
        .Cols = Me.Cols
        .Rows = Me.Rows
        For i = 0 To UBound(mItems)
            Set .ItemArray(i) = mItems(i).Clone
        Next i
    End With
    
    Set Clone = newGD
End Function


'
'   機能: グリッドアイテムをカラム方向に連続して設定するユーティリティー
'
'   備考: 引き数 row - ロウ番号
'                col - カラム番号
'                Text - 表示文字列
'                ToolTip - ツールチップ文字列
'                AlignChar - 配置指定子
'                Link - リンク記号
'                Key - リンクキー
'                BGColor - バックカラー
'                FRColor - フォアカラー
' 　　　  ２段のカラムヘッダの登録に便利
'
Public Sub SetItemMatrix(row As Long, ByRef col As Long _
                        , Optional Text As String = "" _
                        , Optional ToolTip As String = "" _
                        , Optional AlignChar As String = "<-" _
                        , Optional Link As String = "" _
                        , Optional key As String = "" _
                        , Optional BGColor As Long = &H0 _
                        , Optional FRColor As Long = &H0 _
                        , Optional Strikethru As Boolean = False _
                        , Optional SortString As String = "" _
                        )
On Error GoTo Errorhandler
    With mItems(row * mCols + col)
        .Text = Text
        .ToolTip = ToolTip
        Select Case AlignChar
            Case "<^": .Alignment = flexAlignLeftTop
            Case "<-": .Alignment = flexAlignLeftCenter
            Case "<_": .Alignment = flexAlignLeftBottom
            Case "^^": .Alignment = flexAlignCenterTop
            Case "^-": .Alignment = flexAlignCenterCenter
            Case "^_": .Alignment = flexAlignCenterBottom
            Case ">^": .Alignment = flexAlignRightTop
            Case ">-": .Alignment = flexAlignRightCenter
            Case ">_": .Alignment = flexAlignRightBottom
            Case Else
                Err.Raise 0, , "AlignCharが間違い"
        End Select
        .Link = Link
        .key = key
        .BGColor = BGColor
        .FRColor = FRColor
        .Strikethru = Strikethru
        .SortString = SortString
    End With
    col = col + 1
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: ２次元的に指定したセルがキーをもっていれば真
'
'   備考: キーをもっていれば真、もっていなければ偽
'
Public Function HasAKey(row As Long, col As Long) As Boolean
Attribute HasAKey.VB_Description = "あるグリッドアイテムがキーを持っていれば真"
    HasAKey = ((mItems(row * mCols + col).Link <> "") And (mItems(row * mCols + col).key <> ""))
End Function


'
'   機能: 終了処理
'
'   備考: なし
'
Public Sub Free()
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 初期化処理
'
'   備考: なし
'
Private Sub Class_Initialize()
    ' 循環参照チェック用、オブジェクトカウンタ
    gDebugCounter_clsGridData = gDebugCounter_clsGridData + 1
    gApp.InitLog Me, gDebugCounter_clsGridData
    
    mCols = 1
    mRows = 1
    ReDim mItems(0)
    Set mItems(0) = New clsGridItem
    ReDim mblnSortFlag(0)
End Sub


'
'   機能: 終了処理
'
'   備考: なし
'
Private Sub Class_Terminate()
    ' 循環参照チェック用、オブジェクトカウンタ
    gDebugCounter_clsGridData = gDebugCounter_clsGridData - 1
    gApp.TermLog Me, gDebugCounter_clsGridData
    
    Erase mItems
End Sub
