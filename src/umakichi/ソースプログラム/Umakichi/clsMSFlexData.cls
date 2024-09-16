VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSFlexData"
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
    Dim i       As Long
    Dim oldCols As Long
    oldCols = mCols
    mCols = RHS
    ReDim Preserve mItems(0 To (mCols * mRows) - 1)
    If oldCols < mCols Then
        For i = oldCols * mRows To (mRows * mCols) - 1
            Set mItems(i) = New clsGridItem
            If mItems(i) Is Nothing Then
                WriteFlexLog "memory unallocated"
            End If
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
    Dim i As Long
    Dim oldRows As Long
    oldRows = mRows
    mRows = RHS
    ReDim Preserve mItems(0 To (mCols * mRows) - 1)
    If oldRows < mRows Then
        For i = oldRows * mCols To (mRows * mCols) - 1
            Set mItems(i) = New clsGridItem
            If mItems(i) Is Nothing Then
                WriteFlexLog "memory unallocated"
            End If
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
    If Not mItems(Index) Is Nothing Then
        Set ItemArray = mItems(Index)
    End If
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
                        , Optional Alignment As Long = "0" _
                        , Optional Link As String = "" _
                        , Optional Key As String = "" _
                        , Optional BGColor As Long = &H0 _
                        , Optional FRColor As Long = &H0 _
                        , Optional Strikethru As Boolean = False _
                        , Optional SortString As String = "" _
                        )
On Error GoTo ErrorHandler
    With mItems(row * mCols + col)
        .Text = Text
        .ToolTip = ToolTip
        .Alignment = Alignment
        .Link = Link
        .Key = Key
        .BGColor = BGColor
        .FRColor = FRColor
        .Strikethru = Strikethru
        .SortString = SortString
    End With
    col = col + 1
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: ２次元的に指定したセルがキーをもっていれば真
'
'   備考: キーをもっていれば真、もっていなければ偽
'
Public Function HasAKey(row As Long, col As Long) As Boolean
    HasAKey = ((mItems(row * mCols + col).Link <> "") And (mItems(row * mCols + col).Key <> ""))
End Function


'
'   機能: 終了処理
'
'   備考: なし
'
Public Sub Free()
End Sub


'
'   機能: 再アロケート
'
'   備考: なし
'
Public Sub Reallocate()
    Dim i As Long
    ReDim mItems(0 To (mCols * mRows) - 1)
    
    For i = 0 To (mRows * mCols) - 1
        Set mItems(i) = New clsGridItem
        If mItems(i) Is Nothing Then
            WriteFlexLog "**************memory unallocated**************"
        End If
    Next i
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
    
    mCols = 1
    mRows = 1
    ReDim mItems(0)
    Set mItems(0) = New clsGridItem

End Sub


'
'   機能: 終了処理
'
'   備考: なし
'
Private Sub Class_Terminate()
    ' 循環参照チェック用、オブジェクトカウンタ
    
    Erase mItems
End Sub


