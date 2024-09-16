VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   グリッドアイテムクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Text         As String       '' セルに表示する文字列
Public Link         As String       '' リンク先のViewer名
Public Key          As String       '' リンク先用のキー
Public ToolTip      As String       '' セルのツールチップテキスト
Public Alignment    As Long
Public BGColor      As Long         '' 背景色
Public FRColor      As Long         '' 文字色
Public Strikethru   As Boolean      '' 取り消し線

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrSortString   As String  '' ソート用文字列

'
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ソート文字列を返す
'
'   備考: ソート文字列が無いときは表示文字列を返す
'
Public Property Get SortString() As String
    If (mstrSortString <> "") Then
        SortString = mstrSortString
    Else
        SortString = Text
    End If
End Property

'
'   機能: ソート文字列をセットする
'
'   備考: ソート文字列が無いときは表示文字列を返す
'
Public Property Let SortString(strSortString As String)
    mstrSortString = strSortString
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: キーの有無
'
'   備考: なし
'
Public Function HasAKey() As Boolean
    HasAKey = (Link <> "") And (Key <> "")
End Function


'
'   機能: グリッドアイテムを複製する
'
'   備考: なし
'
Public Function Clone() As clsGridItem
    Dim newItem As clsGridItem
    Set newItem = New clsGridItem
    
    With newItem
        .Text = Me.Text
        .Link = Me.Link
        .Key = Me.Key
        .ToolTip = Me.ToolTip
        .Alignment = Me.Alignment
        .BGColor = Me.BGColor
        .FRColor = Me.FRColor
        .Strikethru = Me.Strikethru
    End With
    
    Set Clone = newItem
End Function

