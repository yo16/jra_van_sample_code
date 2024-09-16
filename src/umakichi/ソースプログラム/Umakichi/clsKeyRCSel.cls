VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyRCSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   レコード一覧画面
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrKey         As String   '選択されたキー
Private mlngTitleRow()  As String   'グリッドでタイトルを表示する行の番号
Private mlngTitleIndex  As Long     'mlngTitleRowのインデックス

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キー文字列 代入プロパティ
'
'   備考: なし
'
Public Property Let Key(Str As String)
    mstrKey = Str
End Property


'
'   機能: キー文字列 取得プロパティ
'
'   備考: なし
'
Public Property Get Key() As String
    Key = mstrKey
End Property


'
'   機能: TitleRow 代入プロパティ
'
'   備考: なし
'
Public Property Let SetTitleRow(lng As Long)
    ReDim Preserve mlngTitleRow(mlngTitleIndex)
    mlngTitleRow(mlngTitleIndex) = lng
    mlngTitleIndex = mlngTitleIndex + 1
End Property


'
'   機能: TitleRow 取得プロパティ
'
'   備考: なし
'
Public Property Get TitleRow(Index As Long) As Long
    TitleRow = mlngTitleRow(Index)
End Property


'
'   機能: TitleRowのIndexをデクリメント
'
'   備考: なし
'
Public Property Get TitleRowUb() As Long
    TitleRowUb = mlngTitleIndex - 1
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: TitleRowをクリア
'
'   備考: なし
'
Public Sub TitleRowClr()
    Erase mlngTitleRow
    mlngTitleIndex = 0
End Sub

