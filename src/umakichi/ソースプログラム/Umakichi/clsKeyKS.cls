VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'    騎手マスタ keyクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrKey As String


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: キー文字列 代入プロパティ
'
'   備考: なし
'
Public Property Let Str(RHS As String)
    mstrKey = RHS
End Property 'Str


'
'   機能: キー文字列 取得プロパティ
'
'   備考: なし
'
Public Property Get Str() As String
    Str = mstrKey
End Property 'Str


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: SQL WHERE 文を返す
'
'   備考: なし
'
Public Function SQLWHEREString() As String
    Dim Str As String

    Str = " WHERE [KisyuCode]='" & mstrKey & "'"

    SQLWHEREString = Str
End Function
