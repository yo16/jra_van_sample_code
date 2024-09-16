VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyHCSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   坂路調教選択 keyクラス
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
Public Property Let str(RHS As String)
    mstrKey = RHS
End Property 'Str

'
'   機能: キー文字列 取得プロパティ
'
'   備考: なし
'
Public Property Get str() As String
    str = mstrKey
End Property 'Str

'
'   機能: Year 代入プロパティ
'
'   備考: なし
'
Public Property Let Year(RHS As String)
    Mid$(mstrKey, 1, 4) = RHS
End Property 'Year

'
'   機能: Year 取得プロパティ
'
'   備考: なし
'
Public Property Get Year() As String
    Year = Mid$(mstrKey, 1, 4)
End Property 'Year


'
'   機能: MonthDay 代入プロパティ
'
'   備考: なし
'
Public Property Let MonthDay(RHS As String)
    Mid$(mstrKey, 5, 4) = RHS
End Property 'MonthDay

'
'   機能: MonthDay 取得プロパティ
'
'   備考: なし
'
Public Property Get MonthDay() As String
    MonthDay = Mid$(mstrKey, 5, 4)
End Property 'MonthDay


'
'   機能: MonthDay 代入プロパティ
'
'   備考: なし
'
Public Property Let Basho(RHS As String)
    Mid$(mstrKey, 9, 1) = RHS
End Property 'MonthDay


'
'   機能: MonthDay 取得プロパティ
'
'   備考: なし
'
Public Property Get Basho() As String
    Basho = Mid$(mstrKey, 9, 1)
End Property 'MonthDay


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: SQL WHERE 文を返す
'
'   備考: なし
'
Public Function SQLWHEREString() As String
    Dim str As String
    
    str = " WHERE [ChokyoDate]='" & Left$(mstrKey, 8) & "'"
    Select Case Me.Basho
    Case "1"
        str = str & " AND [TresenKubun]='0'"
    Case "2"
        str = str & " AND [TresenKubun]='1'"
    End Select
    SQLWHEREString = str
End Function

