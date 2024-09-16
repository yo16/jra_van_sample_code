VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   出馬表  keyクラス
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
'   機能: JyoCD 代入プロパティ
'
'   備考: なし
'
Public Property Let JyoCD(RHS As String)
    Mid$(mstrKey, 9, 2) = RHS
End Property 'JyoCD


'
'   機能: JyoCD 取得プロパティ
'
'   備考: なし
'
Public Property Get JyoCD() As String
    JyoCD = Mid$(mstrKey, 9, 2)
End Property 'JyoCD


'
'   機能: Kaiji 代入プロパティ
'
'   備考: なし
'
Public Property Let Kaiji(RHS As String)
    Mid$(mstrKey, 11, 2) = RHS
End Property 'Kaiji


'
'   機能: Kaiji 取得プロパティ
'
'   備考: なし
'
Public Property Get Kaiji() As String
    Kaiji = Mid$(mstrKey, 11, 2)
End Property 'Kaiji


'
'   機能: Nichiji 代入プロパティ
'
'   備考: なし
'
Public Property Let Nichiji(RHS As String)
    Mid$(mstrKey, 13, 2) = RHS
End Property 'Nichiji


'
'   機能: Nichiji 取得プロパティ
'
'   備考: なし
'
Public Property Get Nichiji() As String
    Nichiji = Mid$(mstrKey, 13, 2)
End Property 'Nichiji


'
'   機能: RaceNum 代入プロパティ
'
'   備考: なし
'
Public Property Let RaceNum(RHS As String)
    Mid$(mstrKey, 15, 2) = RHS
End Property 'RaceNum


'
'   機能: RaceNum 取得プロパティ
'
'   備考: なし
'
Public Property Get RaceNum() As String
    RaceNum = Mid$(mstrKey, 15, 2)
End Property 'RaceNum


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
    
    str = " WHERE [Year]='" & Year & "'"
    str = str & " AND [MonthDay]='" & MonthDay & "'"
    str = str & " AND [JyoCD]='" & JyoCD & "'"
    str = str & " AND [Kaiji]='" & Kaiji & "'"
    str = str & " AND [Nichiji]='" & Nichiji & "'"
    str = str & " AND [RaceNum]='" & RaceNum & "'"
    
    SQLWHEREString = str
End Function


'
'   機能: 可読文字列を返す
'
'   備考: なし
'
Public Function ReadableString() As String
    Dim sc As New clsStringConverter
    
    ReadableString = sc.YMD1(Mid$(mstrKey, 1, 8))
End Function


'
'   機能: フィールド名を配列で返す
'
'   備考: なし
'
Public Function FieldArray() As Variant
    FieldArray = Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum")
End Function


'
'   機能: 値を配列で返す
'
'   備考: なし
'
Public Function ValueArray() As Variant
    With Me
        ValueArray = Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
End Function

