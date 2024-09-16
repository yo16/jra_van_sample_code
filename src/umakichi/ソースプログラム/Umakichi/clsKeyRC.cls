VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   レコードマスタ  keyクラス
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
'   機能: RecInfoKubun代入プロパティ
'
'   備考: なし
'
Public Property Let RecInfoKubun(RHS As String)
    Mid$(mstrKey, 1, 1) = RHS
End Property 'RecInfoKubun


'
'   機能: RecInfoKubun取得プロパティ
'
'   備考: なし
'
Public Property Get RecInfoKubun() As String
    RecInfoKubun = Mid$(mstrKey, 1, 1)
End Property 'RecInfoKubun


'
'   機能: Year代入プロパティ
'
'   備考: なし
'
Public Property Let Year(RHS As String)
    Mid$(mstrKey, 2, 4) = RHS
End Property 'Year


'
'   機能: Year取得プロパティ
'
'   備考: なし
'
Public Property Get Year() As String
    Year = Mid$(mstrKey, 2, 4)
End Property 'Year


'
'   機能: MonthDay代入プロパティ
'
'   備考: なし
'
Public Property Let MonthDay(RHS As String)
    Mid$(mstrKey, 6, 4) = RHS
End Property 'MonthDay


'
'   機能: MonthDay取得プロパティ
'
'   備考: なし
'
Public Property Get MonthDay() As String
    MonthDay = Mid$(mstrKey, 6, 4)
End Property 'MonthDay


'
'   機能: JyoCD代入プロパティ
'
'   備考: なし
'
Public Property Let JyoCD(RHS As String)
    Mid$(mstrKey, 10, 2) = RHS
End Property 'JyoCD


'
'   機能: JyoCD取得プロパティ
'
'   備考: なし
'
Public Property Get JyoCD() As String
    JyoCD = Mid$(mstrKey, 10, 2)
End Property 'JyoCD


'
'   機能: Kaiji代入プロパティ
'
'   備考: なし
'
Public Property Let Kaiji(RHS As String)
    Mid$(mstrKey, 12, 2) = RHS
End Property 'Kaiji


'
'   機能: Kaiji取得プロパティ
'
'   備考: なし
'
Public Property Get Kaiji() As String
    Kaiji = Mid$(mstrKey, 12, 2)
End Property 'Kaiji


'
'   機能: Nichiji代入プロパティ
'
'   備考: なし
'
Public Property Let Nichiji(RHS As String)
    Mid$(mstrKey, 14, 2) = RHS
End Property 'Nichiji


'
'   機能: Nichiji取得プロパティ
'
'   備考: なし
'
Public Property Get Nichiji() As String
    Nichiji = Mid$(mstrKey, 14, 2)
End Property 'Nichiji


'
'   機能: RaceNum代入プロパティ
'
'   備考: なし
'
Public Property Let RaceNum(RHS As String)
    Mid$(mstrKey, 16, 2) = RHS
End Property 'RaceNum


'
'   機能: RaceNum取得プロパティ
'
'   備考: なし
'
Public Property Get RaceNum() As String
    RaceNum = Mid$(mstrKey, 16, 2)
End Property 'RaceNum


'
'   機能: TokuNum_SyubetuCD代入プロパティ
'
'   備考: なし
'
Public Property Let TokuNum_SyubetuCD(RHS As String)
    Mid$(mstrKey, 18, 6) = RHS
End Property 'TokuNum_SyubetuCD


'
'   機能: TokuNum_SyubetuCD取得プロパティ
'
'   備考: なし
'
Public Property Get TokuNum_SyubetuCD() As String
    TokuNum_SyubetuCD = Mid$(mstrKey, 18, 6)
End Property 'TokuNum_SyubetuCD


'
'   機能: Kyori代入プロパティ
'
'   備考: なし
'
Public Property Let KYORI(RHS As String)
    Mid$(mstrKey, 24, 4) = RHS
End Property 'Kyori


'
'   機能: Kyori取得プロパティ
'
'   備考: なし
'
Public Property Get KYORI() As String
    KYORI = Mid$(mstrKey, 24, 4)
End Property 'Kyori


'
'   機能: TrackCD代入プロパティ
'
'   備考: なし
'
Public Property Let TrackCD(RHS As String)
    Mid$(mstrKey, 28, 2) = RHS
End Property 'TrackCD


'
'   機能: TrackCD取得プロパティ
'
'   備考: なし
'
Public Property Get TrackCD() As String
    TrackCD = Mid$(mstrKey, 28, 2)
End Property 'TrackCD


'
'   機能: キー配列を返す
'
'   備考: なし
'
Public Property Get ValueArray() As Variant
    ValueArray = Array(RecInfoKubun, Year, MonthDay, JyoCD, Kaiji, Nichiji, _
                RaceNum, TokuNum_SyubetuCD, KYORI, TrackCD)
End Property


'
'   機能: フィールド名配列を返す
'
'   備考: なし
'
Public Property Get FieldArray() As Variant
    ValueArray = Array("RecInfoKubun", "Year", "MonthDay", "JyoCD", "Kaiji", _
                "Nichiji", "RaceNum", "TokuNum_SyubetuCD", "Kyori", "TrackCD")
End Property


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

    str = " WHERE [RecInfoKubun]='" & RecInfoKubun & "'"
    str = str & " AND [Year]='" & Year & "'"
    str = str & " AND [MonthDay]='" & MonthDay & "'"
    str = str & " AND [JyoCD]='" & JyoCD & "'"
    str = str & " AND [Kaiji]='" & Kaiji & "'"
    str = str & " AND [Nichiji]='" & Nichiji & "'"
    str = str & " AND [RaceNum]='" & RaceNum & "'"
    str = str & " AND [TokuNum_SyubetuCD]='" & TokuNum_SyubetuCD & "'"
    str = str & " AND [Kyori]='" & KYORI & "'"
    str = str & " AND [TrackCD]='" & TrackCD & "'"

    SQLWHEREString = str
End Function


'
'   機能: レコードセットから文字列を返す
'
'   備考: なし
'
Public Sub SetFromRS(rs As ADODB.Recordset)
    mstrKey = rs("RecInfoKubun") & rs("Year") & rs("MonthDay") & rs("JyoCD") & _
                rs("Kaiji") & rs("Nichiji") & rs("RaceNum") & rs("TokuNum_SyubetuCD") & _
                rs("Kyori") & rs("TrackCD")
End Sub
