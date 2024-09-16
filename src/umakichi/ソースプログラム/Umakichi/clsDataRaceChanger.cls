VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataRaceChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   タイトルバンドのレースチェンジャーコンボボックスのデータを取得、保持する
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'イベント
Public Event FetchComplete(locationIndex As Long, raceIndex As Long)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mKey As clsKeyRASel             '' Fetch()の引数
Private mstrLocationName(2) As String   '' 開催場     "中山５回８日"
Private mstrRaceName(2, 11) As String   '' レース名   " 9R 有馬記念"
Private mstrRaceKey(2, 11) As String    '' レースキー "2004122606050809" (オッズ・票数画面のときは+"0" or "1")
Private mlngLocationCount As Long       '' 開催場の数
Private mlngRaceCount(2) As Long        '' 開催場毎のレース数

Private mSC As clsStringConverter
Private mCC As clsCodeConverter


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 場数 1～3
'
'   備考: empty=0
'
Public Property Get LocationCount() As Long

    LocationCount = mlngLocationCount
    
End Property

'
'   機能: 場名の表示文字列
'
'   備考: "中山５回８日"
'
Public Property Get LocationName(locationIndex) As String

    LocationName = mstrLocationName(locationIndex)
    
End Property


'
'   機能: 場毎のレース数
'
'   備考: empty=0
'
Public Property Get RaceCount(locationIndex) As Long

    RaceCount = mlngRaceCount(locationIndex)
    
End Property

'
'   機能: レース名の表示文字列
'
'   備考: " 9R 有馬記念"
'
Public Property Get RaceName(locationIndex, raceIndex) As String

    RaceName = mstrRaceName(locationIndex, raceIndex)
    
End Property

'
'   機能: リンク用のキー RaceNameの引数と対応する
'
'   備考: "2004122405060610"（オッズ・票数のとき+"0"or"1"）
'
Public Property Get RaceKey(locationIndex, raceIndex) As String
    
    RaceKey = mstrRaceKey(locationIndex, raceIndex)
    
End Property

'
'   機能: データを取得した時のキーを返す
'
'   備考: 同日判定用　未取得の場合 Nothing
'
Public Property Get Key() As clsKeyRASel

    Set Key = mKey
    
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 取得 公開関数
'
'   備考: 取得後イベント通知するとき、選択したレースキーのインデックスを渡す
'
Public Sub Fetch(Key As clsKeyRASel, blnTKMode As Boolean)
On Error GoTo Errorhandler
    Dim cn        As ADODB.Connection
    Dim rs        As ADODB.Recordset
    Dim strSQL    As String
    Dim strPreJyoCD As Long         '' 開催場CD   "06"
    Dim strPreLocation As String    '' 開催場　   "中山５回８日"
    Dim lngLocationCount As Long    '' 場のカウンタ
    Dim lngRaceCount As Long        '' レースのカウンタ
    Dim lngLocationIndex As Long    '' 選択した開催場のインデックス
    Dim lngRaceIndex As Long        '' 選択したレース番号のインデックス
    Dim strRaceKeyExt As String     '' オッズ・票数のときレースキーに追加されている"0"または"1"
    
    'オッズ・票数画面のときKeyにオッズか票数を表す"0"か"1"が追加されている、strRaceKeyExtに一時保存
    If Len(Key.str) > 16 Then               'オッズ・票数のとき
        strRaceKeyExt = Mid(Key.str, 17, 1)
    Else                                    'レースのとき
        strRaceKeyExt = ""
    End If
    
    ' データ取得元は、TKModeに応じて、特別登録場の場合は、TOKU_RACEから取得する・
    Set cn = IIf(blnTKMode, gApp.GetCN_TOKU_RACE(), gApp.GetCN_RACE)
    
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM " & IIf(blnTKMode, "TOKU_RACE", "RACE")
    strSQL = strSQL & Key.SQLWHEREString
    strSQL = strSQL & " AND [JyoCD] <= '10' "
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & " [Year], [MonthDay], [JyoCD], [RaceNum]"
    
    Set rs = cn.Execute(strSQL, , adCmdText)

    lngLocationCount = 0
    lngRaceCount = 0
    strPreJyoCD = rs("JyoCD")
    strPreLocation = FormatLocation(rs)
    Do While Not rs.EOF
        '場が変わったら場の情報を保存
        If (strPreJyoCD <> rs("JyoCD")) Then
            mstrLocationName(lngLocationCount) = strPreLocation '場を保存　"中山５回８日"
            mlngRaceCount(lngLocationCount) = lngRaceCount      '場ごとにレース数を保存
            strPreJyoCD = rs("JyoCD")                           '新しい場CDを設定
            strPreLocation = FormatLocation(rs)                 '新しいを設定   "中山５回８日"
            lngLocationCount = lngLocationCount + 1             '場をカウントアップ
            lngRaceCount = 0                                    'レース数を初期化
        End If
        'レース名を保存
        mstrRaceName(lngLocationCount, lngRaceCount) = FormatRaceName(rs)
        'レースキーを保存
        mstrRaceKey(lngLocationCount, lngRaceCount) = FormatRaceKey(rs) & strRaceKeyExt 'Extを付加
        
        If mstrRaceKey(lngLocationCount, lngRaceCount) = Key.str Then   'キーのインデックスを保存
            lngLocationIndex = lngLocationCount
            lngRaceIndex = lngRaceCount
        End If
        lngRaceCount = lngRaceCount + 1                         'レース数をカウントアップ
        rs.MoveNext
    Loop
    mstrLocationName(lngLocationCount) = strPreLocation         '最後の場を保存 "中山５回８日"
    mlngRaceCount(lngLocationCount) = lngRaceCount              '最後の場のレース数を保存
    mlngLocationCount = lngLocationCount + 1                    '場の数を保存
    Set mKey = Key                                              'キーを保存

    RaiseEvent FetchComplete(lngLocationIndex, lngRaceIndex)    'キーのインデックスを引数にしてイベント発生
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 場情報から場文字列を作成
'
'   備考: なし
'
Private Function FormatLocation(rs As Recordset) As String
On Error GoTo Errorhandler
    Dim strOut As String
    
    strOut = mCC.KIBJ3(rs("JyoCD")) '開催場
    
    If (rs("Kaiji") < "10") Then    '回
        strOut = strOut & StrConv(Right(rs("Kaiji"), 1), vbWide)    '１桁のとき全角
    Else
        strOut = strOut & rs("Kaiji")
    End If
    strOut = strOut & "回"
    
    If (rs("Nichiji") < "10") Then  '日
        strOut = strOut & StrConv(Right(rs("Nichiji"), 1), vbWide)  '１桁のとき全角
    Else
        strOut = strOut & rs("Nichiji")
    End If
    strOut = strOut & "日"
    
    FormatLocation = strOut
    Exit Function
    
Errorhandler:
    gApp.ErrLog
End Function


'
'   機能: レース名を作成する
'
'   備考: 略称があるときは略称10文字。ないときは競争種別+競争条件
'
Private Function FormatRaceName(rs As Recordset) As String
On Error GoTo Errorhandler
    Dim strOut As String
    
    strOut = mSC.FitSpaceNum(rs("RaceNum"), 2) & "R "       'レース番号（１０の位が０のときスペースに変換）
    
    If (Trim(rs("Ryakusyo10")) <> "") Then
        strOut = strOut & rs("Ryakusyo10")                  '略称
    Else
        strOut = strOut & mCC.KSSB3(rs("SyubetuCD"))        '競争種別
        strOut = strOut & mCC.KSJK4(rs("JyokenCD5"))        '最若年の競争条件
    End If
    FormatRaceName = strOut
    Exit Function

Errorhandler:
    gApp.ErrLog
End Function


'
'   機能: レースキーを作成する
'
'   備考: なし
'
Private Function FormatRaceKey(rs As Recordset) As String
On Error GoTo Errorhandler

    FormatRaceKey = rs("Year") & rs("MonthDay") & rs("JyoCD") & rs("Kaiji") & rs("Nichiji") & rs("RaceNum")
    Exit Function
    
Errorhandler:
    gApp.ErrLog
End Function

'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    
    Set mKey = New clsKeyRASel
    Set mSC = New clsStringConverter
    Set mCC = New clsCodeConverter
    
End Sub
