VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataRASel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "出馬表選択画面データクラス"
'
'   出馬表選択画面データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'イベント
Public Event FetchComplete(gd As clsGridData)   '' 取得完了イベント
Public Event NoData()                           '' データ無しイベント

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter        '' 文字列変換メソッド群
Private mFraTopStr As String


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: トップエリアの文字列を取得する
'
'   備考: なし
'
Public Property Get FraTopStr() As String
    FraTopStr = mFraTopStr
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 取得を開始する
'
'   備考: 初期表示用データの取得
'
Public Function Fetch(ByRef objKey As clsKeyRASel) As Boolean
On Error GoTo ErrorHandler
    Dim cn        As ADODB.Connection
    Dim rs        As ADODB.Recordset
    Dim gd        As clsGridData
    Dim strSQL    As String
    Dim i         As Long
    Dim p         As Long               '' 同日レースのインデックス 0~2
    Dim max       As Long
    Dim WriteFlag As Boolean
    
    
    Set gd = New clsGridData
    Set cn = gApp.GetCN_RACE()
    
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM RACE"
    strSQL = strSQL & objKey.SQLWHEREString
    strSQL = strSQL & " AND [JyoCD] <= '10' "
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & " [Year], [MonthDay], [JyoCD], [RaceNum]"
    
    Set rs = cn.Execute(strSQL, , adCmdText)
    
    If rs.EOF Then
        Fetch = False
        Exit Function
    End If
    
    gd.Rows = 30
    gd.Cols = 3
    
    mFraTopStr = mSC.YMD1(objKey.str) & mCC.YOBI7(rs("YoubiCD"))
    
    i = 1
    p = 0
    max = 1
    ' カラムヘッダに場所コードを記憶する
    gd.ItemMatrix(0, 0).Text = mCC.KIBJ3(rs("JyoCD")) _
                                    & mSC.KN1(rs("Kaiji") & rs("Nichiji"))
    gd.ItemMatrix(0, 0).key = rs("JyoCD")
    Do While Not rs.EOF
        gd.ItemMatrix(i, p).Text = GetString(rs)
        gd.ItemMatrix(i, p).Link = "RA"
        gd.ItemMatrix(i, p).key = rs("Year") & rs("MonthDay") & rs("JyoCD") & rs("Kaiji") & rs("Nichiji") & rs("RaceNum")
        gd.ItemMatrix(i, p).ToolTip = Trim$(rs("Hondai")) & mCC.GRAD2(rs("GradeCD"))
        gd.ItemMatrix(i, p).FRColor = RGB(1, 1, 1)
        
        rs.MoveNext
        If rs.EOF Then Exit Do
        ' 場所コードが異なる場合
        If rs("JyoCD") <> gd.ItemMatrix(0, p).key Then
            p = p + 1
            If p > 2 Then
                p = 2
                gApp.Log "同日に３レース以上データがあります。"
                i = i + 1
            Else
                i = 1
                gd.ItemMatrix(0, p).Text = mCC.KIBJ3(rs("JyoCD")) _
                                         & mSC.KN1(rs("Kaiji") & rs("Nichiji"))
                gd.ItemMatrix(0, p).key = rs("JyoCD")
            End If
        Else
            i = i + 1
        End If
        max = Bigger(max, i)
    Loop
    
    gd.Rows = max + 1
    
    Fetch = True
    
    RaiseEvent FetchComplete(gd)
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 文字列を取得
'
'   備考: なし
'
Private Function GetString(rs As ADODB.Recordset) As String
On Error GoTo ErrorHandler
    Dim out As String
    Dim strSQL_TC As String
    Dim strSQL_CC As String
    Dim strSQLWhereKey As String
    Dim cnTC As ADODB.Connection
    Dim cnCC As ADODB.Connection
    Dim rsTC As ADODB.Recordset
    Dim rsCC As ADODB.Recordset
    
    strSQLWhereKey = " WHERE [Year] = '" & rs("Year") & "'"
    strSQLWhereKey = strSQLWhereKey & " AND [MonthDay] = '" & rs("MonthDay") & "'"
    strSQLWhereKey = strSQLWhereKey & " AND [JyoCD] = '" & rs("JyoCD") & "'"
    strSQLWhereKey = strSQLWhereKey & " AND [Kaiji] = '" & rs("kaiji") & "'"
    strSQLWhereKey = strSQLWhereKey & " AND [Nichiji] = '" & rs("Nichiji") & "'"
    strSQLWhereKey = strSQLWhereKey & " AND [RaceNum] = '" & rs("RaceNum") & "'"
    
    Set cnTC = gApp.GetCN_HASSOU_CHANGE()
    
    strSQL_TC = "SELECT * FROM HASSOU_CHANGE"
    strSQL_TC = strSQL_TC & strSQLWhereKey
    strSQL_TC = strSQL_TC & " ORDER BY"
    strSQL_TC = strSQL_TC & " [Year], [MonthDay], [JyoCD], [RaceNum]"
    
    Set rsTC = cnTC.Execute(strSQL_TC, , adCmdText)

    Set cnCC = gApp.GetCN_COURSE_CHANGE()
    
    strSQL_CC = "SELECT * FROM COURSE_CHANGE"
    strSQL_CC = strSQL_CC & strSQLWhereKey
    strSQL_CC = strSQL_CC & " ORDER BY"
    strSQL_CC = strSQL_CC & " [Year], [MonthDay], [JyoCD], [RaceNum]"
    
    Set rsCC = cnCC.Execute(strSQL_CC, , adCmdText)
    
    out = Right$("   " & val(rs("RaceNum")), 4) & "R"
    out = out & Space(1)
    out = out & MidB_SJIS(Trim$(rs("Ryakusyo6")) & mCC.GRAD3(rs("GradeCD")) & Space(18), 1, 18)
    out = out & Space(1)
    If rs("TrackCD") >= "51" Then ' 障害の場合
        out = out & "障害" & mCC.KSSB6(rs("SyubetuCD")) & " " & mCC.KSJK4(rs("JyoKenCD5"))
    Else
        out = out & mCC.KSSB6(rs("SyubetuCD")) & " " & mCC.KSJK4(rs("JyoKenCD5"))
    End If
    out = out & vbCrLf

    If rs("DataKubun") = 9 Then
        out = out & "中止 "
    ElseIf rs("DataKubun") <= 2 And Not rsTC.EOF Then
            out = out & mSC.HHNN2(rsTC("AtoHassoTime"))
    Else
        out = out & mSC.HHNN2(rs("HassoTime"))
    End If
    
    out = out & Space(1)
    
    If rs("DataKubun") <= 2 And Not rsCC.EOF Then
            out = out & MidB_SJIS(mCC.TRCK2(rsCC("AtoTrackCD")) & Space(18), 1, 18)
            out = out & Space(1)
            out = out & mSC.FitSpaceNum(rsCC("AtoKyori"), 4) & "m"
    Else
        out = out & MidB_SJIS(mCC.TRCK2(rs("TrackCD")) & Space(18), 1, 18)
        out = out & Space(1)
        out = out & mSC.FitSpaceNum(rs("Kyori"), 4) & "m"
    End If
    
    out = out & Space(1)
    out = out & mSC.FitSpaceNum(rs("TorokuTosu"), 2) & "頭"
    
    GetString = out
    
    rsTC.Close
    cnTC.Close
    rsCC.Close
    cnCC.Close
    Set rsTC = Nothing
    Set cnTC = Nothing
    Set rsCC = Nothing
    Set cnCC = Nothing
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function

'
'   機能: 文字列を切り出す
'
'   備考: なし
'
Private Function MidB_SJIS(str As String, start As Long, length As Long) As String
    Dim bytBuf() As Byte
    bytBuf = StrConv(str, vbFromUnicode)
    MidB_SJIS = StrConv(MidB(bytBuf, start, length), vbUnicode)
End Function

'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    Set mCC = New clsCodeConverter
    Set mSC = New clsStringConverter
End Sub

'
'   機能: クラス終了イベント
'
'   備考: なし
'
Private Sub Class_Terminate()
    Set mCC = Nothing
    Set mSC = Nothing
End Sub

