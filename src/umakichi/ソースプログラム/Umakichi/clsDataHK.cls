VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataHK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "変更情報 データクラス"
'
'   変更情報 データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event FetchComplete(gd As clsGridData)
Public Event NoData()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter        '' 文字列変換メソッド群

Private WithEvents mAsyncCN_KISHU_CHANGE As ADODB.Connection
Attribute mAsyncCN_KISHU_CHANGE.VB_VarHelpID = -1
Private WithEvents mCN_TORIKESI_JYOGAI As ADODB.Connection
Attribute mCN_TORIKESI_JYOGAI.VB_VarHelpID = -1
Private WithEvents mCN_TENKO_BABA As ADODB.Connection
Attribute mCN_TENKO_BABA.VB_VarHelpID = -1
Private WithEvents mCN_HASSOU_CHANGE As ADODB.Connection
Attribute mCN_HASSOU_CHANGE.VB_VarHelpID = -1
Private WithEvents mCN_COURSE_CHANGE As ADODB.Connection
Attribute mCN_COURSE_CHANGE.VB_VarHelpID = -1

Private mRS_KISHU_CHANGE As ADODB.Recordset
Private mRS_TORIKESI_JYOGAI As ADODB.Recordset
Private mRS_TENKO_BABA As ADODB.Recordset
Private mRS_HASSOU_CHANGE As ADODB.Recordset
Private mRS_COURSE_CHANGE As ADODB.Recordset

Private mblnNowFetching As Boolean      '' 取得中フラグ
Private mblnCancelFetching As Boolean   '' 中断待ちフラグ

Private mKey As clsKeyRASel             '' キー

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: データ取得状態を返す
'
'   備考: なし
'
Public Property Get NowFetching() As Boolean
    NowFetching = mblnNowFetching
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    gApp.InitLog Me
    Set mCC = New clsCodeConverter
    Set mSC = New clsStringConverter
    
    Set mAsyncCN_KISHU_CHANGE = gApp.GetCN_KISHU_CHANGE
    Set mCN_TENKO_BABA = gApp.GetCN_TENKO_BABA
    Set mCN_TORIKESI_JYOGAI = gApp.GetCN_TORIKESI_JYOGAI
    Set mCN_HASSOU_CHANGE = gApp.GetCN_HASSOU_CHANGE
    Set mCN_COURSE_CHANGE = gApp.GetCN_COURSE_CHANGE
    
End Sub


'
'   機能: キャンセル
'
'   備考: なし
'
Public Sub CancelFetching()
On Error GoTo ErrorHandler
    If mblnNowFetching Then
        gApp.Log "Cancel"
        mblnCancelFetching = True
        
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: データを取得する
'
'   備考: なし
'
Public Sub Fetch(key As clsKeyRASel)
On Error GoTo ErrorHandler
    gApp.Log ">Fetch"
    Dim strSQL As String
    
    Set mKey = key

    mblnCancelFetching = False
    mblnNowFetching = True
    
    strSQL = "SELECT * FROM KISHU_CHANGE"
    strSQL = strSQL & " WHERE [Year]='" & mKey.Year & "'"
    strSQL = strSQL & " AND [MonthDay]='" & mKey.MonthDay & "'"
    mAsyncCN_KISHU_CHANGE.Execute strSQL, , adAsyncExecute
    
    Exit Sub
ErrorHandler:
    gApp.Log "Error@clsDataHK.Fetch"
    gApp.ErrLog
    Resume Next
    
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 騎手変更マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_KISHU_CHANGE_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    Dim strSQL As String
    
    gApp.Log ">mAsyncCN_KISHU_CHANGE_ExecuteComplete"
    
    If pError Is Nothing Then
        
    Set mRS_KISHU_CHANGE = pRecordset
    Set mRS_TENKO_BABA = New ADODB.Recordset
    Set mRS_TORIKESI_JYOGAI = New ADODB.Recordset
    Set mRS_HASSOU_CHANGE = New ADODB.Recordset
    Set mRS_COURSE_CHANGE = New ADODB.Recordset
        
    strSQL = "SELECT * FROM TENKO_BABA"
    strSQL = strSQL & " WHERE [Year]='" & mKey.Year & "'"
    strSQL = strSQL & " AND [MonthDay]='" & mKey.MonthDay & "'"
    mRS_TENKO_BABA.Open strSQL, mCN_TENKO_BABA, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    strSQL = "SELECT * FROM TORIKESI_JYOGAI"
    strSQL = strSQL & " WHERE [Year]='" & mKey.Year & "'"
    strSQL = strSQL & " AND [MonthDay]='" & mKey.MonthDay & "'"
    mRS_TORIKESI_JYOGAI.Open strSQL, mCN_TORIKESI_JYOGAI, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    strSQL = "SELECT * FROM HASSOU_CHANGE"
    strSQL = strSQL & " WHERE [Year]='" & mKey.Year & "'"
    strSQL = strSQL & " AND [MonthDay]='" & mKey.MonthDay & "'"
    mRS_HASSOU_CHANGE.Open strSQL, mCN_HASSOU_CHANGE, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    strSQL = "SELECT * FROM COURSE_CHANGE"
    strSQL = strSQL & " WHERE [Year]='" & mKey.Year & "'"
    strSQL = strSQL & " AND [MonthDay]='" & mKey.MonthDay & "'"
    mRS_COURSE_CHANGE.Open strSQL, mCN_COURSE_CHANGE, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        
        Call DirectMakeData
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    
    gApp.Log "<mAsyncCN_KISHU_CHANGE_ExecuteComplete"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: グリッド表示用データを作成する
'
'   備考: なし
'
Private Sub DirectMakeData()
    gApp.Log "makedata"
    
    Dim rs      As ADODB.Recordset
    Dim RA      As ADODB.Recordset
    Dim RAKey   As clsKeyRA
    Dim RASel   As ADODB.Recordset
    Dim RASelKey   As clsKeyRASel
    Dim gd      As clsGridData
    Dim lngCP   As Long           '' カラムポインタ
    Dim lngRP   As Long           '' ロウポインタ
    Dim i       As Long
    Dim strBABA As String
    
    Set gd = New clsGridData
    Set RAKey = New clsKeyRA
    Set RASelKey = New clsKeyRASel
    Set RA = New ADODB.Recordset
    Call OpenTableDirect(RA, gApp.GetCN_RACE, "RACE")
    
    With gd
        .Cols = 5
        .Rows = 500
        lngCP = 0
        .SetItemMatrix 0, lngCP, "日時", , ">-"
        .SetItemMatrix 0, lngCP, "開催", , "^-"
        .SetItemMatrix 0, lngCP, "発表時間", , "^-"
        .SetItemMatrix 0, lngCP, "変更種別", , "^-"
        .SetItemMatrix 0, lngCP, "変更内容", , "<-"
    End With
    lngRP = 1
    
    Set rs = mRS_KISHU_CHANGE
    Do While Not rs.EOF
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If
        lngCP = 0
        RAKey.str = rs("Year") & rs("MonthDay") & rs("JyoCD") & rs("Kaiji") & rs("Nichiji") & rs("RaceNum")
        Call SafeSeek(RA, RAKey.FieldArray, RAKey.ValueArray)
        With gd
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("Year") & rs("MonthDay")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.KN1(rs("Kaiji") & rs("Nichiji")) & mCC.KIBJ4(rs("JyoCD")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.MDHM(rs("HappyoTime"))
            .SetItemMatrix lngRP, lngCP, "騎手変更"
            .SetItemMatrix lngRP, lngCP, val(rs("RaceNum")) & "R " & _
                                        val(rs("Umaban")) & "番 " & _
                                        Trim$(rs("Bamei")) & "号 " & _
                                        mCC.KSMN1(rs("MaeMinaraiCD")) & _
                                        Trim$(rs("MaeKisyuName")) & IIf(Trim$(rs("MaeKisyuName")) = "未定", " ", "騎手 ") & _
                                        IIf(val(rs("MaeFutan")) = 0, "", Format$(val(rs("MaeFutan")) / 10, "#0.0") & "kg") & " → " & _
                                        mCC.KSMN1(rs("AtoMinaraiCD")) & _
                                        Trim$(rs("AtoKisyuName")) & IIf(Trim$(rs("AtoKisyuName")) = "未定", " ", "騎手 ") & _
                                        IIf(val(rs("AtoFutan")) = 0, "", Format$(val(rs("AtoFutan")) / 10, "#0.0") & "kg")


        End With
        lngRP = lngRP + 1
        rs.MoveNext
    Loop
    
    Set rs = mRS_TENKO_BABA
    Do While Not rs.EOF
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If
        RASelKey.str = rs("Year") & rs("MonthDay") & rs("JyoCD") & rs("Kaiji") & rs("Nichiji")
        Call SafeSeek(RA, RASelKey.FieldArray, RASelKey.ValueArray)
        ' 天候初期値は表示しない
        If rs("HenkoID") = 2 Or rs("HenkoID") = 3 Then
            lngCP = 0
            With gd
                .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("Year") & rs("MonthDay")), , , "RASel", IIf(RA.EOF, "", RASelKey.str)
                .SetItemMatrix lngRP, lngCP, mSC.KN1(rs("Kaiji") & rs("Nichiji")) & mCC.KIBJ4(rs("JyoCD")), , , "RASel", IIf(RA.EOF, "", RASelKey.str)
                .SetItemMatrix lngRP, lngCP, mSC.MDHM(rs("HappyoTime"))
                .SetItemMatrix lngRP, lngCP, mSC.WE_HKBN(rs("HenkoID"))
                Select Case rs("HenkoID")
                ' 天候
                Case "2"
                    .SetItemMatrix lngRP, lngCP, mCC.TNKO1(rs("MaeTenkoCD")) & "→" & mCC.TNKO1(rs("AtoTenkoCD"))
                ' 馬場状態
                Case "3"
                    strBABA = ""
                    If rs("MaeSibaBabaCD") <> rs("AtoSibaBabaCD") Then
                        strBABA = ""
                        If rs("MaeSibaBabaCD") = "0" Then
                            strBABA = strBABA & "芝：" & mCC.BBJT1(rs("AtoSibaBabaCD"))
                        Else
                            strBABA = strBABA & "芝：" & mCC.BBJT1(rs("MaeSibaBabaCD")) & "→" & mCC.BBJT1(rs("AtoSibaBabaCD"))
                        End If
                    End If
                    strBABA = Left$(strBABA & "　　　　　　　", 7)
                    If rs("MaeDirtBabaCD") <> rs("AtoDirtBabaCD") Then
                        If rs("MaeDirtBabaCD") = "0" Then
                            strBABA = strBABA & "　ダート：" & mCC.BBJT1(rs("AtoDirtBabaCD"))
                        Else
                            strBABA = strBABA & "　ダート：" & mCC.BBJT1(rs("MaeDirtBabaCD")) & "→" & mCC.BBJT1(rs("AtoDirtBabaCD"))
                        End If
                    End If
                    strBABA = Trim$(strBABA)
                    .SetItemMatrix lngRP, lngCP, strBABA
                End Select
            End With
            lngRP = lngRP + 1
        End If
        rs.MoveNext
    Loop
    
    Set rs = mRS_TORIKESI_JYOGAI
    Do While Not rs.EOF
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If
        RAKey.str = rs("Year") & rs("MonthDay") & rs("JyoCD") & rs("Kaiji") & rs("Nichiji") & rs("RaceNum")
        Call SafeSeek(RA, RAKey.FieldArray, RAKey.ValueArray)
        lngCP = 0
        With gd
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("Year") & rs("MonthDay")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.KN1(rs("Kaiji") & rs("Nichiji")) & mCC.KIBJ4(rs("JyoCD")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.MDHM(rs("HappyoTime"))
            .SetItemMatrix lngRP, lngCP, mSC.AV_DKBN(rs("DataKubun"))
            .SetItemMatrix lngRP, lngCP, val(rs("RaceNum")) & "R" & Right$(" " & val(rs("Umaban")), 2) & "番 " & Trim$(rs("Bamei")) & "号 " & mSC.AV_DKBN(rs("DataKubun")) & "(" & mSC.AV_JKBN(rs("JiyuKubun")) & ")"
        End With
        lngRP = lngRP + 1
        rs.MoveNext
    Loop
    
    Set rs = mRS_HASSOU_CHANGE
    Do While Not rs.EOF
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If
        RAKey.str = rs("Year") & rs("MonthDay") & rs("JyoCD") & rs("Kaiji") & rs("Nichiji") & rs("RaceNum")
        Call SafeSeek(RA, RAKey.FieldArray, RAKey.ValueArray)
        lngCP = 0
        With gd
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("Year") & rs("MonthDay")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.KN1(rs("Kaiji") & rs("Nichiji")) & mCC.KIBJ4(rs("JyoCD")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.MDHM(rs("HappyoTime"))
            .SetItemMatrix lngRP, lngCP, "発走時刻変更"
            .SetItemMatrix lngRP, lngCP, val(rs("RaceNum")) & "R " & mSC.HHNN2(rs("MaeHassoTime")) & "→" & mSC.HHNN2(rs("AtoHassoTime"))
        End With
        lngRP = lngRP + 1
        rs.MoveNext
    Loop
    
    Set rs = mRS_COURSE_CHANGE
    Do While Not rs.EOF
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If
        RAKey.str = rs("Year") & rs("MonthDay") & rs("JyoCD") & rs("Kaiji") & rs("Nichiji") & rs("RaceNum")
        Call SafeSeek(RA, RAKey.FieldArray, RAKey.ValueArray)
        lngCP = 0
        With gd
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("Year") & rs("MonthDay")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.KN1(rs("Kaiji") & rs("Nichiji")) & mCC.KIBJ4(rs("JyoCD")), , , "RA", IIf(RA.EOF, "", RAKey.str)
            .SetItemMatrix lngRP, lngCP, mSC.MDHM(rs("HappyoTime"))
            .SetItemMatrix lngRP, lngCP, "コース変更"
            .SetItemMatrix lngRP, lngCP, val(rs("RaceNum")) & "R " & mCC.TRCK2(rs("MaeTrackCD")) & ":" & val(rs("MaeKyori")) & "m → " & mCC.TRCK2(rs("AtoTrackCD")) & ":" & val(rs("AtoKyori")) & "m"
        End With
        lngRP = lngRP + 1
        rs.MoveNext
    Loop
    
    gd.Rows = lngRP
    
    If lngRP < 2 Then
        RaiseEvent NoData
    Else
        RaiseEvent FetchComplete(gd)
    End If
    
    mblnNowFetching = False

End Sub

