VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataRAKaiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = " 出馬表開催選択画面データクラス"
'
'   出馬表開催選択画面データクラス
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
Private mGridData As clsGridData         '' グリッドデータ

Private WithEvents mAsyncCN_RACE As ADODB.Connection
Attribute mAsyncCN_RACE.VB_VarHelpID = -1
Private WithEvents mAsyncCN_SCHEDULE As ADODB.Connection
Attribute mAsyncCN_SCHEDULE.VB_VarHelpID = -1
Private WithEvents mAsyncCN_RAKaiSel As ADODB.Connection
Attribute mAsyncCN_RAKaiSel.VB_VarHelpID = -1

Private mRS_RACE As ADODB.Recordset
Private mRS_SCHEDULE As ADODB.Recordset
Private mRS_RAKaiSel As ADODB.Recordset

Private mblnNowFetching As Boolean      '' 取得中フラグ
Private mblnCancelFetching As Boolean   '' 中断待ちフラグ

Private mKey As clsKeyRAKaiSel          '' キー

Private mstrYearList() As String

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

'
'   機能: 開催年を返す
'
'   備考: なし
'
Public Property Get YearList() As String()
    YearList = mstrYearList
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


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
        
        freers mRS_RACE
        freers mRS_SCHEDULE
        
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 取得を開始する
'
'   備考: 初期表示用データの取得
'
Public Sub Fetch(ByVal key As clsKeyRAKaiSel)
On Error GoTo ErrorHandler
    gApp.Log ">Fetch"
    Dim strSQL      As String
    Dim cacheMaker  As frmDBRAKaiSel
    
    Set mKey = key
    
    mblnCancelFetching = False
    mblnNowFetching = True
    
    ' 全場所＆キャッシュが無ければ、生成する
    If key.JyoCD = "00" And Not gApp.R_RAKaiSelCacheExist(mKey.Year) Then
        Set cacheMaker = New frmDBRAKaiSel
        cacheMaker.TargetYear = mKey.Year
        cacheMaker.Show vbModal
        Unload cacheMaker
        Set cacheMaker = Nothing
    End If
    
    Set mAsyncCN_RAKaiSel = gApp.GetCN_RAKaiSel
    
    '　全場所＆キャッシュがあればキャッシュを使う
    If key.JyoCD = "00" And gApp.R_RAKaiSelCacheExist(mKey.Year) Then
        strSQL = "SELECT * FROM RAKaiSel"
        strSQL = strSQL & " WHERE [Year]='" & key.Year & "' "
        strSQL = strSQL & " ORDER BY [Year], MonthDay DESC, JyoCD"
        mAsyncCN_RAKaiSel.Execute strSQL, , adAsyncExecute
    Else
        strSQL = "SELECT * FROM RACE"
        strSQL = strSQL & " WHERE [Year]='" & key.Year & "' "
        If key.JyoCD = "00" Then
            strSQL = strSQL & " AND JyoCD>='00' AND JyoCD<='10'"
        Else
            strSQL = strSQL & " AND JyoCD='" & key.JyoCD & "' "
        End If
        strSQL = strSQL & " ORDER BY [Year], MonthDay DESC, JyoCD"
        gApp.Log ">"
        Set mRS_RACE = New ADODB.Recordset
        mRS_RACE.CursorLocation = adUseClient
        mRS_RACE.Open strSQL, mAsyncCN_RACE, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
        gApp.Log "<"
    End If
    
    Call getYearList
    gApp.Log "'''''''''' " & strSQL
    
    gApp.Log "<Fetch"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub



'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
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
    
    Set mAsyncCN_RACE = gApp.GetCN_RACE
    Set mAsyncCN_SCHEDULE = gApp.GetCN_SCHEDULE

    Set mRS_RACE = New ADODB.Recordset
    Set mRS_SCHEDULE = New ADODB.Recordset
End Sub

'
'   機能: クラス終了イベント
'
'   備考: なし
'
Private Sub Class_Terminate()
    gApp.TermLog Me
    Set mCC = Nothing
    Set mSC = Nothing
End Sub


'
'   機能: レースレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_RACE_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    Dim strSQL As String
    
    gApp.Log ">mAsyncCN_RACE_ExecuteComplete"
    
    If pError Is Nothing Then
        
        Set mRS_RACE = pRecordset
        
        strSQL = "SELECT * FROM SCHEDULE "
        strSQL = strSQL & " WHERE [Year]='" & mKey.Year & "' "
        If mKey.JyoCD <> "00" Then
            strSQL = strSQL & " AND JyoCD='" & mKey.JyoCD & "' "
        Else
            strSQL = strSQL & " AND JyoCD<='10' "
        End If
        strSQL = strSQL & " ORDER BY [Year], MonthDay DESC, JyoCD"
        
        mAsyncCN_SCHEDULE.Execute strSQL, , adAsyncExecute
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    
    gApp.Log "<mAsyncCN_RACE_ExecuteComplete"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: スケジュールレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_SCHEDULE_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    gApp.Log ">mAsyncCN_SCHEDULE_ExecuteComplete"
    
    If pError Is Nothing Then
        Set mRS_SCHEDULE = pRecordset
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    
    Call DirectMakeData
    
    gApp.Log "<mAsyncCN_SCHEDULE_ExecuteComplete"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: レース開催レコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_RAKaiSel_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    gApp.Log ">mAsyncCN_RAKaiSel_ExecuteComplete"
    
    If pError Is Nothing Then
        Set mRS_RAKaiSel = pRecordset
        Call DirectMakeDateRAKaiSel
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    
    gApp.Log "<mAsyncCN_RAKaiSel_ExecuteComplete"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 有効なグレードコードなら真
'
'   備考: なし
'
Private Function GradeTest(GradeCD As String) As Boolean
    GradeTest = (InStr("ABCFGH", GradeCD) > 0)
End Function


'
'   機能: キャッシュを使わず、RA SC から読み込んで書き込む（旧コード、場所フィルタの場合使用）
'
'   備考: なし
'
Private Sub DirectMakeData()
    gApp.Log "makedata"
    
    Dim rs  As ADODB.Recordset
    
    Dim gd As clsGridData
    Dim lngCP As Long           '' カラムポインタ
    Dim lngRP As Long           '' ロウポインタ
    Dim blnWriteFlag As Boolean '' 上書きフラグ
    Dim p As Long               '' 書き込みエリア
    Dim i As Long
    Dim J As Long
    Dim strPrevJyokenCD As String '' 比較用の前回条件コード
    
    Set gd = New clsGridData
    
    With gd
        .Cols = 7
        .Rows = 500
        lngCP = 0
        .SetItemMatrix 0, lngCP, "日付", , ">-"
        .SetItemMatrix 0, lngCP, "場所", , "^-"
        .SetItemMatrix 0, lngCP, "場所", , "^-"
        .SetItemMatrix 0, lngCP, "場所", , "^-"
        .SetItemMatrix 0, lngCP, "主要レース名", , "<-"
        .SetItemMatrix 0, lngCP, "主要レース名", , "<-"
        .SetItemMatrix 0, lngCP, "主要レース名", , "<-"
    End With
    
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        Exit Sub
    End If
    
    lngRP = 1
    blnWriteFlag = True
    Set rs = ActiveRecordset()
    Do While Not BothEOF()
    
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If
        
        If lngRP = 1 Then
            blnWriteFlag = True
        Else
            ' 日付が同じかどうか調べる
            If (gd.ItemMatrix(lngRP - 1, 0).key = rs("Year") & rs("MonthDay")) Then
                ' 日付が同じ場合は、
                lngRP = lngRP - 1 ' 前回と同じ行に戻す
                ' 場所が同じかどうか調べる
                If rs("JyoCD") = gd.ItemMatrix(lngRP, 1 + p).key Then
                    ' 場所が同じ場合は
                    ' グレードがより高いかどうか調べる
                    ' より主なレースにふさわしいかどうか調べる
                    If CompareMajorRace(rs, gd.ItemMatrix(lngRP, 4 + p).key & strPrevJyokenCD) Then
                        ' グレードが高い場合は
                        blnWriteFlag = True  ' 次回は書き込む
                    Else
                        ' グレードが低い場合は
                        blnWriteFlag = False ' 次回は書き込まない
                    End If
                Else
                    ' 場所が異なる場合は
                    p = p + 1 ' 右にずらす
                    If p >= 3 Then ' ３箇所以上もデータがあった場合
                        gApp.Log "開催が３箇所以上同時に存在します" & ":" & rs("Year") & rs("MonthDay") & mCC.KIBJ1(rs("JyoCD"))
                        p = 2 ' ３番目の場所で処理する
                    End If
                    blnWriteFlag = True ' 次回は書き込む
                End If
            Else
                ' 日付が異なる場合は
                p = 0            ' １番目の場所にセット
                blnWriteFlag = True ' 次回は書き込む
            End If

        End If
    
        If blnWriteFlag Then
            With gd
                ' 日付カラム
                With .ItemMatrix(lngRP, 0)
                    .Text = mSC.YMD1(rs("Year") & rs("MonthDay")) & mCC.YOBI7(rs("YoubiCD"))
                    .key = rs("Year") & rs("MonthDay") ' リンクの有無にかかわらず、比較用に日付を記憶する
                End With

                ' レコードがレース情報ならリンク
                If rs Is mRS_RACE Then
                    .ItemMatrix(lngRP, 0).Link = "RASel" ' リンクする
                    For i = 0 To 6
                        .ItemMatrix(lngRP, i).BGColor = IIf(lngRP Mod 2 = 0, RGB(240, 240, 255), RGB(223, 223, 255))
                    Next i
                End If

                ' 場所カラム (p=書き込みエリア)
                With .ItemMatrix(lngRP, 1 + p)
                    .Text = mCC.KIBJ3(rs("JyoCD"))
                    .ToolTip = mCC.KIBJ1(rs("JyoCD"))
                    .Link = ""             'リンクしない
                    .key = rs("JyoCD")     'リンクの有無にかかわらず、比較用に場所コードを記憶する
                End With
                
                ' 重賞名カラム (p=書き込みエリア)
                With .ItemMatrix(lngRP, 4 + p)
                    If GetRyakusyo6(rs) <> "" Then
                        .Text = Trim$(GetRyakusyo6(rs)) & Trim$(mCC.GRAD2(GetGradeCD(rs)))
                        .ToolTip = Trim$(GetHondai(rs))
                    End If
                    .Link = ""             'リンクしない
                    .key = GetGradeCD(rs)  'リンクの有無にかかわらず、比較用にグレードコードを記憶する
                    strPrevJyokenCD = GetJokenCD(rs)  ' 条件コードも付け足す
                End With
            End With
        End If

        lngRP = lngRP + 1
        If lngRP >= gd.Rows Then
            gd.Rows = gd.Rows + 100
        End If
        rs.MoveNext
        Set rs = ActiveRecordset()
    Loop
    
    gd.Rows = lngRP
    
    For i = 1 To gd.Rows - 1
        If gd.ItemMatrix(i, 0).Link = "RASel" Then
            For J = 1 To 6
                With gd.ItemMatrix(i, J)
                    .Link = "RASel"
                    .key = gd.ItemMatrix(i, 0).key
                    .FRColor = 1
                End With
            Next J
        End If
    Next i
    
    mRS_RACE.Close
    mRS_SCHEDULE.Close
    
    If lngRP < 2 Then
        RaiseEvent NoData
    Else
        RaiseEvent FetchComplete(gd)
    End If
    
    mblnNowFetching = False
End Sub

'
'   機能: 主なレースによりふさわしいかどうかの比較
'
'   備考: 左辺の方がふさわしければTrue
'
'           ① グレードコードが小さいもの
'           ② グレードコードが同じ場合は条件コードが同じもの
'           ③ グレードコード､条件コードが同じ場合はレース番号の小さいもの
'           但し､グレードコードの順番は以下のとおり
'           優先度
'           1    A G1(平地競走)
'           3　　F J・G1（障害競走）
'           3    B G2(平地競走)
'           4　　G J・G2（障害競走）
'           5    C G3(平地競走)
'           6　　H J・G3（障害競走）
'           7    D グレードのない重賞
'           8    E 重賞以外の特別競走
'           9      その他
'
Private Function CompareMajorRace(LHSrs As ADODB.Recordset, RHS As String) As Boolean
On Error GoTo ErrorHandler
    Dim lngLHSLevel As Long
    Dim lngRHSLevel As Long
    
    lngLHSLevel = LevelOfGrace(GetGradeCD(LHSrs))
    lngRHSLevel = LevelOfGrace(Left(RHS, 1))
    
    ' グレードコード比較
    If lngLHSLevel < lngRHSLevel Then
        CompareMajorRace = True
        Exit Function
    ElseIf lngLHSLevel > lngRHSLevel Then
        CompareMajorRace = False
        Exit Function
    End If

    
    lngLHSLevel = LevelOfJyokenCD(GetJokenCD(LHSrs))
    lngRHSLevel = LevelOfJyokenCD(Right(RHS, 3))
    
    '  競争条件比較
    If lngLHSLevel < lngRHSLevel Then
        CompareMajorRace = True
        Exit Function
    ElseIf lngLHSLevel > lngRHSLevel Then
        CompareMajorRace = False
        Exit Function
    End If


    ' レース番号比較は、若い順に入っているはずなのでLHSが勝ることは無い
    
    CompareMajorRace = False
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Debug.Assert False
End Function


'
'   機能: 条件コードの優先順位
'
'   備考: なし
'
Private Function LevelOfJyokenCD(JyokenCD As String) As Long
    Select Case JyokenCD
    Case "999"
        LevelOfJyokenCD = 0
    Case "001" To "100"
        LevelOfJyokenCD = 100 - CLng(JyokenCD)
    Case "703"
        LevelOfJyokenCD = 101
    Case "702"
        LevelOfJyokenCD = 102
    Case "701"
        LevelOfJyokenCD = 103
    Case "---"
        LevelOfJyokenCD = 104
    Case Else
        LevelOfJyokenCD = 105
    End Select
End Function


'
'   機能: グレードコードの優先順位
'
'   備考: なし
'
Private Function LevelOfGrace(GradeCD As String) As Long
    Select Case GradeCD
    Case "A"
        LevelOfGrace = 1
    Case "F"
        LevelOfGrace = 2
    Case "B"
        LevelOfGrace = 3
    Case "G"
        LevelOfGrace = 4
    Case "C"
        LevelOfGrace = 5
    Case "H"
        LevelOfGrace = 6
    Case "D"
        LevelOfGrace = 7
    Case "E"
        LevelOfGrace = 8
    Case Else
        LevelOfGrace = 9
    End Select
End Function


'
'   機能: 両方ともEOFかどうか
'
'   備考: なし
'
Private Function BothEOF() As Boolean
    BothEOF = (mRS_RACE.EOF And mRS_SCHEDULE.EOF)
End Function


'
'   機能: 有効なレコードセットを返す
'
'   備考: なし
'
Private Function ActiveRecordset() As ADODB.Recordset
On Error GoTo ErrorHandler
    Dim strRADate As String
    Dim strSCDate As String
        
    If Not mRS_RACE.EOF Then
        strRADate = mRS_RACE("Year") & mRS_RACE("MonthDay")
    End If
    If Not mRS_SCHEDULE.EOF Then
        strSCDate = mRS_SCHEDULE("Year") & mRS_SCHEDULE("MonthDay")
    End If
    
    If mRS_RACE.EOF Then
        Set ActiveRecordset = mRS_SCHEDULE
    ElseIf mRS_SCHEDULE.EOF Then
        Set ActiveRecordset = mRS_RACE
    ElseIf strRADate < strSCDate Then
        Set ActiveRecordset = mRS_RACE
    ElseIf strRADate > strSCDate Then
        Set ActiveRecordset = mRS_SCHEDULE
    ElseIf mRS_RACE("JyoCD") > mRS_SCHEDULE("JyoCD") Then
        Set ActiveRecordset = mRS_SCHEDULE
    ElseIf mRS_RACE("JyoCD") < mRS_SCHEDULE("JyoCD") Then
        Set ActiveRecordset = mRS_RACE
    Else
        mRS_SCHEDULE.MoveNext
        Set ActiveRecordset = mRS_RACE
    End If
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function

'
'   機能: 条件コードを取得
'
'   備考: なし
'
Private Function GetJokenCD(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetJokenCD = rs("JyokenCD5")
    ElseIf rs("RecordSpec") = "YS" Then
        GetJokenCD = "---"
    End If
End Function

'
'   機能: レース略称6文字を取得
'
'   備考: なし
'
Private Function GetRyakusyo6(rs As ADODB.Recordset) As String
    If rs Is mRS_RACE Then
        GetRyakusyo6 = rs("Ryakusyo6")
    ElseIf rs Is mRS_SCHEDULE Then
        GetRyakusyo6 = rs("Jyusyo1Ryakusyo6")
    End If
End Function

'
'   機能: レース本題を取得
'
'   備考: なし
'
Private Function GetHondai(rs As ADODB.Recordset) As String
    If rs Is mRS_RACE Then
        GetHondai = rs("Hondai")
    ElseIf rs Is mRS_SCHEDULE Then
        GetHondai = rs("Jyusyo1Hondai")
    End If
End Function

'
'   機能: グレードコードを取得
'
'   備考: なし
'
Private Function GetGradeCD(rs As ADODB.Recordset) As String
    If rs Is mRS_RACE Then
        GetGradeCD = rs("GradeCD")
    ElseIf rs Is mRS_SCHEDULE Then
        GetGradeCD = rs("Jyusyo1GradeCD")
    End If
End Function


'
'   機能: キャッシュから読み込んで書き込む
'
'   備考: なし
'
Private Sub DirectMakeDateRAKaiSel()
On Error GoTo ErrorHandler
        
    Dim rs  As ADODB.Recordset
    
    Dim gd As clsGridData
    Dim lngCP As Long           '' ロウポインタ
    Dim lngRP As Long           '' ロウポインタ
    Dim blnWriteFlag As Boolean '' 上書きフラグ
    Dim p As Long               '' 書き込みエリア
    Dim i As Long
    Dim J As Long
    
    Set gd = New clsGridData
    Set rs = mRS_RAKaiSel
    
    With gd
        .Cols = 7
        .Rows = 500
        lngCP = 0
        .SetItemMatrix 0, lngCP, "日付", , ">-"
        .SetItemMatrix 0, lngCP, "場所", , "^-"
        .SetItemMatrix 0, lngCP, "場所", , "^-"
        .SetItemMatrix 0, lngCP, "場所", , "^-"
        .SetItemMatrix 0, lngCP, "主要レース名", , "<-"
        .SetItemMatrix 0, lngCP, "主要レース名", , "<-"
        .SetItemMatrix 0, lngCP, "主要レース名", , "<-"
    End With
    
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        Exit Sub
    End If
    
    lngRP = 1
    Do While Not rs.EOF
        

        
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If
        
        lngCP = 0
        With gd
            
            If rs("CanLink") = "0" Then
                .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("Year") & rs("MonthDay")) & mCC.YOBI7(rs("YoubiCD")), , ">-"
            Else
                .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("Year") & rs("MonthDay")) & mCC.YOBI7(rs("YoubiCD")), , ">-", "RASel", rs("Year") & rs("MonthDay")
                For i = 0 To 6
                    .ItemMatrix(lngRP, i).BGColor = IIf(lngRP Mod 2 = 0, RGB(240, 240, 255), RGB(223, 223, 255))
                Next i
            End If
            .SetItemMatrix lngRP, lngCP, mCC.KIBJ3(nr(rs, "JyoCD0")), mCC.KIBJ1(nr(rs, "JyoCD0"))
            .SetItemMatrix lngRP, lngCP, mCC.KIBJ3(nr(rs, "JyoCD1")), mCC.KIBJ1(nr(rs, "JyoCD1"))
            .SetItemMatrix lngRP, lngCP, mCC.KIBJ3(nr(rs, "JyoCD2")), mCC.KIBJ1(nr(rs, "JyoCD2"))
            .SetItemMatrix lngRP, lngCP, Trim$(nr(rs, "Dai0")) & mCC.GRAD2(nr(rs, "GradeCD0")), Trim$(nr(rs, "DaiToolTip0"))
            .SetItemMatrix lngRP, lngCP, Trim$(nr(rs, "Dai1")) & mCC.GRAD2(nr(rs, "GradeCD1")), Trim$(nr(rs, "DaiToolTip1"))
            .SetItemMatrix lngRP, lngCP, Trim$(nr(rs, "Dai2")) & mCC.GRAD2(nr(rs, "GradeCD2")), Trim$(nr(rs, "DaiToolTip2"))
            
            If rs("CanLink") = "1" Then
                For i = 0 To 6
                    .ItemMatrix(lngRP, i).BGColor = IIf(lngRP Mod 2 = 0, RGB(240, 240, 255), RGB(223, 223, 255))
                Next i
            End If
        
        End With
        
        lngRP = lngRP + 1
        If lngRP >= gd.Rows Then
            gd.Rows = gd.Rows + 100
        End If
        rs.MoveNext
    Loop
    gd.Rows = lngRP
    
    mRS_RAKaiSel.Close

    For i = 0 To gd.Rows - 1
        If gd.ItemMatrix(i, 0).Link = "RASel" Then
            For J = 1 To 6
                With gd.ItemMatrix(i, J)
                    .Link = "RASel"
                    .key = gd.ItemMatrix(i, 0).key
                    .FRColor = 1
                End With
            Next J
        End If
    Next i

    If lngRP < 2 Then
        RaiseEvent NoData
    Else
        RaiseEvent FetchComplete(gd)
    End If
    
    mblnNowFetching = False
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: NULLだったら空文字を返す
'
'   備考: なし
'
Private Function nr(rs As ADODB.Recordset, field As String) As String
    If IsNull(rs(field)) Then
        nr = ""
    Else
        nr = rs(field)
    End If
End Function


'
'   機能: 存在する年度のリストを作る
'
'   備考: なし
'
Private Sub getYearList()
On Error GoTo ErrorHandler
    Dim RA As ADODB.Recordset
    Dim YS As ADODB.Recordset
    Dim y As String
    Dim i As Long
    
    ReDim mstrYearList(0 To 100)
    
    Set RA = New ADODB.Recordset
    Call OpenTableDirect(RA, gApp.GetCN_RACE, "RACE")
    Set YS = New ADODB.Recordset
    Call OpenTableDirect(YS, gApp.GetCN_SCHEDULE, "SCHEDULE")
    
    y = Year(Now) + 1
    i = 0
    Do While y > 1900
        RA.Seek Array(y), adSeekFirstEQ
        YS.Seek Array(y), adSeekFirstEQ
        If IfExist(RA, "Year") = y Then
            mstrYearList(i) = RA("Year")
            i = i + 1
        ElseIf IfExist(YS, "Year") = y Then
            mstrYearList(i) = YS("Year")
            i = i + 1
        End If
        y = y - 1
    Loop
    
    ReDim Preserve mstrYearList(0 To i - 1)
    RA.Close
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

