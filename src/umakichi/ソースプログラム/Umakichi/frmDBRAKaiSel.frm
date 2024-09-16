VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBRAKaiSel 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "開催一覧の作成中"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   4680
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Timer tmrTrigger 
      Enabled         =   0   'False
      Left            =   3780
      Top             =   -30
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmDBRAKaiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   開催情報作成 ダイアログ
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mRTDates As Dictionary

Private mTargetYear As String

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 開催年を設定する
'
'   備考: なし
'
Public Property Let TargetYear(RHS As String)
    mTargetYear = RHS
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 開催情報作成処理
'
'   備考: なし
'
Private Function MakeRAKaiSel() As Boolean
On Error GoTo ErrorHandler
    gApp.Log "MakeRAKaiSel"
    
    Dim CN    As ADODB.Connection
    Dim rsRA  As ADODB.Recordset    '' RACE レコードセット
    Dim rsSC  As ADODB.Recordset    '' SCHEDULE レコードセット
    Dim rsOut As ADODB.Recordset    '' 出力先
    Dim rs    As ADODB.Recordset    '' カレントレコードセット
    
    Dim cc As clsCodeConverter
    
    Dim gd As clsGridData
    Dim lngCP As Long               '' カラムポインタ
    Dim lngRP As Long               '' ロウポインタ
    Dim blnWriteFlag As Boolean     '' 上書きフラグ
    Dim blnNewRecordFlag As Boolean '' 新レコードフラグ
    Dim p As Long                   '' 書き込みエリア
    Dim i As Long
    Dim strPrevJyokenCD As String '' 比較用の前回条件コード
    
    Dim strRTDate As String
    
    Set CN = New ADODB.Connection
    
    Set rsRA = New ADODB.Recordset
    Set rsSC = New ADODB.Recordset
    Set rsOut = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set cc = New clsCodeConverter
    Set mRTDates = New Dictionary
    
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & gApp.R_DBPath & "\" & gstrMDBName(47)
    On Error Resume Next
    CN.Execute "DELETE * FROM RAKaiSel WHERE [Year]='" & mTargetYear & "'", , adExecuteNoRecords
    If Err.Number <> 0 Then
        gApp.ErrLog
        gApp.Log "RAKaiSel用先読みテーブルの削除エラー"
        MakeRAKaiSel = False
        Exit Function
    End If
    On Error GoTo ErrorHandler

    rsRA.Open "SELECT * FROM RACE WHERE [Year] = '" & mTargetYear & "' ORDER BY MonthDay, JyoCD, Kaiji, Nichiji, RaceNum", _
                gApp.GetCN_RACE, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsSC.Open "SELECT * FROM SCHEDULE WHERE [Year] = '" & mTargetYear & "' ORDER BY MonthDay, JyoCD, Kaiji, Nichiji", _
                gApp.GetCN_SCHEDULE, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsOut.Open "RAKaiSel", CN, adOpenKeyset, adLockOptimistic, adCmdTable
    
    lngRP = 1
    blnWriteFlag = True
    
    Set rs = ActiveRecordset(rsRA, rsSC)
    Do While Not (rsRA.EOF And rsSC.EOF)
    
        DoEvents
        
        If rs("JyoCD") > "10" Then  ' 中央競馬以外は書き込まない
            blnNewRecordFlag = False
            blnWriteFlag = False
        ElseIf rsOut.BOF Then       ' 最初は書き込む
            blnNewRecordFlag = True
            blnWriteFlag = True
        Else                        ' 二回目以降は、比較して書き込むか決める
            ' 日付が同じか調べる
            If rsOut("Year") = rs("Year") And rsOut("MonthDay") = rs("MonthDay") Then
                ' 日付が同じ場合は、
                blnNewRecordFlag = False
                ' 場所が同じか調べる
                If rs("JyoCD") = rsOut("JyoCD" & p) Then
                    ' 場所が同じ場合は
                    ' グレードがより高いか
                    ' より主要レースか調べる
                    If CompareMajorRace(rs, rsOut("GradeCD" & p).value & strPrevJyokenCD) Then
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
                        gApp.Log "開催が３箇所以上同時に存在します" & ":" & rs("Year") & rs("MonthDay") & rs("JyoCD")
                        p = 2 ' ３番目の場所で処理する
                    End If
                    blnWriteFlag = True ' 次回は書き込む
                End If
            Else
                ' 日付が異なる場合は
                p = 0            ' １番目の場所にセット
                blnNewRecordFlag = True
                blnWriteFlag = True ' 次回は書き込む
            End If

        End If
        
        If blnNewRecordFlag Then  ' 新しい行に書き込む場合
            Call rsOut.AddNew
            prgBar.value = 100 * (Left$(rs("MonthDay"), 2)) / 12
        End If
    
        If blnWriteFlag Then
            With gd
                ' 日付カラム
                rsOut("Year") = rs("Year")
                rsOut("MonthDay") = rs("MonthDay")
                rsOut("JyoCD") = "00"
                rsOut("YoubiCD") = rs("YoubiCD")

                ' レコードがレース情報ならリンク、開催スケジュールならリンクしない
                If rs("RecordSpec") = "RA" Then
                    rsOut("CanLink") = "1" ' リンクする
                Else
                    If IsNull(rsOut("CanLink")) Then  ' リンク設定を消さない
                        rsOut("CanLink") = "0" ' リンクしない
                    End If
                End If

                ' 場所カラム (p=書き込みエリア)
                rsOut("JyoCD" & p) = rs("JyoCD")
                
                ' 重賞名カラム (p=書き込みエリア)
                If GetRyakusyo6(rs) <> "" Then
                    rsOut("Dai" & p) = GetRyakusyo6(rs)
                    rsOut("DaiToolTip" & p) = GetHondai(rs)
                    rsOut("GradeCD" & p) = GetGradeCD(rs)
                    strPrevJyokenCD = GetJokenCD(rs)
                End If
                    
            End With
        End If
                
        ' 速報取得用の必要日付コレクションの作成
        strRTDate = rs("Year") & rs("MonthDay")
        If rs("RecordSpec") = "RA" And rs("DataKubun") <= "6" And Not mRTDates.Exists(strRTDate) Then
            mRTDates.Add strRTDate, strRTDate
        End If
        
        rs.MoveNext
        Set rs = ActiveRecordset(rsRA, rsSC)
        
    Loop
    
    If Not rsOut.EOF Then
        rsOut.Update
    End If
    
    rsRA.Close
    rsSC.Close
    MakeRAKaiSel = True
    
    ' 速報取得用の必要日付コレクションを記録する
    If Join(mRTDates.Keys, ",") <> "" Then
        gApp.R_RTDates = Join(mRTDates.Keys, ",")
    End If
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    MakeRAKaiSel = False
End Function

'
'   機能: 主要レースの判定
'
'   備考: 主なレースによりふさわしいかの比較
'         左辺の方がふさわしければTrue
'
'       ① グレードコードが小さいもの
'       ② グレードコードが同じ場合は条件コードが同じもの
'       ③ グレードコード､条件コードが同じ場合はレース番号の小さいもの
'       但し､グレードコードの順番は以下のとおり
'       優先度
'       1    A G1(平地競走)
'       3　　F J・G1（障害競走）
'       3    B G2(平地競走)
'       4　　G J・G2（障害競走）
'       5    C G3(平地競走)
'       6　　H J・G3（障害競走）
'       7    D グレードのない重賞
'       8    E 重賞以外の特別競走
'       9      その他
'
'       競争条件コード
'       9 701 新馬
'       8 702 未出走
'       7 703 未勝利
'       6 001 １００万円以下
'       5 002 ２００万円以下
'       4 003 ３００万円以下
'        .          .
'        .          .
'        .          .
'       3 099 ９９００万円以下
'       2 100 １億円以下
'       1 999 オープン
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
'   機能: 有効なレコードセットを返す
'
'   備考: なし
'
Private Function ActiveRecordset(RA As ADODB.Recordset, sc As ADODB.Recordset) As ADODB.Recordset
On Error GoTo ErrorHandler
    Dim strRADate As String
    Dim strSCDate As String
    
    If Not RA.EOF Then
        strRADate = RA("Year") & RA("MonthDay")
    End If
    If Not sc.EOF Then
        strSCDate = sc("Year") & sc("MonthDay")
    End If
    
    If RA.EOF Then
        Set ActiveRecordset = sc
    ElseIf sc.EOF Then
        Set ActiveRecordset = RA
    ElseIf strRADate < strSCDate Then
        Set ActiveRecordset = RA
    ElseIf strRADate > strSCDate Then
        Set ActiveRecordset = sc
    ElseIf RA("JyoCD") > sc("JyoCD") Then
        Set ActiveRecordset = sc
    ElseIf RA("JyoCD") < sc("JyoCD") Then
        Set ActiveRecordset = RA
    Else
        sc.MoveNext
        Set ActiveRecordset = RA
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
'   機能: グレードコードを取得
'
'   備考: なし
'
Private Function GetGradeCD(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetGradeCD = rs("GradeCD")
    ElseIf rs("RecordSpec") = "YS" Then
        GetGradeCD = rs("Jyusyo1GradeCD")
    End If
End Function


'
'   機能: 略称6を取得
'
'   備考: なし
'
Private Function GetRyakusyo6(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetRyakusyo6 = rs("Ryakusyo6")
    ElseIf rs("RecordSpec") = "YS" Then
        GetRyakusyo6 = rs("Jyusyo1Ryakusyo6")
    End If
End Function


'
'   機能: 本題を取得
'
'   備考: なし
'
Private Function GetHondai(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetHondai = rs("Hondai")
    ElseIf rs("RecordSpec") = "YS" Then
        GetHondai = rs("Jyusyo1Hondai")
    End If
End Function


'
'   機能: フォームロードイベント
'
'   備考: なし
'
Private Sub Form_Load()
On Error GoTo ErrorHandler
    prgBar.max = 100
    prgBar.Min = 0
    tmrTrigger.Interval = 100
    tmrTrigger.Enabled = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: トリガータイマーイベント
'
'   備考: なし
'
Private Sub tmrTrigger_Timer()
On Error GoTo ErrorHandler
    tmrTrigger.Enabled = False
    gApp.R_RAKaiSelCacheExist(mTargetYear) = MakeRAKaiSel
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
