VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataTKKaiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = " 特別登録馬開催選択画面データクラス"
'
'   特別登録馬開催選択画面データクラス
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

Private WithEvents mAsyncCN_TOKU_RACE As ADODB.Connection
Attribute mAsyncCN_TOKU_RACE.VB_VarHelpID = -1

Private mRS_TOKU_RACE As ADODB.Recordset

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
'   機能: 開催年を取得
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
    gApp.Log "Cancel"
    mblnCancelFetching = True
    
    freers mRS_TOKU_RACE
    
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
    Dim strSQL As String
    Dim t As Single
    
    Set mKey = key
    
    mblnCancelFetching = False
    mblnNowFetching = True
    
    strSQL = "SELECT * FROM TOKU_RACE"
    strSQL = strSQL & " WHERE [Year]='" & key.Year & "' "
    If key.JyoCD <> "00" Then
        strSQL = strSQL & " AND JyoCD='" & key.JyoCD & "' "
    Else
            strSQL = strSQL & " AND JyoCD<='10' "
    End If
    strSQL = strSQL & " ORDER BY MonthDay DESC, JyoCD"
    gApp.Log ">"
    Set mRS_TOKU_RACE = New ADODB.Recordset
    mRS_TOKU_RACE.CursorLocation = adUseClient
    mRS_TOKU_RACE.Open strSQL, mAsyncCN_TOKU_RACE, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    gApp.Log "<"
    
    
    Call getYearList
    
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
'   機能: 有効なグレードコードなら真
'
'   備考: なし
'
Private Function GradeTest(GradeCD As String) As Boolean
    GradeTest = (InStr("ABCFGH", GradeCD) > 0)
End Function


'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    gApp.InitLog Me
    Set mCC = New clsCodeConverter
    Set mSC = New clsStringConverter
    
    Set mAsyncCN_TOKU_RACE = gApp.GetCN_TOKU_RACE

    Set mRS_TOKU_RACE = New ADODB.Recordset
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
'   機能: 特別登録馬レースレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_TOKU_RACE_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    Dim strSQL As String
    
    gApp.Log ">mAsyncCN_TOKU_RACE_ExecuteComplete"
    
    If pError Is Nothing Then
        
        Set mRS_TOKU_RACE = pRecordset
        
        Call DirectMakeData
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    
    gApp.Log "<mAsyncCN_TOKU_RACE_ExecuteComplete"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 特別登録馬開催一覧を作る
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
    Dim j As Long
    
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
    Set rs = mRS_TOKU_RACE
    Do While Not rs.EOF()
    
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
                    If rs("GradeCD") <= gd.ItemMatrix(lngRP, 4 + p).key _
                        And rs("GradeCD") >= "A" Or gd.ItemMatrix(lngRP, 4 + p).key = " " Then
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

                .ItemMatrix(lngRP, 0).Link = "TKSel" ' リンクする
                For i = 0 To 6
                    .ItemMatrix(lngRP, i).BGColor = IIf(lngRP Mod 2 = 0, RGB(240, 240, 255), RGB(223, 223, 255))
                Next i

                ' 場所カラム (p=書き込みエリア)
                With .ItemMatrix(lngRP, 1 + p)
                    .Text = mCC.KIBJ3(rs("JyoCD"))
                    .ToolTip = mCC.KIBJ1(rs("JyoCD"))
                    .Link = ""             'リンクしない
                    .key = rs("JyoCD")     'リンクの有無にかかわらず、比較用に場所コードを記憶する
                End With
                
                ' 重賞名カラム (p=書き込みエリア)
                If rs("Ryakusyo6") <> "" Then
                    .ItemMatrix(lngRP, 4 + p).Text = Trim$(rs("Ryakusyo6")) & Trim$(mCC.GRAD2(rs("GradeCD")))
                End If
                .ItemMatrix(lngRP, 4 + p).ToolTip = Trim$(rs("Hondai"))
                .ItemMatrix(lngRP, 4 + p).Link = ""             'リンクしない
                .ItemMatrix(lngRP, 4 + p).key = rs("GradeCD")  'リンクの有無にかかわらず、比較用にグレードコードを記憶する
            End With
        End If


        lngRP = lngRP + 1
        If lngRP >= gd.Rows Then
            gd.Rows = gd.Rows + 100
        End If
        rs.MoveNext
    Loop
    
    gd.Rows = lngRP
    For i = 1 To gd.Rows - 1
        If gd.ItemMatrix(i, 0).Link = "TKSel" Then
            For j = 1 To 6
                With gd.ItemMatrix(i, j)
                    .Link = "TKSel"
                    .key = gd.ItemMatrix(i, 0).key
                    .FRColor = 1
                End With
            Next j
        End If
    Next i
    mRS_TOKU_RACE.Close
    
    If lngRP < 2 Then
        RaiseEvent NoData
    Else
        RaiseEvent FetchComplete(gd)
    End If
    
    mblnNowFetching = False
End Sub


'
'   機能: 存在する年度のリストを作る
'
'   備考: なし
'
Private Sub getYearList()
On Error GoTo ErrorHandler
    Dim TK As ADODB.Recordset
    Dim Y As String
    Dim i As Long
    
    ReDim mstrYearList(0 To 100)
    
    Set TK = New ADODB.Recordset
    Call OpenTableDirect(TK, gApp.GetCN_TOKU_RACE, "TOKU_RACE")
    Y = Year(Now) + 1
    i = 0
    Do While Y > 1900
        TK.Seek Array(Y), adSeekFirstEQ
        If Not TK.BOF And Not TK.EOF Then
            If TK("Year") = Y Then
                mstrYearList(i) = TK("Year")
                i = i + 1
            End If
        End If
        Y = Y - 1
    Loop
    ReDim Preserve mstrYearList(0 To i - 1)
    TK.Close
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

