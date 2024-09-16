VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataRCSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "レコード一覧画面 データクラス"
'
'   レコード一覧画面 データクラス
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

Private mCC     As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC     As clsStringConverter        '' JV-Data のコード変換メソッド群
Private mKey    As clsKeyRCSel

Private WithEvents mAsyncCN_RECORD      As ADODB.Connection
Attribute mAsyncCN_RECORD.VB_VarHelpID = -1
Private WithEvents mAsyncCN_UMA         As ADODB.Connection
Attribute mAsyncCN_UMA.VB_VarHelpID = -1
Private WithEvents mAsyncCN_UMA_RACE_A  As ADODB.Connection
Attribute mAsyncCN_UMA_RACE_A.VB_VarHelpID = -1

Private mRS_RECORD      As ADODB.Recordset
Private mRS_UMA         As ADODB.Recordset
Private mRS_UMA_RACE_A  As ADODB.Recordset

Private mblnNowFetching     As Boolean      '' 取得中フラグ
Private mblnCancelFetching  As Boolean      '' 中断待ちフラグ

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
'   機能: キャンセル
'
'   備考: なし
'
Public Sub CancelFetching()
On Error GoTo ErrorHandler
    If mblnNowFetching Then
        gApp.Log "Cancel"
        mblnCancelFetching = True
        freers mRS_RECORD
        freers mRS_UMA
        freers mRS_UMA_RACE_A
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 取得を開始する
'
'   備考: なし
'
Public Sub Fetch(ByVal Key As clsKeyRCSel)
On Error GoTo ErrorHandler
    gApp.Log ">Fetch"
    Dim strSQL          As String
    Dim strCondition    As String
    
    If mblnCancelFetching Then Exit Sub
    
    Set mKey = Key
    mblnNowFetching = True

    Select Case mKey.Key
    Case 11:    '芝
        strCondition = "([TrackCD] like '1_' OR [TrackCD]='20' OR [TrackCD]='21' OR [TrackCD]='22')" _
            & " AND [RecInfoKubun]='1'" _
            & " ORDER BY Kyori ASC, RecTime ASC"
    Case 12:    'ダート
        strCondition = "([TrackCD] like '2_' AND NOT([TrackCD]='20' OR [TrackCD]='21' OR [TrackCD]='22'))" _
            & " AND [RecInfoKubun]='1'" _
            & " ORDER BY Kyori ASC, RecTime ASC"
    Case 13:    'GI競争
        strCondition = "( [RecInfoKubun]='2' AND ([RecKubun]='2' or [RecKubun]='3') )" _
            & " ORDER BY Kyori"
    Case Else
        strCondition = "[JyoCD]='" & mKey.Key & "'" _
            & " AND [RecInfoKubun]='1'" _
            & " ORDER BY Kyori ASC, Year DESC, MonthDay DESC, RaceNum DESC"
    End Select
    
    strSQL = "SELECT * FROM RECORD WHERE " & strCondition

    gApp.Log ">" & strSQL
    Set mRS_RECORD = New ADODB.Recordset
    mRS_RECORD.CursorLocation = adUseClient
    mRS_RECORD.Open strSQL, mAsyncCN_RECORD, adOpenForwardOnly, adLockReadOnly, adAsyncExecute
    
    gApp.Log "<"
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
    
    Set mAsyncCN_RECORD = gApp.GetCN_RECORD
    Set mAsyncCN_UMA = gApp.GetCN_UMA
    Set mAsyncCN_UMA_RACE_A = gApp.GetCN_UMA_RACE_A
    
    Set mRS_RECORD = New ADODB.Recordset
    Set mRS_UMA = New ADODB.Recordset
    Set mRS_UMA_RACE_A = New ADODB.Recordset
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
    Set mKey = Nothing
End Sub


'
'   機能: レコードレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_RECORD_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    If pError Is Nothing Then
        Set mRS_RECORD = pRecordset
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    Call DirectMakeData
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 競走馬レコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_UMA_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    If pError Is Nothing Then
        Set mRS_UMA = pRecordset
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 競走馬レースレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_UMA_RACE_A_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    If pError Is Nothing Then
        Set mRS_UMA_RACE_A = pRecordset
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: レコード一覧グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData()
On Error GoTo ErrorHandler
    gApp.Log "makedata"
    
    Dim KS              As ADODB.Recordset
    Dim UM              As ADODB.Recordset
    
    Dim rs_copy         As ADODB.Recordset
    Dim arrStrChecked() As String
    Dim lngCt           As Long
    Dim gd              As clsGridData
    Dim lngRP           As Long             '' ロウポインタ
    Dim strKyosoCD      As String           '' 競走種別コード
    Dim intYear         As Integer          '' 開催年
    Dim intTrackCD      As Integer          '' トラックコード
    Dim strChecked      As String           '' データの種類(例．サラブレッド系2歳)
    Dim arrStrHondai()  As String           '' GIレース競走名保存用
    Dim HondaiMinTime   As New Dictionary   '' 本題に対する最短タイムを記憶
    Dim lngHondaiIndex  As Long             '' arrStrHondaiのIndex
    Dim j               As Long
    
    Set KS = New ADODB.Recordset
    Set UM = New ADODB.Recordset
    
    Call OpenTableDirect(KS, gApp.GetCN_KISHU, "KISHU")
    Call OpenTableDirect(UM, gApp.GetCN_UMA, "UMA")
    
    Set gd = New clsGridData
    With gd
        .Cols = 11
        .Rows = 500
        .SetItemMatrix 0, 0, "距離" & vbCrLf & "メートル"
        .SetItemMatrix 0, 1, "タイム"
        .SetItemMatrix 0, 2, "馬名"
        .SetItemMatrix 0, 3, "性年齢"
        .SetItemMatrix 0, 4, "血統 (父/母)"
        .SetItemMatrix 0, 5, "重量"
        .SetItemMatrix 0, 6, "騎手"
        .SetItemMatrix 0, 7, "年月日"
        .SetItemMatrix 0, 8, "場名"
        .SetItemMatrix 0, 9, "天候"
        .SetItemMatrix 0, 10, "馬場状態"
    End With

    'キャンセル
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        Exit Sub
    End If

    'データをマーキングしながら配列へ入れる
    Set rs_copy = mRS_RECORD
    lngCt = 0
    Do Until rs_copy.EOF
    
        
        '割込みにより現在の処理をキャンセル
        DoEvents
        If mblnCancelFetching Then
            mblnNowFetching = False
            mblnCancelFetching = False
            Exit Sub
        End If

        intYear = CInt(rs_copy("Year"))
        strKyosoCD = Right$(rs_copy("TokuNum_SyubetuCD"), 2)
        intTrackCD = CInt(rs_copy("TrackCD"))

        '芝かダートの時
        If (mKey.Key = "11") Or (mKey.Key = "12") Then
            If strKyosoCD = "11" Then
                strChecked = "サラブレッド系2歳"
            ElseIf strKyosoCD <= "19" Then
                strChecked = "サラブレッド系3歳以上"
            End If
            
        ElseIf mKey.Key = "13" Then
        'GI
            strChecked = Trim(rs_copy("Hondai"))
            If rs_copy("RecKubun") = "3" Then
                strChecked = strChecked & "(参考記録)"
            End If
            '同じ本題が無かったら配列に追加
            If lngHondaiIndex = 0 Then
                ReDim Preserve arrStrHondai(lngHondaiIndex)
                arrStrHondai(lngHondaiIndex) = strChecked
                lngHondaiIndex = lngHondaiIndex + 1
            Else
                For j = 0 To lngHondaiIndex - 1
                    If arrStrHondai(j) = strChecked Then Exit For ' rs_copy("Hondai") Then Exit For
                Next j
                If j > lngHondaiIndex - 1 Then
                    ReDim Preserve arrStrHondai(lngHondaiIndex)
                    arrStrHondai(lngHondaiIndex) = strChecked
                    lngHondaiIndex = lngHondaiIndex + 1
                End If
            End If
            ' 本題の最短タイムを記録
            If HondaiMinTime.Exists(strChecked) Then
                If HondaiMinTime.item(strChecked) > rs_copy("RecTime") Then
                    HondaiMinTime.item(strChecked) = rs_copy("RecTime")
                End If
            Else
                HondaiMinTime.item(strChecked) = rs_copy("RecTime")
            End If
        Else
        '場の時
            If 10 <= intTrackCD And intTrackCD <= 22 Then
                If strKyosoCD = "11" Then
                    strChecked = "芝 サラブレッド系2歳"
                ElseIf strKyosoCD <= "19" Then
                    strChecked = "芝 サラブレッド系3歳以上"
                End If
            ElseIf 23 <= intTrackCD And intTrackCD <= 29 Then
                If strKyosoCD = "11" Then
                    strChecked = "ダート サラブレッド系2歳"
                ElseIf strKyosoCD <= "19" Then
                    strChecked = "ダート サラブレッド系3歳以上"
                End If
            End If
        End If

        '条件に合致するデータならば配列に格納する
        If strChecked <> "" Then
        
            ReDim Preserve arrStrChecked(0 To 14, lngCt)
            Call SafeSeek(UM, Array("KettoNum"), Array(rs_copy("RecUmaKettoNum1").value))
            arrStrChecked(0, lngCt) = rs_copy("Kyori")
            arrStrChecked(1, lngCt) = Format$(rs_copy("RecTime"), "@:@@.@")
            arrStrChecked(2, lngCt) = RTrim$(rs_copy("RecUmaBamei1"))
            arrStrChecked(3, lngCt) = mCC.SEIB4(rs_copy("RecUmaSexCD1")) & _
                                    mSC.ValStr(GetBarei(rs_copy))
            arrStrChecked(4, lngCt) = Trim$(IfExist(UM, "Ketto3InfoBamei1")) & vbCrLf _
                                    & Trim$(IfExist(UM, "Ketto3InfoBamei2"))
                                    'Join(GetFubo(rs_copy("RecUmaKettoNum1")), vbCrLf)
            arrStrChecked(5, lngCt) = Format$(rs_copy("RecUmaFutan1") / 10, "#0.0")
            arrStrChecked(6, lngCt) = RTrim$(rs_copy("RecUmaKisyuName1"))
            arrStrChecked(7, lngCt) = mSC.YMD2(rs_copy("Year") & rs_copy("MonthDay"))
            arrStrChecked(8, lngCt) = mCC.KIBJ3(rs_copy("JyoCD"))
            arrStrChecked(9, lngCt) = mCC.TNKO1(rs_copy("TenkoCD"))
            arrStrChecked(10, lngCt) = mSC.BabaGrid(rs_copy)
            arrStrChecked(11, lngCt) = strChecked
            'リンク用キー 騎手
            Call SafeSeek(KS, Array("KisyuCode"), Array(rs_copy("RecUmaKisyuCode1").value))
            arrStrChecked(12, lngCt) = IfExist(KS, "KisyuCode")
            ' リンク用キー レコード
            arrStrChecked(13, lngCt) = rs_copy("RecInfoKubun") & _
                                        rs_copy("Year") & _
                                        rs_copy("MonthDay") & _
                                        rs_copy("JyoCD") & _
                                        rs_copy("Kaiji") & _
                                        rs_copy("Nichiji") & _
                                        rs_copy("RaceNum") & _
                                        rs_copy("TokuNum_SyubetuCD") & _
                                        rs_copy("Kyori") & _
                                        rs_copy("TrackCD")
            'リンク用キー 競走馬
            arrStrChecked(14, lngCt) = IfExist(UM, "KettoNum")
            
            lngCt = lngCt + 1
        
        
            ' 同着1
            If val(rs_copy("RecUmaKettoNum2")) <> 0 Then
                ReDim Preserve arrStrChecked(0 To 14, lngCt)
                Call SafeSeek(UM, Array("KettoNum"), Array(rs_copy("RecUmaKettoNum2").value))
                arrStrChecked(0, lngCt) = rs_copy("Kyori")
                arrStrChecked(1, lngCt) = Format$(rs_copy("RecTime"), "@:@@.@")
                arrStrChecked(2, lngCt) = RTrim$(rs_copy("RecUmaBamei2"))
                arrStrChecked(3, lngCt) = mCC.SEIB4(rs_copy("RecUmaSexCD2")) & _
                                        mSC.ValStr(GetBarei(rs_copy))
                arrStrChecked(4, lngCt) = Trim$(IfExist(UM, "Ketto3InfoBamei1")) & vbCrLf _
                                        & Trim$(IfExist(UM, "Ketto3InfoBamei2"))
                arrStrChecked(5, lngCt) = Format$(rs_copy("RecUmaFutan2") / 10, "#0.0")
                arrStrChecked(6, lngCt) = RTrim$(rs_copy("RecUmaKisyuName2"))
                arrStrChecked(7, lngCt) = mSC.YMD2(rs_copy("Year") & rs_copy("MonthDay"))
                arrStrChecked(8, lngCt) = mCC.KIBJ3(rs_copy("JyoCD"))
                arrStrChecked(9, lngCt) = mCC.TNKO1(rs_copy("TenkoCD"))
                arrStrChecked(10, lngCt) = mSC.BabaGrid(rs_copy)
                arrStrChecked(11, lngCt) = strChecked
                'リンク用キー 騎手
                Call SafeSeek(KS, Array("KisyuCode"), Array(rs_copy("RecUmaKisyuCode2").value))
                arrStrChecked(12, lngCt) = IfExist(KS, "KisyuCode")
                ' リンク用キー レコード
                arrStrChecked(13, lngCt) = rs_copy("RecInfoKubun") & _
                                            rs_copy("Year") & _
                                            rs_copy("MonthDay") & _
                                            rs_copy("JyoCD") & _
                                            rs_copy("Kaiji") & _
                                            rs_copy("Nichiji") & _
                                            rs_copy("RaceNum") & _
                                            rs_copy("TokuNum_SyubetuCD") & _
                                            rs_copy("Kyori") & _
                                            rs_copy("TrackCD")
                'リンク用キー 競走馬
                arrStrChecked(14, lngCt) = IfExist(UM, "KettoNum")
                
                lngCt = lngCt + 1
            End If
        
        
            ' 同着2
            If val(rs_copy("RecUmaKettoNum3")) <> 0 Then
                ReDim Preserve arrStrChecked(0 To 14, lngCt)
                Call SafeSeek(UM, Array("KettoNum"), Array(rs_copy("RecUmaKettoNum3").value))
                arrStrChecked(0, lngCt) = rs_copy("Kyori")
                arrStrChecked(1, lngCt) = Format$(rs_copy("RecTime"), "@:@@.@")
                arrStrChecked(2, lngCt) = RTrim$(rs_copy("RecUmaBamei3"))
                arrStrChecked(3, lngCt) = mCC.SEIB4(rs_copy("RecUmaSexCD3")) & _
                                        mSC.ValStr(GetBarei(rs_copy))
                arrStrChecked(4, lngCt) = Trim$(IfExist(UM, "Ketto3InfoBamei1")) & vbCrLf _
                                        & Trim$(IfExist(UM, "Ketto3InfoBamei2"))
                arrStrChecked(5, lngCt) = Format$(rs_copy("RecUmaFutan3") / 10, "#0.0")
                arrStrChecked(6, lngCt) = RTrim$(rs_copy("RecUmaKisyuName3"))
                arrStrChecked(7, lngCt) = mSC.YMD2(rs_copy("Year") & rs_copy("MonthDay"))
                arrStrChecked(8, lngCt) = mCC.KIBJ3(rs_copy("JyoCD"))
                arrStrChecked(9, lngCt) = mCC.TNKO1(rs_copy("TenkoCD"))
                arrStrChecked(10, lngCt) = mSC.BabaGrid(rs_copy)
                arrStrChecked(11, lngCt) = strChecked
                'リンク用キー 騎手
                Call SafeSeek(KS, Array("KisyuCode"), Array(rs_copy("RecUmaKisyuCode3").value))
                arrStrChecked(12, lngCt) = IfExist(KS, "KisyuCode")
                ' リンク用キー レコード
                arrStrChecked(13, lngCt) = rs_copy("RecInfoKubun") & _
                                            rs_copy("Year") & _
                                            rs_copy("MonthDay") & _
                                            rs_copy("JyoCD") & _
                                            rs_copy("Kaiji") & _
                                            rs_copy("Nichiji") & _
                                            rs_copy("RaceNum") & _
                                            rs_copy("TokuNum_SyubetuCD") & _
                                            rs_copy("Kyori") & _
                                            rs_copy("TrackCD")
                'リンク用キー 競走馬
                arrStrChecked(14, lngCt) = IfExist(UM, "KettoNum")
                
                lngCt = lngCt + 1
            End If
        
            strChecked = ""
        
        End If
        rs_copy.MoveNext
    Loop
    
    rs_copy.Close
    Set rs_copy = Nothing
    
    '配列をグリッドに書き出す
    If lngCt > 0 Then
        If (mKey.Key = "11") Or (mKey.Key = "12") Then
            '芝かダートの時
            lngRP = SetArrGrid("サラブレッド系2歳", arrStrChecked, gd, 1)
            lngRP = SetArrGrid("サラブレッド系3歳以上", arrStrChecked, gd, lngRP)
        ElseIf mKey.Key = "13" Then
            'GIの時
            lngRP = 1
            For j = 0 To UBound(arrStrHondai)
                ' 参考記録は、次に引っ付けて表示
                If Right(arrStrHondai(j), 6) <> "(参考記録)" Then
                    lngRP = SetArrGrid(arrStrHondai(j), arrStrChecked, gd, lngRP, HondaiMinTime)
                    lngRP = SetArrGrid(arrStrHondai(j) & "(参考記録)", arrStrChecked, gd, lngRP, HondaiMinTime)
                End If
            Next j
        Else
            '場の時
            lngRP = SetArrGrid("芝 サラブレッド系2歳", arrStrChecked, gd, 1)
            lngRP = SetArrGrid("芝 サラブレッド系3歳以上", arrStrChecked, gd, lngRP)
            lngRP = SetArrGrid("ダート サラブレッド系2歳", arrStrChecked, gd, lngRP)
            lngRP = SetArrGrid("ダート サラブレッド系3歳以上", arrStrChecked, gd, lngRP)
        End If
    End If
    
    If lngRP < 2 Then
        RaiseEvent NoData
    Else
        gd.Rows = lngRP
        RaiseEvent FetchComplete(gd)
    End If
    
    mblnCancelFetching = False
    mblnNowFetching = False
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 指定された分類のデータを配列からグリッドに書き出す
'
'   備考: 次の行数を返す
'
Private Function SetArrGrid(strBunrui As String, arrData() As String, gd As clsGridData, lngRP As Long, Optional HondaiMinTime As Dictionary) As Long
On Error GoTo ErrorHandler
    Dim i               As Long
    Dim j               As Long
    Dim lngCt           As Long             '' 配列のインデックス
    Dim lngColorCt      As Long             '' 背景色用カウンタ
    Dim strKyori        As String
    Dim blnKyoriFlag    As Boolean
    Dim blnSkipFlag     As Boolean


    ' グループが空なら出力しない
    blnSkipFlag = True
    For lngCt = 0 To UBound(arrData, 2)
        If arrData(11, lngCt) = strBunrui Then
            blnSkipFlag = False
            Exit For
        End If
    Next lngCt
    If blnSkipFlag Then
        SetArrGrid = lngRP
        Exit Function
    End If
    
    'セルを連結して分類名を表示
    For i = 0 To 10
        gd.SetItemMatrix lngRP, j, strBunrui
        gd.ItemMatrix(lngRP, i).BGColor = RGB(128, 128, 128)
    Next i
    mKey.SetTitleRow = lngRP
    lngRP = lngRP + 1
    
    '特定のデータだけ書き出す
    For lngCt = 0 To UBound(arrData, 2)
    
        If arrData(0, lngCt) <> strKyori Then
            blnKyoriFlag = True
            strKyori = arrData(0, lngCt)
        End If
        
        blnSkipFlag = False
        If Not HondaiMinTime Is Nothing Then
            If Format$(HondaiMinTime(strBunrui), "@:@@.@") <> arrData(1, lngCt) Then
                blnSkipFlag = True
            End If
        End If
        
        If blnKyoriFlag And Not blnSkipFlag Then
            If arrData(11, lngCt) = strBunrui Then
                j = 0
                For i = 0 To 10
                    'リンクの設定のため，処理を分ける
                    If i = 0 Then       '距離
                        gd.SetItemMatrix lngRP, j, arrData(i, lngCt), , , "RC", arrData(13, lngCt)
                    ElseIf i = 2 Then       '馬名
                        gd.SetItemMatrix lngRP, j, arrData(i, lngCt), , , "UM", arrData(14, lngCt)
                    ElseIf i = 6 Then   '騎手
                        gd.SetItemMatrix lngRP, j, arrData(i, lngCt), , , "KS", arrData(12, lngCt)
                    Else
                        gd.SetItemMatrix lngRP, j, arrData(i, lngCt)
                    End If
                    'セルの背景色の設定
                    gd.ItemMatrix(lngRP, i).BGColor = IIf(lngColorCt Mod 2 = 0, RGB(240, 240, 255), RGB(223, 223, 255))
                Next i
                lngRP = lngRP + 1
                lngColorCt = lngColorCt + 1
                blnKyoriFlag = False
            End If
        End If
    Next lngCt

    SetArrGrid = lngRP
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function

        
'
'   機能: 馬毎レース情報_前半から馬齢を取得
'
'   備考: 該当データがない場合，エラーが発生した場合は"  "を返す
'
Private Function GetBarei(rc As ADODB.Recordset) As String
On Error GoTo ErrorHandler
    Dim strSQL As String
    Dim strRtrn As String
    
    Set mRS_UMA_RACE_A = New ADODB.Recordset
    strSQL = "SELECT [Barei] FROM UMA_RACE_A WHERE [Year]='" & rc("Year") & "'"
    strSQL = strSQL & " AND [MonthDay]='" & rc("MonthDay") & "'"
    strSQL = strSQL & " AND [JyoCD]='" & rc("JyoCD") & "'"
    strSQL = strSQL & " AND [Kaiji]='" & rc("Kaiji") & "'"
    strSQL = strSQL & " AND [Nichiji]='" & rc("Nichiji") & "'"
    strSQL = strSQL & " AND [RaceNum]='" & rc("RaceNum") & "'"
    strSQL = strSQL & " AND [KettoNum]='" & rc("RecUmaKettoNum1") & "'"
    mRS_UMA_RACE_A.Open strSQL, mAsyncCN_UMA_RACE_A, adOpenStatic, adLockReadOnly
    If mRS_UMA_RACE_A.RecordCount < 1 Then
        strRtrn = "  "
    Else
        strRtrn = mRS_UMA_RACE_A.Fields(0)
    End If
    mRS_UMA_RACE_A.Close
    Set mRS_UMA_RACE_A = Nothing
    GetBarei = strRtrn
    Exit Function

ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
