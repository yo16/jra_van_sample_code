VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataHCSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "坂路調教選択画面データクラス"
'
'   坂路調教選択画面データクラス
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

Private mCC As clsCodeConverter         '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter       '' 文字列変換メソッド群
Private mGridData     As clsGridData    '' グリッドデータ
Private mlngNumRecord As Long
Private mstrFetchDate As String

Private mblnNowFetching As Boolean      '' 取得中フラグ
Private mblnCancelFetching As Boolean   '' 中断待ちフラグ

Private WithEvents mAsyncCN_Hanro As ADODB.Connection
Attribute mAsyncCN_Hanro.VB_VarHelpID = -1
Private mRS_HANRO As ADODB.Recordset


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: グリッドデータを返す
'
'   備考: なし
'
Public Property Get GridDatas() As clsGridData
    Set GridDatas = mGridData
End Property

'
'   機能: レコード番号を返す
'
'   備考: なし
'
Public Property Get NumRecord() As Long
    NumRecord = mlngNumRecord
End Property

'
'   機能: 取得したデータを返す
'
'   備考: なし
'
Public Property Get FetchDate() As String
    FetchDate = mstrFetchDate
End Property

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
        
        freers mRS_HANRO
        
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
Public Function Fetch(ByRef key As clsKeyHCSel) As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String
    Dim i As Long
    Dim col As Long

    ' グリッドデータオブジェクトを生成
    Set mGridData = New clsGridData
    
    ' レコードセットを取得する
    
    strSQL = "SELECT * FROM HANRO"
    If key.str = "" Then ' キーが空の場合、データのある最新の日付を表示する
        strSQL = strSQL & " WHERE [ChokyoDate]=(SELECT MAX(ChokyoDate) FROM HANRO)"
    Else
        strSQL = strSQL & " " & key.SQLWHEREString
    End If
    strSQL = strSQL & " ORDER BY [ChokyoTime]"
    
    mblnNowFetching = True
    mAsyncCN_Hanro.Execute strSQL, , adAsyncExecute
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
 

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    Set mSC = New clsStringConverter
    Set mCC = New clsCodeConverter
    Set mAsyncCN_Hanro = gApp.GetCN_HANRO
End Sub

'
'   機能: 坂路調教マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_Hanro_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    Dim rs      As ADODB.Recordset
    Dim UM      As ADODB.Recordset
    Dim gd      As clsGridData
    Dim lngCP   As Long
    Dim i       As Long
    
    If Not pError Is Nothing Then
        Debug.Print pError.Description
        mstrFetchDate = Format$(Now, "YYYY/MM/DD")
        mblnNowFetching = False
        mblnCancelFetching = False
        RaiseEvent NoData
    ElseIf pRecordset.EOF Then
        mstrFetchDate = Format$(Now, "YYYY/MM/DD")
        mblnNowFetching = False
        mblnCancelFetching = False
        RaiseEvent NoData
    Else
        
        Set rs = pRecordset
        
        mstrFetchDate = Format$(rs("ChokyoDate").value, "@@@@/@@/@@")
        
        Set UM = New ADODB.Recordset
        UM.Index = "PrimaryKey"
        Call OpenTableDirect(UM, gApp.GetCN_UMA, "UMA")
        
        Set gd = New clsGridData
        
        With gd
            .Rows = 3000
            .Cols = 10
            i = 0
            .SetItemMatrix 0, i, "調教時刻", , ">_"
            .SetItemMatrix 0, i, "ﾄﾚｾﾝ", "トレーニングセンター", "^_"
            .SetItemMatrix 0, i, "馬名", , "<_"
            .SetItemMatrix 0, i, "４ハロン", "800m", "^_"
            .SetItemMatrix 0, i, "４ハロン", "800m", "^_"
            .SetItemMatrix 0, i, "３ハロン", "600m", "^_"
            .SetItemMatrix 0, i, "３ハロン", "600m", "^_"
            .SetItemMatrix 0, i, "２ハロン", "400m", "^_"
            .SetItemMatrix 0, i, "２ハロン", "400m", "^_"
            .SetItemMatrix 0, i, "１ハロン", "200m", "^_"
            i = 0
            .SetItemMatrix 1, i, "調教時刻", , ">_"
            .SetItemMatrix 1, i, "ﾄﾚｾﾝ", "トレーニングセンター", "^_"
            .SetItemMatrix 1, i, "馬名", , "<_"
            .SetItemMatrix 1, i, "合計", "0m～800m", ">_"
            .SetItemMatrix 1, i, "ﾗｯﾌﾟﾀｲﾑ", "600m～800m", ">_"
            .SetItemMatrix 1, i, "合計", "0m～600m", ">_"
            .SetItemMatrix 1, i, "ﾗｯﾌﾟﾀｲﾑ", "400m～600m", ">_"
            .SetItemMatrix 1, i, "合計", "0m～400m", ">_"
            .SetItemMatrix 1, i, "ﾗｯﾌﾟﾀｲﾑ", "200m～400m", ">_"
            .SetItemMatrix 1, i, "ﾗｯﾌﾟﾀｲﾑ", "0m～200m", ">_"
        End With
    
        mstrFetchDate = mSC.YMD3(rs("ChokyoDate"))
    
        i = 2
        mlngNumRecord = 0
        Do While Not rs.EOF
        
            ' バックグラウンド処理
            DoEvents
            If mblnCancelFetching Then
                mblnCancelFetching = False
                mblnNowFetching = False
                rs.Close
                Exit Sub
            End If
            
            Call SafeSeek(UM, Array("KettoNum"), Array(rs("KettoNum").value))
            
            mlngNumRecord = mlngNumRecord + 1
            With gd
                lngCP = 0  ' 調教時間
                .ItemMatrix(i, lngCP).Text = mSC.HHNN1(rs("ChokyoTime"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
                lngCP = lngCP + 1  ' ﾄﾚｾﾝ区分
                .ItemMatrix(i, lngCP).Text = mSC.HC_TKBN(rs("TresenKubun"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignCenterTop
                lngCP = lngCP + 1  ' 馬名
                If Not UM.EOF And Not UM.BOF Then
                    .ItemMatrix(i, lngCP).Text = Trim$(UM("Bamei"))
                    .ItemMatrix(i, lngCP).Link = "UM"
                    .ItemMatrix(i, lngCP).key = UM("KettoNum")
                Else
                    .ItemMatrix(i, lngCP).Text = Trim$(rs("KettoNum"))
                    .ItemMatrix(i, lngCP).ToolTip = "この馬の詳細情報はありません"
                End If
                .ItemMatrix(i, lngCP).Alignment = flexAlignLeftTop
                lngCP = lngCP + 1  ' 4h 合計
                .ItemMatrix(i, lngCP).Text = mSC.SSSs(rs("HaronTime4"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
                lngCP = lngCP + 1  ' 4h ﾗｯﾌﾟﾀｲﾑ
                .ItemMatrix(i, lngCP).Text = mSC.SSSs(rs("LapTime4"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
                lngCP = lngCP + 1  ' 3h 合計
                .ItemMatrix(i, lngCP).Text = mSC.SSSs(rs("HaronTime3"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
                lngCP = lngCP + 1  ' 3h ﾗｯﾌﾟﾀｲﾑ
                .ItemMatrix(i, lngCP).Text = mSC.SSSs(rs("LapTime3"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
                lngCP = lngCP + 1  ' 2h 合計
                .ItemMatrix(i, lngCP).Text = mSC.SSSs(rs("HaronTime2"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
                lngCP = lngCP + 1  ' 2h ﾗｯﾌﾟﾀｲﾑ
                .ItemMatrix(i, lngCP).Text = mSC.SSSs(rs("LapTime2"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
                lngCP = lngCP + 1  ' 1h ﾗｯﾌﾟﾀｲﾑ
                .ItemMatrix(i, lngCP).Text = mSC.SSSs(rs("LapTime1"))
                .ItemMatrix(i, lngCP).Alignment = flexAlignRightTop
            End With
            i = i + 1
            If gd.Rows <= i Then
                gd.Rows = gd.Rows + 100
            End If
            rs.MoveNext
        Loop
        gd.Rows = i
        RaiseEvent FetchComplete(gd)
        mblnNowFetching = False
        mblnCancelFetching = False
    End If

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
