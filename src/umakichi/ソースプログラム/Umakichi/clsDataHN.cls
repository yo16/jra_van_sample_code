VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataHN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "繁殖馬 データクラス"
'
'   繁殖馬 データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event NoData()
Public Event NoDataSANKU()
Public Event NoDataHANSYOKU()

Public Event FetchCompleteSANKU(gd As clsGridData)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter          '' JV-Data のコード変換メソッド群
' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 3) As String     '' テキストデータ
Private mstrLinkLabels(0 To 2) As clsGridItem     '' リンクデータ
'モジュール内で使用する為のデータ格納変数
Private mKey As clsKeyHN                '' キー
Private mBuf_HN As JV_HN_HANSYOKU

'コネクション
Private WithEvents mAsyncCN_SANKU As ADODB.Connection
Attribute mAsyncCN_SANKU.VB_VarHelpID = -1
Private mCN_HANSYOKU As ADODB.Connection
Private mCN_UMA As ADODB.Connection
'レコードセット
Private mRS_SANKU As ADODB.Recordset
Private mRS_UMA As ADODB.Recordset
Private mRS_HANSYOKU As ADODB.Recordset

Private mblnCancelFetching As Boolean


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ラベル用文字列を返す
'
'   備考: なし
'
Public Property Get Labels(Index As Integer) As String
    Labels = mstrLabels(Index)
End Property

'
'   機能: リンクラベル用文字列を返す
'
'   備考: なし
'
Public Property Get LinkLabels(Index As Integer) As clsGridItem
    Set LinkLabels = mstrLinkLabels(Index)
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データを取得する
'
'   備考: なし
'
Public Function Fetch(ByRef key As clsKeyHN) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set mKey = key
    
    Set cn = gApp.GetCN_HANSYOKU
    
    'HANSYOKU (繁殖馬マスタ)
    strSQL = "SELECT * FROM HANSYOKU " & mKey.SQLWHEREString

    Set rs = cn.Execute(strSQL)
    
    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        RaiseEvent NoData
        Exit Function
    End If
    
    Call SetDataFromRS_HN(rs, mBuf_HN)
    
    rs.Close
    Set rs = Nothing
    
    Call Set_HANSYOKU
    
    Set mAsyncCN_SANKU = gApp.GetCN_SANKU
    strSQL = "SELECT * FROM SANKU WHERE "
    If mBuf_HN.SexCD = "1" Then 'オス馬
        strSQL = strSQL & "(SANKU.FNum)='" & mKey.str & "' "
    ElseIf mBuf_HN.SexCD = "2" Then 'メス馬
        strSQL = strSQL & "(SANKU.MNum)='" & mKey.str & "' "
    Else '"0","3"
        strSQL = strSQL & "(SANKU.FNum)='" & mKey.str & "' OR "
        strSQL = strSQL & "(SANKU.MNum)='" & mKey.str & "' "
    End If
    mAsyncCN_SANKU.Execute strSQL, , adAsyncExecute
    
    gApp.Log "HN Fetch Complete"
    
    Fetch = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
    

'
'   機能: キャンセル
'
'   備考: なし
'
Public Sub CancelFetching()
On Error GoTo ErrorHandler
    mblnCancelFetching = True
    
    Call freers(mRS_SANKU)

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


'
'   機能: 産駒マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_SANKU_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    
    If pError Is Nothing Then
        Set mRS_SANKU = pRecordset
        
        If mRS_SANKU.EOF Then
            RaiseEvent NoDataSANKU
        Else
            Call Set_SANKU
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataSANKU
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 繁殖馬グリッドを作る
'
'   備考: なし
'
Private Sub Set_HANSYOKU()
On Error GoTo ErrorHandler
    Dim i As Long
    Dim HN As ADODB.Recordset
    
    For i = 0 To 2
        Set mstrLinkLabels(i) = New clsGridItem
    Next i
    
    With mBuf_HN
        
        ' 作成年月日
        With .head.MakeDate
            mstrLabels(0) = mstrLabels(0) & mSC.YMD1(.Year & .Month & .Day) & "作成データ"
        End With
        
        '競走馬テーブルを開く
        Set mCN_UMA = gApp.GetCN_UMA
        Set mRS_UMA = New ADODB.Recordset
        Call OpenTableDirect(mRS_UMA, mCN_UMA, "UMA")
    
        ' 太字レース名等グレーバー用文字列生成
        mstrLinkLabels(0).Text = Trim$(.BAMEI)
        mstrLinkLabels(0).Link = "UM"
        Call SafeSeek(mRS_UMA, Array("KettoNum"), Array(.KettoNum))
        If Not (mRS_UMA.EOF Or mRS_UMA.BOF) Then
            mstrLinkLabels(0).key = .KettoNum
        End If
        mstrLinkLabels(0).ToolTip = ""
        
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & IIf(Trim$(.SanchiName) = "", "", "産地：" & Trim$(.SanchiName)) & vbCrLf
        mstrLabels(1) = mstrLabels(1) & IIf(val(.BirthYear) = 0, "　　　    ", "生年：" & .BirthYear)
        mstrLabels(1) = mstrLabels(1) & IIf(val(.ImportYear) = 0, "", "　　輸入年：" & .ImportYear)
        mstrLabels(1) = mstrLabels(1) & vbCrLf
        mstrLabels(1) = mstrLabels(1) & mCC.HINS2(.HinsyuCD) & " "  '品種
        mstrLabels(1) = mstrLabels(1) & mCC.SEIB4(.SexCD) & " "  '性別
        mstrLabels(1) = mstrLabels(1) & mCC.KEIR1(.KeiroCD) & " "  '毛色
        mstrLabels(1) = mstrLabels(1) & mSC.HN_HansyokuMochiKubun(.HansyokuMochiKubun) & " " '繁殖馬持込区分
        mstrLabels(1) = mstrLabels(1) & mSC.HN_DelKubun(.DelKubun) & " " '

        'HANSYOKUを開く
        Set HN = New ADODB.Recordset
        Call OpenTableDirect(HN, gApp.GetCN_HANSYOKU, "HANSYOKU")
        

        mstrLabels(2) = ""
        Call SafeSeek(HN, Array("HansyokuNum"), Array(.HansyokuFNum))
        If Not HN.EOF Then
            mstrLabels(2) = mstrLabels(2) & "父："
            mstrLinkLabels(1).Text = Trim$(HN("Bamei"))
            mstrLinkLabels(1).Link = "HN"
            mstrLinkLabels(1).key = .HansyokuFNum
            mstrLinkLabels(1).ToolTip = ""
        Else
            If Not val(.HansyokuFNum) = 0 Then
                mstrLabels(2) = mstrLabels(2) & "父："
                mstrLinkLabels(1).Text = .HansyokuFNum
            End If
        End If
        mstrLabels(2) = mstrLabels(2) & vbCrLf
        Call SafeSeek(HN, Array("HansyokuNum"), Array(.HansyokuMNum))
        If Not HN.EOF Then
            mstrLabels(2) = mstrLabels(2) & "母："
            mstrLinkLabels(2).Text = Trim$(HN("Bamei"))
            mstrLinkLabels(2).Link = "HN"
            mstrLinkLabels(2).key = .HansyokuMNum
            mstrLinkLabels(2).ToolTip = ""
        Else
            If Not val(.HansyokuMNum) = 0 Then
                mstrLabels(2) = mstrLabels(2) & "母："
                mstrLinkLabels(2).Text = .HansyokuMNum
            End If
        End If
        
        Set mRS_HANSYOKU = HN
        
    End With
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 産駒グリッドを作る
'
'   備考: なし
'
Private Sub Set_SANKU()
On Error GoTo ErrorHandler

    Dim gd      As clsGridData
    Dim rs      As ADODB.Recordset
    Dim HN      As ADODB.Recordset
    Dim BR      As ADODB.Recordset
    Dim UM      As ADODB.Recordset
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim strTmp As String
        
    Set BR = New ADODB.Recordset
    Call OpenTableDirect(BR, gApp.GetCN_SEISAN, "SEISAN")
    
    Set mCN_UMA = gApp.GetCN_UMA
    Set UM = New ADODB.Recordset
    Call OpenTableDirect(UM, mCN_UMA, "UMA")
    
    Set HN = mRS_HANSYOKU
    Set rs = mRS_SANKU
        
    ' 総合成績　タブ
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData

    With gd
        .Rows = 10
        .Cols = 9
        lngRP = 0
        lngCP = 0
        .SetItemMatrix lngRP, lngCP, "生年月日"
        .SetItemMatrix lngRP, lngCP, "馬名"
        .SetItemMatrix lngRP, lngCP, "性"
        .SetItemMatrix lngRP, lngCP, "品種"
        .SetItemMatrix lngRP, lngCP, "毛色"
        .SetItemMatrix lngRP, lngCP, "生産者名"
        .SetItemMatrix lngRP, lngCP, "産地名"
        .SetItemMatrix lngRP, lngCP, "輸入年"
        .SetItemMatrix lngRP, lngCP, "持込区分"
    End With
    rs.MoveFirst
    
    lngRP = 1
    Do While Not rs.EOF
        ' バックグラウンド
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        lngCP = 0
        With gd
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("BirthDate"))  '生年月日
            Call SafeSeek(UM, Array("KettoNum"), Array(rs("KettoNum").value))
            If UM.EOF Or UM.BOF Then
                Call SafeSeek(HN, Array("KettoNum"), Array(rs("KettoNum").value))
                If HN.EOF Or HN.BOF Then
                    .SetItemMatrix lngRP, lngCP, rs("KettoNum"), "", "<-", "SK", rs("KettoNum")  '"馬名"
                Else
                    .SetItemMatrix lngRP, lngCP, Trim$(HN("Bamei")), "", "<-", "SK", rs("KettoNum") '"馬名"
                End If
            Else
                .SetItemMatrix lngRP, lngCP, Trim$(UM("Bamei")), "", "<-", "UM", rs("KettoNum") '"馬名"
            End If
            .SetItemMatrix lngRP, lngCP, mCC.SEIB4(rs("SexCD")), mCC.SEIB3(rs("SexCD")) '性
            .SetItemMatrix lngRP, lngCP, mCC.HINS2(rs("HinsyuCD")), mCC.HINS3(rs("HinsyuCD")) '"品種"
            .SetItemMatrix lngRP, lngCP, mCC.KEIR1(rs("KeiroCD"))   '"毛色"
            Call SafeSeek(BR, Array("BreederCode"), Array(rs("BreederCode").value))
            If Not BR.EOF Then
                .SetItemMatrix lngRP, lngCP, Trim$(BR("BreederName")), , , "BR", BR("BreederCode") '"生産者名"
            Else
                .SetItemMatrix lngRP, lngCP, ""
            End If
            .SetItemMatrix lngRP, lngCP, Trim$(rs("SanchiName"))     '"産地名"
            .SetItemMatrix lngRP, lngCP, IfBe(rs("ImportYear"))      '"輸入年"
            .SetItemMatrix lngRP, lngCP, mSC.HN_HansyokuMochiKubun(rs("SankuMochiKubun")) '"持込区分"
        End With
        rs.MoveNext
        lngRP = lngRP + 1
        If lngRP >= gd.Rows Then
            gd.Rows = gd.Rows + 10
        End If
    Loop
    rs.Close
    HN.Close
    BR.Close
    
    gd.Rows = lngRP
    
    RaiseEvent FetchCompleteSANKU(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
