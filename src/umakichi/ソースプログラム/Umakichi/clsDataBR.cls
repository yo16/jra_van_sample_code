VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "生産者 データクラス"
'
'   生産者 データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event NoDataSEISEKI()
Public Event NoDataUMA()

Public Event FetchCompleteSEISEKI(gd As clsGridData)
Public Event FetchCompleteUMA(gd As clsGridData)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter          '' JV-Data のコード変換メソッド群
' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 3) As String     '' テキストデータ
'モジュール内で使用する為のデータ格納変数
Private mKey As clsKeyBR                '' キー
Private mBuf_BR As JV_BR_BREEDER

'コネクション
Private WithEvents mAsyncCN_UMA As ADODB.Connection
Attribute mAsyncCN_UMA.VB_VarHelpID = -1
'レコードセット
Private mRS_UMA As ADODB.Recordset

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

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データを取得する
'
'   備考: なし
'
Public Function Fetch(ByRef key As clsKeyBR) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set mKey = key
    
    Set cn = gApp.GetCN_SEISAN
    
    'SEISAN (生産者マスタ)
    strSQL = "SELECT * FROM SEISAN " & mKey.SQLWHEREString

    Set rs = cn.Execute(strSQL)
    
    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        Exit Function
    End If

    Call SetDataFromRS_BR(rs, mBuf_BR)
    
    Call Set_SEISAN
    Call Set_SEISEKI
    
    rs.Close
   
    Set rs = Nothing
    
    Set mAsyncCN_UMA = gApp.GetCN_UMA
    strSQL = "SELECT * FROM UMA WHERE [BreederCode]='" & mKey.str & "'"
    strSQL = strSQL & " ORDER BY [BirthDate]"
    
    mAsyncCN_UMA.Execute strSQL, , adAsyncExecute
    
    gApp.Log "BR Fetch Complete"
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
    
    Call freers(mRS_UMA)

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
'   機能: 生産者マスタレコードセットの取得完了イベントハンドラ
'
'   備考: Fetchから抜けて裏で行われる処理
'
Private Sub mAsyncCN_UMA_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler

    If pError Is Nothing Then
        Set mRS_UMA = pRecordset
        
        If mRS_UMA.EOF Then
            RaiseEvent NoDataUMA
        Else
            Call Set_Uma
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataUMA
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 生産者画面用データを作る
'
'   備考: なし
'
Private Sub Set_SEISAN()
On Error GoTo ErrorHandler
    With mBuf_BR
        
        ' 作成年月日
        With .head.MakeDate
            mstrLabels(0) = mstrLabels(0) & mSC.YMD1(.Year & .Month & .Day) & "作成データ"
        End With
        
        ' 太字レース名等グレーバー用文字列生成
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & Trim$(IIf(.BreederName_Co = "", .BreederName, .BreederName_Co))
        mstrLabels(2) = ""
        mstrLabels(2) = mstrLabels(2) & mSC.CTrim(.BreederNameKana)
        
        mstrLabels(3) = ""
        mstrLabels(3) = mstrLabels(3) & Trim$(.BreederNameEng) & vbCrLf
        mstrLabels(3) = mstrLabels(3) & IIf(Trim$(.Address) = "", "", "所在地：" & Trim$(.Address))
    
    End With
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 成績グリッドを作る
'
'   備考: なし
'
Private Sub Set_SEISEKI()
On Error GoTo ErrorHandler
    Dim gd      As clsGridData
    Dim lngCP   As Long
    Dim lngRP   As Long
    Dim i       As Long
    
    ' 成績　タブ
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData

    With gd
        .Rows = 4
        .Cols = 14
        lngCP = 0
        lngRP = -1
        .SetItemMatrix lngRP + 1, lngCP, " "
        .SetItemMatrix lngRP + 1, lngCP, " "
        .SetItemMatrix lngRP + 1, lngCP, " "
        .SetItemMatrix lngRP + 1, lngCP, " "
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        .SetItemMatrix lngRP + 1, lngCP, "総合"
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP + 1, lngCP, " "
        .SetItemMatrix lngRP + 1, lngCP, "設定年"
        .SetItemMatrix lngRP + 1, lngCP, "本賞金"
        .SetItemMatrix lngRP + 1, lngCP, "付加賞金"
        .SetItemMatrix lngRP + 1, lngCP, "1着"
        .SetItemMatrix lngRP + 1, lngCP, "2着"
        .SetItemMatrix lngRP + 1, lngCP, "3着"
        .SetItemMatrix lngRP + 1, lngCP, "4着"
        .SetItemMatrix lngRP + 1, lngCP, "5着"
        .SetItemMatrix lngRP + 1, lngCP, "着外"
        .SetItemMatrix lngRP + 1, lngCP, "合計"
        .SetItemMatrix lngRP + 1, lngCP, "勝率"
        .SetItemMatrix lngRP + 1, lngCP, "連対率"
        .SetItemMatrix lngRP + 1, lngCP, "3着内率"
        lngCP = 0
        lngRP = 1
        .SetItemMatrix lngRP + 1, lngCP, "本年"
        .SetItemMatrix lngRP + 1, lngCP, mBuf_BR.HonRuikei(0).SetYear & "年"
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BR.HonRuikei(0).HonSyokinTotal)
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BR.HonRuikei(0).Fukasyokin)
        For i = 0 To 9
            .SetItemMatrix lngRP + 1, lngCP, mSC.Chakukaisu2(mBuf_BR.HonRuikei(0).Chakukaisu, i)
        Next i
        lngCP = 0
        lngRP = 2
        .SetItemMatrix lngRP + 1, lngCP, "累計"
        .SetItemMatrix lngRP + 1, lngCP, ""
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BR.HonRuikei(1).HonSyokinTotal)
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BR.HonRuikei(1).Fukasyokin)
        For i = 0 To 9
            .SetItemMatrix lngRP + 1, lngCP, mSC.Chakukaisu2(mBuf_BR.HonRuikei(1).Chakukaisu, i)
        Next i
    End With
    
    RaiseEvent FetchCompleteSEISEKI(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
    

'
'   機能: 競走馬グリッドを作る
'
'   備考: なし
'
Private Sub Set_Uma()
On Error GoTo ErrorHandler
    Dim gd      As clsGridData
    Dim rs      As ADODB.Recordset
    Dim BN      As ADODB.Recordset
    Dim CH      As ADODB.Recordset
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Set rs = mRS_UMA
    
    Dim buf_UM As JV_UM_UMA
    Dim i As Long
    Dim j As Long
    
    ' 競走馬　タブ
    
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData
    
    ' 馬主リンク確認用
    Set BN = New ADODB.Recordset
    Call OpenTableDirect(BN, gApp.GetCN_BANUSI, "BANUSI")
    
    ' 調教師リンク確認用
    Set CH = New ADODB.Recordset
    Call OpenTableDirect(CH, gApp.GetCN_CHOKYO, "CHOKYO")
    
    With gd
        
        ' グリッドデータにロウ数、カラム数を設定
        .Rows = 10
        .Cols = 8
        
        'カラムヘッダ
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, "馬記号"
        .SetItemMatrix lngRP, lngCP, "馬名"
        .SetItemMatrix lngRP, lngCP, "性齢"
        .SetItemMatrix lngRP, lngCP, "生年月日"
        .SetItemMatrix lngRP, lngCP, "馬主"
        .SetItemMatrix lngRP, lngCP, "調教師"
        .SetItemMatrix lngRP, lngCP, "産地"
        .SetItemMatrix lngRP, lngCP, "抹消"
    End With
    rs.MoveFirst
    
    lngRP = 1
    Do While Not rs.EOF
        ' バックグラウンド
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        
        Call SafeSeek(BN, Array("BanusiCode"), Array(rs("BanusiCode").value))
        Call SafeSeek(CH, Array("ChokyosiCode"), Array(rs("ChokyosiCode").value))
        
        lngCP = 0
        With gd
            .SetItemMatrix lngRP, lngCP, mCC.UMKG1(rs("UmaKigoCD")), mCC.UMKG3(rs("UmaKigoCD"))
            .SetItemMatrix lngRP, lngCP, Trim$(rs("Bamei")), , , "UM", rs("KettoNum")
            .SetItemMatrix lngRP, lngCP, mCC.SEIB4(rs("SexCD")) & mSC.Age(mBuf_BR.head.MakeDate.Year, rs("BirthDate"))
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("BirthDate"))
            .SetItemMatrix lngRP, lngCP, Trim$(rs("BanusiName")), , , "BN", IfExist(BN, "BanusiCode")
            .SetItemMatrix lngRP, lngCP, Trim$(rs("ChokyosiRyakusyo")), , , "CH", IfExist(CH, "ChokyosiCode")
            .SetItemMatrix lngRP, lngCP, Trim$(rs("SanchiName"))
            .SetItemMatrix lngRP, lngCP, Trim$(IIf(val(rs("DelKubun")) = 1, "抹消", ""))
        End With
        rs.MoveNext
        lngRP = lngRP + 1
        If lngRP >= gd.Rows Then
            gd.Rows = gd.Rows + 10
        End If
    Loop
    rs.Close
  
    Set rs = Nothing
    
    gd.Rows = lngRP
    
    RaiseEvent FetchCompleteUMA(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

