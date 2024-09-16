VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "馬主 データクラス"
'
'   馬主 データクラス
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
Private mCC As clsCodeConverter         '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter       '' JV-Data のコード変換メソッド群
' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 3) As String    '' テキストデータ
'モジュール内で使用する為のデータ格納変数
Private mBuf_BN As JV_BN_BANUSI
Private mKey As clsKeyBN                '' キー

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
Public Function Fetch(ByRef key As clsKeyBN) As Boolean
On Error GoTo ErrorHandler
    Dim cn      As ADODB.Connection
    Dim rs      As ADODB.Recordset
    Dim strSQL  As String
    
    Set mKey = key
    
    Set cn = gApp.GetCN_BANUSI
    
    'BANUSI (馬主マスタ)
    strSQL = "SELECT * FROM BANUSI " & mKey.SQLWHEREString
    Set rs = cn.Execute(strSQL)

    ' データがなければ、終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        Exit Function
    End If
    
    Call SetDataFromRS_BN(rs, mBuf_BN)
    
    rs.Close
    Set rs = Nothing
    
    Call MakeDataBANUSI
    Call MakeDataSEISEKI
    
    Set mAsyncCN_UMA = gApp.GetCN_UMA
    strSQL = "SELECT * FROM UMA WHERE [BanusiCode]='" & mKey.str & "'"
    strSQL = strSQL & " ORDER BY [BirthDate]"
    
    mAsyncCN_UMA.Execute strSQL, , adAsyncExecute

    gApp.Log "BN Fetch Complete"
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
'   機能: 競走馬マスタレコードセットの取得完了イベントハンドラ
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
            Call MakeDataUMA
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
'   機能: 馬主画面用データを作る
'
'   備考: なし
'
Private Sub MakeDataBANUSI()
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    
    Set rs = mRS_UMA
    
    With mBuf_BN
        ' 作成年月日
        With .head.MakeDate
            mstrLabels(0) = mstrLabels(0) & mSC.YMD1(.Year & .Month & .Day) & "作成データ"
        End With

        ' 太字レース名等グレーバー用文字列生成
        mstrLabels(1) = mstrLabels(1) & Trim$(IIf(.BanusiName_Co = "", .BanusiName, .BanusiName_Co))
        mstrLabels(2) = mstrLabels(2) & mSC.CTrim(.BanusiNameKana)

        ' ヘッダ領域
        mstrLabels(3) = mstrLabels(3) & Trim$(.BanusiNameEng) & vbCrLf
        mstrLabels(3) = mstrLabels(3) & IIf(Trim$(.Fukusyoku) = "", "", "服色：" & Trim$(.Fukusyoku))
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
Private Sub MakeDataSEISEKI()
On Error GoTo ErrorHandler
    Dim gd      As clsGridData
    Dim lngCP   As Long
    Dim lngRP   As Long
    Dim i       As Long
    
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
        .SetItemMatrix lngRP + 1, lngCP, mBuf_BN.HonRuikei(0).SetYear & "年"
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BN.HonRuikei(0).HonSyokinTotal)
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BN.HonRuikei(0).Fukasyokin)
        For i = 0 To 9
            .SetItemMatrix lngRP + 1, lngCP, mSC.Chakukaisu2(mBuf_BN.HonRuikei(0).Chakukaisu, i)
        Next i
        lngCP = 0
        lngRP = 2
        .SetItemMatrix lngRP + 1, lngCP, "累計"
        .SetItemMatrix lngRP + 1, lngCP, ""
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BN.HonRuikei(1).HonSyokinTotal)
        .SetItemMatrix lngRP + 1, lngCP, mSC.Money(mBuf_BN.HonRuikei(1).Fukasyokin)
        For i = 0 To 9
            .SetItemMatrix lngRP + 1, lngCP, mSC.Chakukaisu2(mBuf_BN.HonRuikei(1).Chakukaisu, i)
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
Private Sub MakeDataUMA()
On Error GoTo ErrorHandler
    Dim gd      As clsGridData
    Dim rs      As ADODB.Recordset
    Dim CH      As ADODB.Recordset
    Dim BR      As ADODB.Recordset
    Dim lngCP   As Long
    Dim lngRP   As Long
    Set rs = mRS_UMA
    
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData
    
    ' 調教師リンク確認用
    Set CH = New ADODB.Recordset
    Call OpenTableDirect(CH, gApp.GetCN_CHOKYO, "CHOKYO")
    
    ' 馬主リンク確認用
    Set BR = New ADODB.Recordset
    Call OpenTableDirect(BR, gApp.GetCN_SEISAN, "SEISAN")
    
    With gd
        ' グリッドデータにロウ数、カラム数を設定
        .Rows = 50
        .Cols = 8

        'カラムヘッダ
        lngCP = 0
        lngRP = -1
        .SetItemMatrix lngRP + 1, lngCP, "馬記号"
        .SetItemMatrix lngRP + 1, lngCP, "馬名"
        .SetItemMatrix lngRP + 1, lngCP, "性齢"
        .SetItemMatrix lngRP + 1, lngCP, "生年月日"
        .SetItemMatrix lngRP + 1, lngCP, "調教師"
        .SetItemMatrix lngRP + 1, lngCP, "生産者"
        .SetItemMatrix lngRP + 1, lngCP, "産地"
        .SetItemMatrix lngRP + 1, lngCP, "抹消"
    End With

    lngRP = 1
    Do While Not rs.EOF
        ' バックグラウンド
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        
        Call SafeSeek(CH, Array("ChokyosiCode"), Array(rs("ChokyosiCode").value))
        Call SafeSeek(BR, Array("BreederCode"), Array(rs("BreederCode").value))
        
        lngCP = 0
        With gd
            .SetItemMatrix lngRP, lngCP, mCC.UMKG1(rs("UmaKigoCD")), mCC.UMKG3(rs("UmaKigoCD"))
            .SetItemMatrix lngRP, lngCP, Trim$(rs("Bamei")), , , "UM", rs("KettoNum")
            .SetItemMatrix lngRP, lngCP, mCC.SEIB4(rs("SexCD")) & mSC.Age(mBuf_BN.head.MakeDate.Year, rs("BirthDate"))
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("BirthDate"))
            .SetItemMatrix lngRP, lngCP, Trim$(rs("ChokyosiRyakusyo")), , , "CH", IfExist(CH, "ChokyosiCode")
            .SetItemMatrix lngRP, lngCP, Trim$(rs("BreederName")), , , "BR", IfExist(BR, "BreederCode")
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
    If lngRP > 1 Then
        RaiseEvent FetchCompleteUMA(gd)
    Else
        RaiseEvent NoDataUMA
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
