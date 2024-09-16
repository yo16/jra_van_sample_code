VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "産駒 データクラス"
'
'   産駒 データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event NoData()
Public Event NoDataKetto()
Public Event FetchCompleteKetto(gd As clsGridData)
Public Event FetchCompleteLinklabels()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter          '' JV-Data のコード変換メソッド群
' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 1) As String     '' テキストデータ
Private mstrLinkLabels(0 To 1) As clsGridItem     '' リンクデータ
'モジュール内で使用する為のデータ格納変数
Private mKey As clsKeySK                '' キー
Private mBuf_SK As JV_SK_SANKU

'コネクション
Private WithEvents mAsyncCN_SEISAN As ADODB.Connection
Attribute mAsyncCN_SEISAN.VB_VarHelpID = -1
Private mCN_HANSYOKU As ADODB.Connection
Private mCN_SANKU As ADODB.Connection
Private mCN_UMA As ADODB.Connection
'レコードセット
Private mRS_SEISAN As ADODB.Recordset
Private mRS_HANSYOKU As ADODB.Recordset
Private mRS_SANKU As ADODB.Recordset
Private mRS_UMA As ADODB.Recordset

Private mblnCancelFetching As Boolean


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ラベル情報を取得するプロパティ
'
'   備考: なし
'
Public Property Get Labels(Index As Integer) As String
    Labels = mstrLabels(Index)
End Property

'
'   機能: リンクラベル情報を取得するプロパティ
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
'   機能: 取得を開始する
'
'   備考: 初期表示用データの取得
'
Public Function Fetch(ByRef key As clsKeySK) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set mKey = key
    
    Set cn = gApp.GetCN_SANKU
    
    'HANSYOKU (産駒マスタ)
    strSQL = "SELECT * FROM SANKU " & key.SQLWHEREString
    
    Set rs = cn.Execute(strSQL)
    
    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        RaiseEvent NoData
        Exit Function
    End If

    Call SetDataFromRS_SK(rs, mBuf_SK)
    
    Set mRS_SANKU = rs
    
    Call Set_SANKU
    
    Set mAsyncCN_SEISAN = gApp.GetCN_SEISAN
    
    strSQL = "SELECT * FROM SEISAN WHERE [BreederCode]='" & mBuf_SK.BreederCode & "'"
    mAsyncCN_SEISAN.Execute strSQL, , adAsyncExecute
    
    gApp.Log "SK Fetch Complete"
    
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
'   機能: 生産者レコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_SEISAN_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    
    If pError Is Nothing Then
        If Not (pRecordset.EOF Or pRecordset.BOF) Then
            Call Set_LinkLabels(pRecordset)
        End If
        Call Set_Ketto
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
'   機能: 産駒画面用データを作成
'
'   備考: なし
'
Private Sub Set_SANKU()
On Error GoTo ErrorHandler
    Dim i As Long
    Dim strSQL  As String
    Dim BR As ADODB.Recordset
    
    ' リンクデータオブジェクトを生成する
    For i = 0 To 1
        Set mstrLinkLabels(i) = New clsGridItem
    Next i
    
    With mBuf_SK
        
        ' 作成年月日
        With .head.MakeDate
            mstrLabels(0) = mstrLabels(0) & mSC.YMD1(.Year & .Month & .Day) & "作成データ"
        End With
                
        'header
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & IIf(Trim$(.SanchiName) = "", "", "産地　：" & Trim$(.SanchiName)) & vbCrLf
        mstrLabels(1) = mstrLabels(1) & vbCrLf
        mstrLabels(1) = mstrLabels(1) & IIf(val(.ImportYear) = 0, "", "輸入年：" & .ImportYear) & vbCrLf
        With .BirthDate
            mstrLabels(1) = mstrLabels(1) & IIf(val(.Year) = 0, "", mSC.YMD1(.Year & .Month & .Day) & "生 ") & vbCrLf
        End With
        mstrLabels(1) = mstrLabels(1) & mCC.HINS2(.HinsyuCD) & " "  '品種
        mstrLabels(1) = mstrLabels(1) & mCC.SEIB4(.SexCD) & " "   '性別
        mstrLabels(1) = mstrLabels(1) & mCC.KEIR1(.KeiroCD) & " "  '毛色
    
        '繁殖馬テーブルを開く
        Set mCN_HANSYOKU = gApp.GetCN_HANSYOKU
        Set mRS_HANSYOKU = New ADODB.Recordset
        Call OpenTableDirect(mRS_HANSYOKU, mCN_HANSYOKU, "HANSYOKU")
        
        '競走馬テーブルを開く
        Set mCN_UMA = gApp.GetCN_UMA
        Set mRS_UMA = New ADODB.Recordset
        Call OpenTableDirect(mRS_UMA, mCN_UMA, "UMA")
        
        ' 太字レース名等グレーバー用文字列生成
        ' リンク用文字列生成
        Call set_LinkLabel0(mRS_UMA, mRS_HANSYOKU, .KettoNum)
    End With
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: リンクラベルを作成
'
'   備考: なし
'
Private Sub Set_LinkLabels(BR As ADODB.Recordset)
On Error GoTo ErrorHandler

    mstrLinkLabels(1).Text = Trim$(IIf(Trim$(BR("BreederName_Co")) = "", BR("BreederName"), BR("BreederName_Co")))
    mstrLinkLabels(1).Link = "BR"
    mstrLinkLabels(1).key = mBuf_SK.BreederCode
    mstrLinkLabels(1).ToolTip = ""

    RaiseEvent FetchCompleteLinklabels
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 血統画面用データを作成
'
'   備考: なし
'
Private Sub Set_Ketto()
On Error GoTo ErrorHandler

    Dim gd      As clsGridData
    Dim rs      As ADODB.Recordset
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim i As Long
    Dim j As Long
    
    Set rs = mRS_SANKU
    
    '　血統タブ
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData

    With gd
        .Rows = 9
        .Cols = 3
        ' カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "親(父/母)"
        .SetItemMatrix 0, lngCP, "二代"
        .SetItemMatrix 0, lngCP, "三代"
    End With
    
    Call KettoSub(rs, gd, 7, "M", "MM", "MMM")
    Call KettoSub(rs, gd, 6, "M", "MM", "MMF")
    Call KettoSub(rs, gd, 5, "M", "MF", "MFM")
    Call KettoSub(rs, gd, 4, "M", "MF", "MFF")
    Call KettoSub(rs, gd, 3, "F", "FM", "FMM")
    Call KettoSub(rs, gd, 2, "F", "FM", "FMF")
    Call KettoSub(rs, gd, 1, "F", "FF", "FFM")
    Call KettoSub(rs, gd, 0, "F", "FF", "FFF")

    rs.Close
    
    Set rs = Nothing
    
    RaiseEvent FetchCompleteKetto(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: 血統情報作成用サブルーチン
'
'   備考: なし
'
Private Sub KettoSub(SK As ADODB.Recordset, gd As clsGridData, RowP As Long, Field1 As String, Field2 As String, Field3 As String)
    Dim HN As ADODB.Recordset
    Dim lngCP As Long
    Dim strKey As String
    Dim strBamei As String
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Set HN = mRS_HANSYOKU
    
    lngCP = 0
    With gd
        strKey = SK(Field1 & "Num")
        strBamei = seek_HN(HN, strKey)
        strKey = IIf(strBamei = "", "", SK(Field1 & "Num"))
        strBamei = IIf(strBamei = "", Space(Int(RowP / 4) + 1), Trim$(strBamei))
        .SetItemMatrix RowP + rowOffset, lngCP, strBamei, "" _
            , , "HN", strKey, IIf(Right$(Field1, 1) = "M", ColorMother, ColorFather)
        
        strKey = SK(Field2 & "Num")
        strBamei = seek_HN(HN, strKey)
        strKey = IIf(strBamei = "", "", SK(Field2 & "Num"))
        strBamei = IIf(strBamei = "", Space(Int(RowP / 2) + 1), Trim$(strBamei))
        .SetItemMatrix RowP + rowOffset, lngCP, strBamei, "" _
            , , "HN", strKey, IIf(Right$(Field2, 1) = "M", ColorMother, ColorFather)
        
        strKey = SK(Field3 & "Num")
        strBamei = Trim$(seek_HN(HN, strKey))
        .SetItemMatrix RowP + rowOffset, lngCP, strBamei, "" _
            , , "HN", IIf(strBamei = "", "", strKey), IIf(Right$(Field3, 1) = "M", ColorMother, ColorFather)
    End With
End Sub


'
'   機能: 繁殖馬に該当レコードが存在するか返す
'
'   備考: なし
'
Private Function seek_HN(rs As ADODB.Recordset, key As String) As String
On Error GoTo ErrorHandler
    Call SafeSeek(rs, Array("HansyokuNum"), Array(key))
    
    ' データが無かったら""を返して終了
    If rs.EOF Or rs.BOF Then
        seek_HN = ""
        Exit Function
    End If
    
    seek_HN = Trim$(rs("Bamei"))
        
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: リンクラベルを作成
'
'   備考: なし
'
Private Sub set_LinkLabel0(UM As ADODB.Recordset, HN As ADODB.Recordset, key As String)
On Error GoTo ErrorHandler
    
    mstrLinkLabels(0).Link = "UM"
    mstrLinkLabels(0).ToolTip = ""
    
    Call SafeSeek(UM, Array("KettoNum"), Array(key))
    
    ' データが無かったらstrKeyを返して終了
    If UM.EOF Or UM.BOF Then
        Call SafeSeek(HN, Array("KettoNum"), Array(key))
        If HN.EOF Or HN.BOF Then
            mstrLinkLabels(0).Text = key
        Else
            mstrLinkLabels(0).Text = Trim$(HN("Bamei"))
        End If
    Else
        mstrLinkLabels(0).Text = Trim$(UM("Bamei"))
        mstrLinkLabels(0).key = key
    End If
    
    UM.Close
    Set UM = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
