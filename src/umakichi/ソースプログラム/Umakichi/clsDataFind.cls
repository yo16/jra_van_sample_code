VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "検索画面データ取得オブジェクト"
'
'   検索画面データ取得オブジェクト
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mGridData(0 To 5) As clsGridData
Private mCC As clsCodeConverter
Private mSC As clsStringConverter
Private mRX As RegExp

Private WithEvents mAsyncCN_UMA         As ADODB.Connection
Attribute mAsyncCN_UMA.VB_VarHelpID = -1
Private mRS_UMA                         As ADODB.Recordset
Private WithEvents mAsyncCN_KISHU       As ADODB.Connection
Attribute mAsyncCN_KISHU.VB_VarHelpID = -1
Private mRS_KISHU                       As ADODB.Recordset
Private WithEvents mAsyncCN_CHOKYO      As ADODB.Connection
Attribute mAsyncCN_CHOKYO.VB_VarHelpID = -1
Private mRS_CHOKYO                      As ADODB.Recordset
Private WithEvents mAsyncCN_BANUSI      As ADODB.Connection
Attribute mAsyncCN_BANUSI.VB_VarHelpID = -1
Private mRS_BANUSI                      As ADODB.Recordset
Private WithEvents mAsyncCN_SEISAN      As ADODB.Connection
Attribute mAsyncCN_SEISAN.VB_VarHelpID = -1
Private mRS_SEISAN                      As ADODB.Recordset
Private WithEvents mAsyncCN_HANSYOKU    As ADODB.Connection
Attribute mAsyncCN_HANSYOKU.VB_VarHelpID = -1
Private mRS_HANSYOKU                    As ADODB.Recordset

Private mblnCancelFetching As Boolean
Private mblnNowFetching As Boolean
Private mblnNowCanceling As Boolean

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' イベント
Public Event FetchCompleteUMA(gd As clsGridData)
Public Event FetchCompleteKISHU(gd As clsGridData)
Public Event FetchCompleteCHOKYO(gd As clsGridData)
Public Event FetchCompleteBANUSI(gd As clsGridData)
Public Event FetchCompleteSEISAN(gd As clsGridData)
Public Event FetchCompleteHANSYOKU(gd As clsGridData)
Public Event NoDataUMA()
Public Event NoDataKISHU()
Public Event NoDataCHOKYO()
Public Event NoDataBANUSI()
Public Event NoDataSEISAN()
Public Event NoDataHANSYOKU()


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'
'   機能: グリッドデータを返す
'
'   備考: なし
'
Public Property Get GridData(ByVal Index As Long) As clsGridData
    Set GridData = mGridData(Index)
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
'   機能: データを取得する
'
'   備考: なし
'
Public Sub Fetch(Index As Long, key As String, Mode As Long)
On Error GoTo ErrorHandler
    Dim strSQL  As String
    Dim rx As RegExp
    
    gApp.Log "Fetch " & key
    
    ' スペースをすべて削除する
    Set rx = New RegExp
    rx.Global = True
    rx.Pattern = "[\s　]"
    key = rx.Replace(key, "")
    
    mblnNowFetching = True
    mblnCancelFetching = False
    
    mRX.Pattern = ByModeRegExp(key, Mode)

    Select Case Index
    Case 0                                  ' 馬名 Search

        Set mAsyncCN_UMA = gApp.GetCN_UMA

        strSQL = "SELECT * FROM UMA WHERE"
        strSQL = strSQL & " [Bamei] LIKE " & ByModeUMA(key, Mode)
        strSQL = strSQL & " OR [BameiEng] LIKE " & ByModeUMA(key, Mode)
        strSQL = strSQL & " ORDER BY [Bamei]"

        mAsyncCN_UMA.Execute strSQL, , adAsyncExecute
    Case 1                                  ' 騎手名 Search

        Set mAsyncCN_KISHU = gApp.GetCN_KISHU

        strSQL = "SELECT * FROM KISHU WHERE"
        strSQL = strSQL & " [KisyuName] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [KisyuNameKana] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [KisyuRyakusyo] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [KisyuNameEng] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " ORDER BY [KisyuNameKana]"

        mAsyncCN_KISHU.Execute strSQL, , adAsyncExecute
    Case 2                                  ' 調教師名 Search

        Set mAsyncCN_CHOKYO = gApp.GetCN_CHOKYO

        strSQL = "SELECT * FROM CHOKYO WHERE"
        strSQL = strSQL & " [ChokyosiName] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [ChokyosiNameKana] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [ChokyosiRyakusyo] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [ChokyosiNameEng] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " ORDER BY [ChokyosiNameKana]"
    
        mAsyncCN_CHOKYO.Execute strSQL, , adAsyncExecute
    Case 3                                  ' 馬主名 Search

        Set mAsyncCN_BANUSI = gApp.GetCN_BANUSI

        strSQL = "SELECT * FROM BANUSI WHERE"
        strSQL = strSQL & " [BanusiName_Co] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BanusiName] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BanusiNameKana] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BanusiNameEng] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " ORDER BY [BanusiName]"
        
        mAsyncCN_BANUSI.Execute strSQL, , adAsyncExecute
    Case 4                                  ' 生産者名 Search

        Set mAsyncCN_SEISAN = gApp.GetCN_SEISAN

        strSQL = "SELECT * FROM SEISAN WHERE"
        strSQL = strSQL & " [BreederName_Co] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BreederName] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BreederNameKana] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BreederNameEng] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " ORDER BY [BreederName]"

        mAsyncCN_SEISAN.Execute strSQL, , adAsyncExecute
    Case 5                                  ' 繁殖馬名 Search

        Set mAsyncCN_HANSYOKU = gApp.GetCN_HANSYOKU

        strSQL = "SELECT * FROM HANSYOKU WHERE"
        strSQL = strSQL & " [Bamei] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BameiKana] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " OR [BameiEng] LIKE " & ByMode(key, Mode)
        strSQL = strSQL & " ORDER BY [BameiKana]"

        mAsyncCN_HANSYOKU.Execute strSQL, , adAsyncExecute
    End Select
    
    gApp.Log "Find SQL: " & strSQL
    gApp.Log "Find RegExp: " & mRX.Pattern
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    mblnNowFetching = False
    mblnNowCanceling = False
End Sub


'
'   機能: キャンセル(データ取得)
'
'   備考: なし
'
Public Sub CancelFetching()
On Error GoTo ErrorHandler
    mblnCancelFetching = True
    
    freers mRS_UMA
    freers mRS_KISHU
    freers mRS_CHOKYO
    freers mRS_BANUSI
    freers mRS_SEISAN
    freers mRS_HANSYOKU
    
    freecn mAsyncCN_UMA
    freecn mAsyncCN_KISHU
    freecn mAsyncCN_CHOKYO
    freecn mAsyncCN_BANUSI
    freecn mAsyncCN_SEISAN
    freecn mAsyncCN_HANSYOKU
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: キャンセル(検索)
'
'   備考: なし
'
Public Sub CancelFind()
On Error GoTo ErrorHandler
        
    If mblnNowCanceling = False Then
        mblnNowCanceling = True
        ' コネクションが取得実行中ならキャンセルする
        
        gApp.Log "Cancel Flag On"
        mblnCancelFetching = True
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 競走馬グリッドを作る
'
'   備考: なし
'
Private Sub MakeData_UMA()
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim lngCP   As Long
    Dim lngRP   As Long
    Dim lngYear As Long
    
    Dim UM_SeiNen As String
    Dim UM_NenRei As String
    Dim strBameiEng As String
    
    gApp.Log "MakeData_UMA"
    
    gApp.Log "Check Cancel: MakeData_UMA " & mblnCancelFetching
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        Exit Sub
    End If
    
    Set mGridData(0) = New clsGridData
    
    Set rs = mRS_UMA
    
    With mGridData(0)
        .Cols = 12
        .Rows = 500
        .SetItemMatrix 0, lngCP, "馬名"
        .SetItemMatrix 0, lngCP, "馬名欧字"
        .SetItemMatrix 0, lngCP, "馬記号"
        .SetItemMatrix 0, lngCP, "馬齢"
        .SetItemMatrix 0, lngCP, "抹消区分"
        .SetItemMatrix 0, lngCP, "性別"
        .SetItemMatrix 0, lngCP, "品種"
        .SetItemMatrix 0, lngCP, "毛色"
        .SetItemMatrix 0, lngCP, "東西所属"
        .SetItemMatrix 0, lngCP, "調教師名略称"
        .SetItemMatrix 0, lngCP, "生産者名"
        .SetItemMatrix 0, lngCP, "馬主名"
        
        lngYear = Year(Now)
        
        lngRP = 1
        Do While Not rs.EOF
            ' バックグラウンド
            If lngRP Mod 5 = 0 Then
                DoEvents
            End If
            If mblnCancelFetching Then
                mblnNowFetching = False
                mblnCancelFetching = False
                mblnNowCanceling = False
                Exit Sub
            End If
            
            strBameiEng = StrConv(LeftB(StrConv(rs("BameiEng"), vbFromUnicode), 60), vbUnicode) '左から60文字だけを検索・表示対象にする
            
            If mRX.Test(Trim$(rs("Bamei"))) _
            Or mRX.Test(Trim$(strBameiEng)) Then
                
                lngCP = 0
                
                UM_SeiNen = Mid$(rs("BirthDate"), 1, 4)
                UM_NenRei = lngYear - UM_SeiNen
                
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("Bamei"))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(strBameiEng)), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.UMKG1(Trim$(ContractSpace(rs("UmaKigoCD")))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, UM_NenRei, , ">-", "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, mSC.UM_DelKubun(Trim$(ContractSpace(rs("DelKubun")))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.SEIB1(Trim$(ContractSpace(rs("SexCD")))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.HINS1(Trim$(ContractSpace(rs("HinsyuCD")))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.KEIR1(Trim$(ContractSpace(rs("KeiroCD")))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.TZSZ1(Trim$(ContractSpace(rs("TozaiCD")))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("ChokyosiRyakusyo"))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BreederName"))), , , "UM", rs("KettoNum"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BanusiName"))), , , "UM", rs("KettoNum"), , 1
                
                lngRP = lngRP + 1
                If .Rows <= lngRP Then
                    .Rows = Int(.Rows * 1.5)
                    gApp.Log "Rows Up: " & .Rows
                End If
            End If
            rs.MoveNext
        Loop
        .Rows = lngRP
    End With
    If lngRP > 1 Then
        RaiseEvent FetchCompleteUMA(mGridData(0))
    Else
        RaiseEvent NoDataUMA
    End If
    
    gApp.Log "RegExpHit: " & lngRP
    
    rs.Close
    Set rs = Nothing
    mblnNowFetching = False
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 騎手画面用データを作る
'
'   備考: なし
'
Private Sub MakeData_KISHU()
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim lngCP As Long
    Dim lngRP As Long
        
    Set mGridData(1) = New clsGridData
    
    Set rs = mRS_KISHU
    
    With mGridData(1)
        .Cols = 8
        .Rows = 50
        lngCP = 0
        .SetItemMatrix 0, lngCP, "騎手名"
        .SetItemMatrix 0, lngCP, "騎手名半角ｶﾅ"
        .SetItemMatrix 0, lngCP, "騎手名略称"
        .SetItemMatrix 0, lngCP, "騎手名欧字"
        .SetItemMatrix 0, lngCP, "性別"
        .SetItemMatrix 0, lngCP, "騎乗資格"
        .SetItemMatrix 0, lngCP, "騎手見習"
        .SetItemMatrix 0, lngCP, "騎手東西所属"
        
        lngRP = 1
        Do While Not rs.EOF
            ' バックグラウンド
            DoEvents
            If mblnCancelFetching Then
                mblnNowFetching = False
                mblnCancelFetching = False
                mblnNowCanceling = False
                Exit Sub
            End If
            
            If mRX.Test(Trim$(rs("KisyuName"))) _
            Or mRX.Test(Trim$(rs("KisyuNameKana"))) _
            Or mRX.Test(Trim$(rs("KisyuRyakusyo"))) _
            Or mRX.Test(Trim$(rs("KisyuNameEng"))) _
            Then
    
                lngCP = 0
                
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("KisyuName"))), , , "KS", rs("KisyuCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("KisyuNameKana"))), , , "KS", rs("KisyuCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("KisyuRyakusyo"))), , , "KS", rs("KisyuCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("KisyuNameEng"))), , , "KS", rs("KisyuCode"), , 1
                .SetItemMatrix lngRP, lngCP, mSC.KS_SexCD(Trim$(ContractSpace(rs("SexCD")))), , , "KS", rs("KisyuCode"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.KSSK1(Trim$(ContractSpace(rs("SikakuCD")))), , , "KS", rs("KisyuCode"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.KSMN1(Trim$(ContractSpace(rs("MinaraiCD")))), , , "KS", rs("KisyuCode"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.TZSZ1(Trim$(ContractSpace(rs("TozaiCD")))), , , "KS", rs("KisyuCode"), , 1
                
                lngRP = lngRP + 1
                If .Rows <= lngRP Then
                    .Rows = Int(.Rows * 1.5)
                End If
            End If
            rs.MoveNext
        Loop
        .Rows = lngRP
    End With
    If lngRP > 1 Then
        RaiseEvent FetchCompleteKISHU(mGridData(1))
    Else
        RaiseEvent NoDataKISHU
    End If

    rs.Close
    Set rs = Nothing
    mblnNowFetching = False

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 調教師画面用データを作る
'
'   備考: なし
'
Private Sub MakeData_CHOKYO()
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim lngCP As Long
    Dim lngRP As Long
    
    Set mGridData(2) = New clsGridData

    Set rs = mRS_CHOKYO
    
    With mGridData(2)
        .Cols = 6
        .Rows = 50
        lngCP = 0
        .SetItemMatrix 0, lngCP, "調教師名"
        .SetItemMatrix 0, lngCP, "調教師名半角ｶﾅ"
        .SetItemMatrix 0, lngCP, "調教師名略称"
        .SetItemMatrix 0, lngCP, "調教師名欧字"
        .SetItemMatrix 0, lngCP, "性別"
        .SetItemMatrix 0, lngCP, "調教師東西所属"
        
        lngRP = 1
        Do While Not rs.EOF
            ' バックグラウンド
            DoEvents
            If mblnCancelFetching Then
                mblnNowFetching = False
                mblnCancelFetching = False
                mblnNowCanceling = False
                Exit Sub
            End If
            
            If mRX.Test(Trim$(rs("ChokyosiName"))) _
            Or mRX.Test(Trim$(rs("ChokyosiNameKana"))) _
            Or mRX.Test(Trim$(rs("ChokyosiRyakusyo"))) _
            Or mRX.Test(Trim$(rs("ChokyosiNameEng"))) _
            Then
            
                lngCP = 0
                
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("ChokyosiName"))), , , "CH", rs("ChokyosiCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("ChokyosiNameKana"))), , , "CH", rs("ChokyosiCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("ChokyosiRyakusyo"))), , , "CH", rs("ChokyosiCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("ChokyosiNameEng"))), , , "CH", rs("ChokyosiCode"), , 1
                .SetItemMatrix lngRP, lngCP, mSC.CH_SexCD(Trim$(ContractSpace(rs("SexCD")))), , , "CH", rs("ChokyosiCode"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.TZSZ1(Trim$(ContractSpace(rs("TozaiCD")))), , , "CH", rs("ChokyosiCode"), , 1
                
                lngRP = lngRP + 1
                If .Rows <= lngRP Then
                    .Rows = Int(.Rows * 1.5)
                End If
            End If
            rs.MoveNext
        Loop
        .Rows = lngRP
    End With
    If lngRP > 1 Then
        RaiseEvent FetchCompleteCHOKYO(mGridData(2))
    Else
        RaiseEvent NoDataCHOKYO
    End If
    
    rs.Close
    Set rs = Nothing
    mblnNowFetching = False
    
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
Private Sub MakeData_BANUSI()
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim lngCP As Long
    Dim lngRP As Long
    
    Set mGridData(3) = New clsGridData

    Set rs = mRS_BANUSI
    
    With mGridData(3)
        .Cols = 5
        .Rows = 50
        lngCP = 0
        .SetItemMatrix 0, lngCP, "馬主名 法人格有"
        .SetItemMatrix 0, lngCP, "馬主名 法人格無"
        .SetItemMatrix 0, lngCP, "馬主名半角ｶﾅ"
        .SetItemMatrix 0, lngCP, "馬主名欧字"
        .SetItemMatrix 0, lngCP, "服色標示"
        
        lngRP = 1
        Do While Not rs.EOF
            ' バックグラウンド
            DoEvents
            If mblnCancelFetching Then
                mblnNowFetching = False
                mblnCancelFetching = False
                mblnNowCanceling = False
                Exit Sub
            End If
            
            If mRX.Test(Trim$(rs("BanusiName_Co"))) _
            Or mRX.Test(Trim$(rs("BanusiName"))) _
            Or mRX.Test(Trim$(rs("BanusiNameKana"))) _
            Or mRX.Test(Trim$(rs("BanusiNameEng"))) _
            Then
            
                lngCP = 0
                
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BanusiName_Co"))), , , "BN", rs("BanusiCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BanusiName"))), , , "BN", rs("BanusiCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BanusiNameKana"))), , , "BN", rs("BanusiCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BanusiNameEng"))), , , "BN", rs("BanusiCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("Fukusyoku"))), , , "BN", rs("BanusiCode"), , 1
                
                lngRP = lngRP + 1
                If .Rows <= lngRP Then
                    .Rows = Int(.Rows * 1.5)
                End If
            End If
            rs.MoveNext
        Loop
        .Rows = lngRP
    End With
    If lngRP > 1 Then
        RaiseEvent FetchCompleteBANUSI(mGridData(3))
    Else
        RaiseEvent NoDataBANUSI
    End If
    
    rs.Close
    Set rs = Nothing
    mblnNowFetching = False
    
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
Private Sub MakeData_SEISAN()
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim lngRP As Long
    Dim lngCP As Long
    
    Set mGridData(4) = New clsGridData

    Set rs = mRS_SEISAN

    With mGridData(4)
        .Cols = 5
        .Rows = 50
        lngCP = 0
        .SetItemMatrix 0, lngCP, "生産者名 法人格有"
        .SetItemMatrix 0, lngCP, "生産者名 法人格無"
        .SetItemMatrix 0, lngCP, "生産者名半角ｶﾅ"
        .SetItemMatrix 0, lngCP, "生産者名欧字"
        .SetItemMatrix 0, lngCP, "生産者住所自治省名"
        
        lngRP = 1
        Do While Not rs.EOF
            ' バックグラウンド
            DoEvents
            If mblnCancelFetching Then
                mblnNowFetching = False
                mblnCancelFetching = False
                mblnNowCanceling = False
                Exit Sub
            End If
            
            If mRX.Test(Trim$(rs("BreederName_Co"))) _
            Or mRX.Test(Trim$(rs("BreederName"))) _
            Or mRX.Test(Trim$(rs("BreederNameKana"))) _
            Or mRX.Test(Trim$(rs("BreederNameEng"))) _
            Then
            
                lngCP = 0
                
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BreederName_Co"))), , , "BR", rs("BreederCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BreederName"))), , , "BR", rs("BreederCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BreederNameKana"))), , , "BR", rs("BreederCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BreederNameEng"))), , , "BR", rs("BreederCode"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("Address"))), , , "BR", rs("BreederCode"), , 1
                
                lngRP = lngRP + 1
                If .Rows <= lngRP Then
                    .Rows = Int(.Rows * 1.5)
                End If
            End If
            rs.MoveNext
        Loop
        .Rows = lngRP
    End With
    If lngRP > 1 Then
        RaiseEvent FetchCompleteSEISAN(mGridData(4))
    Else
        RaiseEvent NoDataSEISAN
    End If
    
    rs.Close
    Set rs = Nothing
    mblnNowFetching = False
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 繁殖馬画面用データを作る
'
'   備考: なし
'
Private Sub MakeData_HANSYOKU()
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim lngCP As Long
    Dim lngRP As Long
    
    Set mGridData(5) = New clsGridData

    Set rs = mRS_HANSYOKU
    
    With mGridData(5)
        .Cols = 8
        .Rows = 50
        lngCP = 0
        .SetItemMatrix 0, lngCP, "馬名"
        .SetItemMatrix 0, lngCP, "馬名半角ｶﾅ"
        .SetItemMatrix 0, lngCP, "馬名欧字"
        .SetItemMatrix 0, lngCP, "性別"
        .SetItemMatrix 0, lngCP, "品種"
        .SetItemMatrix 0, lngCP, "毛色"
        .SetItemMatrix 0, lngCP, "輸入年"
        .SetItemMatrix 0, lngCP, "産地名"

        lngRP = 1
        Do While Not rs.EOF
            ' バックグラウンド
            DoEvents
            If mblnCancelFetching Then
                mblnNowFetching = False
                mblnCancelFetching = False
                mblnNowCanceling = False
                Exit Sub
            End If
            
            If mRX.Test(Trim$(rs("Bamei"))) _
            Or mRX.Test(Trim$(rs("BameiKana"))) _
            Or mRX.Test(Trim$(rs("BameiEng"))) _
            Then
                
                lngCP = 0
                
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("Bamei"))), , , "HN", rs("HansyokuNum"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BameiKana"))), , , "HN", rs("HansyokuNum"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("BameiEng"))), , , "HN", rs("HansyokuNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.SEIB1(Trim$(ContractSpace(rs("SexCD")))), , , "HN", rs("HansyokuNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.HINS1(Trim$(ContractSpace(rs("HinsyuCD")))), , , "HN", rs("HansyokuNum"), , 1
                .SetItemMatrix lngRP, lngCP, mCC.KEIR1(Trim$(ContractSpace(rs("KeiroCD")))), , , "HN", rs("HansyokuNum"), , 1
                .SetItemMatrix lngRP, lngCP, IIf(val(rs("ImportYear")) = 0, "", Trim$(ContractSpace(rs("ImportYear")))), , , "HN", rs("HansyokuNum"), , 1
                .SetItemMatrix lngRP, lngCP, Trim$(ContractSpace(rs("SanchiName"))), , , "HN", rs("HansyokuNum"), , 1
                
                lngRP = lngRP + 1
                If .Rows <= lngRP Then
                    .Rows = Int(.Rows * 1.5)
                End If
            End If
            rs.MoveNext
        Loop
        .Rows = lngRP
    End With
    If lngRP > 1 Then
        RaiseEvent FetchCompleteHANSYOKU(mGridData(5))
    Else
        RaiseEvent NoDataHANSYOKU
    End If
    
    rs.Close
    Set rs = Nothing
    mblnNowFetching = False

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 正規表現処理
'
'   備考: なし
'
Private Function ByModeRegExp(key As String, Mode As Long) As String
    Dim i As Long
    Dim j   As Long
    Dim strCase(0 To 7) As String
    Dim strTmp As String
    Dim regPattern As Dictionary
    
    Set regPattern = New Dictionary
    
    strCase(0) = StrConv(key, vbHiragana + vbWide)
    strCase(1) = StrConv(key, vbKatakana + vbWide)
    strCase(2) = StrConv(key, vbHiragana + vbNarrow)
    strCase(3) = StrConv(key, vbKatakana + vbNarrow)
    strCase(4) = StrConv(key, vbUpperCase + vbWide)
    strCase(5) = StrConv(key, vbLowerCase + vbWide)
    strCase(6) = StrConv(key, vbUpperCase + vbNarrow)
    strCase(7) = StrConv(key, vbLowerCase + vbNarrow)

    For i = 0 To 7
        strTmp = IIf(Mode = 0, "^", "") & CharGroup(Left$(strCase(i), 1))
        For j = 2 To Len(strCase(i))
            strTmp = strTmp & "\s*" & CharGroup(Mid$(strCase(i), j, 1))
        Next j
        strTmp = strTmp & IIf(Mode = 1, "\s*$", "")
        If Not regPattern.Exists(strTmp) Then
            regPattern.Add strTmp, strTmp
        End If
    Next i
    
    ByModeRegExp = Join(regPattern.Keys, "|")
End Function


'
'   機能: 正規表現処理
'
'   備考: なし
'
Private Function ByMode(key As String, Mode As Long) As String
    Dim i   As Long
    Dim tmp As String

    
    tmp = Left$(key, 1)
    For i = 2 To Len(key)
        tmp = tmp & "%" & Mid$(key, i, 1)
    Next i
    
    tmp = Replace(tmp, "'", "''")
    tmp = CharGroup(tmp)
        
    Select Case Mode
    Case 0
        ByMode = "'" & tmp & "%'"
    Case 1
        ByMode = "'%" & tmp & "%'" ' 後方一致は、後ろの不定数スペースをＳＱＬで指定できないので、正規表現側でフィルタします。
    Case 2
        ByMode = "'%" & tmp & "%'"
    End Select
End Function


'
'   機能: 正規表現処理(競走馬)
'
'   備考: なし
'
Private Function ByModeUMA(key As String, Mode As Long) As String
    Dim i   As Long
    Dim tmp As String

    
    tmp = key
    
    tmp = Replace(tmp, "'", "''")
        
    Select Case Mode
    Case 0
        ByModeUMA = "'" & tmp & "%'"
    Case 1
        ByModeUMA = "'%" & tmp & "%'" ' 後方一致は、後ろの不定数スペースをＳＱＬで指定できないので、正規表現側でフィルタします。
    Case 2
        ByModeUMA = "'%" & tmp & "%'"
    End Select
End Function


'
'   機能: 検索文字の代替
'
'   備考: なし
'
Private Function CharGroup(str As String) As String
    Dim i As Long
    Dim c As String
    Dim out As String
    Dim posHD As Long
    Dim posKD As Long
    Dim posKS As Long
    Dim posG As Long
    
    Const HDSet As String = "ぁあぃいぅうぇえぉおっつゃやゅゆょよゎわ"
    Const KDSet As String = "ァアィイゥウェエォオッツャヤュユョヨヮワ"
    Const KSSet As String = "ｧｱｨｲｩｳｪｴｫｵｯﾂｬﾔｭﾕｮﾖ"
    
    For i = 0 To Len(str) - 1
        c = Mid$(str, i + 1, 1)
        If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Then
            ' アルファベットの場合
            out = out & "[" & StrConv(c, vbProperCase) & StrConv(c, vbLowerCase) & "]"
        Else
            posHD = InStr(1, HDSet, c) ' 含まれていなければ0
            posKD = InStr(1, KDSet, c)
            posKS = InStr(1, KSSet, c)
            
            ' グループ化対象日本文字かどうか
            If posHD + posKD + posKS > 0 Then
                ' 0 オリジンのカーソル位置を得る
                posG = (posHD + posKD + posKS - 1) - (posHD + posKD + posKS - 1) Mod 2
                out = out & "[" & _
                    Mid$(HDSet, 1 + posG, 2) & _
                    Mid$(KDSet, 1 + posG, 2) & _
                    Mid$(KSSet, 1 + posG, 2) & "]"
            Else
                ' それ以外の文字はなにもしないでコピー
                out = out & c
            End If
        End If
    Next i
    CharGroup = out
End Function


'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    Set mCC = New clsCodeConverter
    Set mSC = New clsStringConverter
    Set mRX = New RegExp
End Sub


'
'   機能: クラス終了イベント
'
'   備考: なし
'
Private Sub Class_Terminate()
    Set mCC = Nothing
    Set mSC = Nothing
    Set mRX = Nothing
End Sub



'
'   機能: 馬主マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_BANUSI_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    
    gApp.Log "Check Cancel: ExecuteComplete " & mblnCancelFetching
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        mblnNowCanceling = False
        Exit Sub
    End If
    

    If pError Is Nothing Then
        Set mRS_BANUSI = pRecordset
        
        If mRS_BANUSI.EOF Then
            RaiseEvent NoDataBANUSI
            mblnNowFetching = False
        Else
            Call MakeData_BANUSI
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataBANUSI
        mblnNowFetching = False
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 調教師マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_CHOKYO_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    
    gApp.Log "Check Cancel: ExecuteComplete " & mblnCancelFetching
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        mblnNowCanceling = False
        Exit Sub
    End If
    

    If pError Is Nothing Then
        Set mRS_CHOKYO = pRecordset
        
        If mRS_CHOKYO.EOF Then
            RaiseEvent NoDataCHOKYO
            mblnNowFetching = False
        Else
            Call MakeData_CHOKYO
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataCHOKYO
        mblnNowFetching = False
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 繁殖馬マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_HANSYOKU_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    
    gApp.Log "Check Cancel: ExecuteComplete " & mblnCancelFetching
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        mblnNowCanceling = False
        Exit Sub
    End If
    

    If pError Is Nothing Then
        Set mRS_HANSYOKU = pRecordset
        
        If mRS_HANSYOKU.EOF Then
            RaiseEvent NoDataHANSYOKU
            mblnNowFetching = False
        Else
            Call MakeData_HANSYOKU
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataHANSYOKU
        mblnNowFetching = False
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 騎手マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_KISHU_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    
    gApp.Log "Check Cancel: ExecuteComplete " & mblnCancelFetching
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        mblnNowCanceling = False
        Exit Sub
    End If
    

    If pError Is Nothing Then
        Set mRS_KISHU = pRecordset
        
        If mRS_KISHU.EOF Then
            RaiseEvent NoDataKISHU
            mblnNowFetching = False
        Else
            Call MakeData_KISHU
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataKISHU
        mblnNowFetching = False
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 生産者マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_SEISAN_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    
    gApp.Log "Check Cancel: ExecuteComplete " & mblnCancelFetching
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        mblnNowCanceling = False
        Exit Sub
    End If
    
    If pError Is Nothing Then
        Set mRS_SEISAN = pRecordset
        
        If mRS_SEISAN.EOF Then
            RaiseEvent NoDataSEISAN
            mblnNowFetching = False
        Else
            Call MakeData_SEISAN
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataSEISAN
        mblnNowFetching = False
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 競走馬マスタレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_UMA_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler

    gApp.Log "Check Cancel: ExecuteComplete " & mblnCancelFetching
    If mblnCancelFetching Then
        mblnNowFetching = False
        mblnCancelFetching = False
        mblnNowCanceling = False
        Exit Sub
    End If

    If pError Is Nothing Then
        Set mRS_UMA = pRecordset
        
        If mRS_UMA.EOF Then
            RaiseEvent NoDataUMA
            mblnNowFetching = False
        Else
            Call MakeData_UMA
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataUMA
        mblnNowFetching = False
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: コネクションを開放する
'
'   備考: なし
'
Private Sub freecn(cn As ADODB.Connection)

    If Not cn Is Nothing Then
        Do While cn.State And adStateExecuting
            Call cn.Cancel
            gApp.Log "freecn Cancel"
        Loop
        cn.Close
        Set cn = Nothing
    Else
        gApp.Log "freecn Nothing"
    End If
End Sub

