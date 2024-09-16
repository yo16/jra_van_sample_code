VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "レコード データクラス"
'
'   レコード データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event FetchComplete(gd As clsGridData, col As Long)
Public Event NoData()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter          '' JV-Data のコード変換メソッド群
Private mstrLabels(0 To 3) As String     '' テキストデータ
Private mstrLinkLabels(0 To 1) As clsGridItem     '' リンクデータ
Private mKey As clsKeyRC                '' キー

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
Public Function Fetch(ByRef key As clsKeyRC) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rc As clsRCSearch
    Dim gd As clsGridData
    Dim strSQL As String
    Dim i As Long
    Dim lngCols As Long
    
    Set mKey = key

    Set cn = gApp.GetCN_RECORD
    
    'RECORD (レコードマスタ)
    strSQL = "SELECT * FROM RECORD " & key.SQLWHEREString
    Set rs = cn.Execute(strSQL)
    
    ' データが無かったらFalseを返して終了
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        RaiseEvent NoData
        Exit Function
    End If
    
    mstrLabels(0) = mSC.YMD1(rs("MakeDate")) & "作成データ"   '作成年月日
    
    ' 太字レース名等グレーバー用文字列生成
    If rs("RecInfoKubun") = "1" Then
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & mCC.KIBJ3(rs("JyoCD")) & " "
        mstrLabels(1) = mstrLabels(1) & val(rs("Kyori")) & "m "
        mstrLabels(1) = mstrLabels(1) & mCC.TRCK1(rs("TrackCD")) & " "
        mstrLabels(1) = mstrLabels(1) & mCC.KSSB4(Mid$(rs("TokuNum_SyubetuCD"), 5, 2)) & " "
    ElseIf rs("RecInfoKubun") = "2" Then
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & Trim$(rs("Hondai"))
        mstrLabels(1) = mstrLabels(1) & mCC.GRAD2(rs("GradeCD"))
        mstrLabels(1) = mstrLabels(1) & mCC.KIBJ3(rs("JyoCD")) & " "
        mstrLabels(1) = mstrLabels(1) & val(rs("Kyori")) & "m "
        mstrLabels(1) = mstrLabels(1) & mCC.TRCK1(rs("TrackCD")) & " "
        mstrLabels(1) = mstrLabels(1) & mCC.KSSB4(Mid$(rs("TokuNum_SyubetuCD"), 5, 2)) & " "
    End If
    
    mstrLabels(2) = ""
    mstrLabels(2) = mstrLabels(2) & mSC.YMD1(rs("Year") & rs("MonthDay")) & " "
    mstrLabels(2) = mstrLabels(2) & IIf(val(rs("RaceNum")) = 0, "", val(rs("RaceNum")) & "R　")
    mstrLabels(2) = mstrLabels(2) & mCC.TNKO1(rs("TenkoCD")) & " "
    mstrLabels(2) = mstrLabels(2) & IIf(rs("SibaBabaCD") = "0", "", "芝：" & _
                                      mCC.BBJT1(rs("SibaBabaCD"))) & _
                                      IIf(rs("DirtBabaCD") = "0", "", " ダ：" & _
                                      mCC.BBJT1(rs("DirtBabaCD")))
                                      
    ' リンクデータオブジェクトを生成する
    For i = 0 To 1
        Set mstrLinkLabels(i) = New clsGridItem
    Next i
    Set rc = New clsRCSearch
    Set rc.CurrentRecordKey = key

    ' リンク用文字列生成
    If Not rc.NextRecordKey(False) Is Nothing Then
        With rc.NextRecordKey(False)
            mstrLinkLabels(0).Text = Trim$(mSC.YMD1(.Year & .MonthDay) & " " & _
                                        IIf(val(.RaceNum) = 0, "", val(.RaceNum) & "R　"))
        End With
        mstrLinkLabels(0).Link = "RC"
        mstrLinkLabels(0).key = rc.NextRecordKey(False).str
    End If
    If Not rc.PreviousRecordKey(False) Is Nothing Then
        With rc.PreviousRecordKey(False)
            mstrLinkLabels(1).Text = Trim$(mSC.YMD1(.Year & .MonthDay) & " " & _
                                        IIf(val(.RaceNum) = 0, "", val(.RaceNum) & "R　"))
        End With
        mstrLinkLabels(1).Link = "RC"
        mstrLinkLabels(1).key = rc.PreviousRecordKey(False).str
    End If
    
    ' 以下Grid処理
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData

    With gd
        .Cols = 4
        .Rows = 5
        .SetItemMatrix 0, 0, "レースタイム"
        .SetItemMatrix 1, 0, "馬名"
        .SetItemMatrix 2, 0, "馬性別"
        .SetItemMatrix 3, 0, "騎手"
        .SetItemMatrix 4, 0, "調教師"

        For i = 0 To 2
            If Len(RTrim$(rs("RecUmaBamei" & i + 1))) <> 0 Then
                .SetItemMatrix 0, i + 1, IIf(val(rs("RecTime")) = 0, "", Format$(rs("RecTime"), "@分@@秒@"))
                .SetItemMatrix 1, i + 1, Trim$(rs("RecUmaBamei" & i + 1))
                .SetItemMatrix 2, i + 1, mCC.SEIB4(rs("RecUmaSexCD" & i + 1))
                .SetItemMatrix 3, i + 1, Trim$(rs("RecUmaKisyuName" & i + 1)) & "(" & Format$(rs("RecUmaFutan" & i + 1) / 10, "0.0") & ")"
                .SetItemMatrix 4, i + 1, Trim$(rs("RecUmaChokyosiName" & i + 1))
                lngCols = i + 2
            End If
        Next i
    End With
    
    rs.Close
    
    Set rs = Nothing
    
    RaiseEvent FetchComplete(gd, lngCols)
    
    Fetch = True
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

