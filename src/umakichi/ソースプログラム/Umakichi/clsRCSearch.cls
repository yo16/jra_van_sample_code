VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   レコード検索 クラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mCn     As ADODB.Connection
Private mKey    As clsKeyRC


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: レコード検索キークラス 代入プロパティ
'
'   備考: なし
'
Public Property Set CurrentRecordKey(ByRef RHS As clsKeyRC)
    Set mKey = RHS
End Property


'
'   機能: レコード検索キークラス 取得プロパティ
'
'   備考: なし
'
Public Property Get CurrentRecordKey() As clsKeyRC
    Set CurrentRecordKey = mKey
End Property


'
'   機能: レコード検索キー文字列 代入プロパティ
'
'   備考: なし
'
Public Property Let CurrentRecordKeyStr(RHS As String)
    Set mKey = New clsKeyRC
    mKey.Str = RHS
End Property


'
'   機能: レコード検索キー文字列 取得プロパティ
'
'   備考: なし
'
Public Property Get CurrentRecordKeyStr() As String
     CurrentRecordKeyStr = mKey.Str
End Property


'
'   機能: レコード検索キークラス 取得プロパティ
'
'   備考: なし
'
Public Property Get PreviousRecordKey(opt As Boolean) As clsKeyRC
    Set PreviousRecordKey = GetPreviousRecordKey(opt)
End Property


'
'   機能: レコード検索キークラス 取得プロパティ
'
'   備考: なし
'
Public Property Get NextRecordKey(opt As Boolean) As clsKeyRC
    Set NextRecordKey = GetNextRecordKey(opt)
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: レコードのキーを返す
'
'   備考: 引き数 opt -true:現在のキーがあれば現在のキーを返す
'                    -false:現在のキーがあっても前回のキーを返す
'
Private Function GetPreviousRecordKey(opt As Boolean) As clsKeyRC
On Error GoTo ErrorHandler
    Dim newKey As clsKeyRC
    
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Call OpenTableDirect(rs, mCn, "Record")
    
    ' 末端に移動
    If Not rs.EOF Then
        rs.MoveLast
    End If
    
    ' 未来分は読み飛ばす
    Do While Not rs.BOF
        If opt Then
            If rs("RecInfoKubun") & rs("Year") & rs("MonthDay") & rs("JyoCD") & _
            rs("Kaiji") & rs("Nichiji") & rs("RaceNum") & rs("TokuNum_SyubetuCD") & _
            rs("Kyori") & rs("TrackCD") <= mKey.Str Then
                Exit Do
            End If
        Else
            If rs("RecInfoKubun") & rs("Year") & rs("MonthDay") & rs("JyoCD") & _
            rs("Kaiji") & rs("Nichiji") & rs("RaceNum") & rs("TokuNum_SyubetuCD") & _
            rs("Kyori") & rs("TrackCD") < mKey.Str Then
                Exit Do
            End If
        End If
        rs.MovePrevious
    Loop

    ' 同条件の最も近い過去のレコードを探す
    Do While Not rs.BOF
        If mKey.RecInfoKubun = "1" Then
            If SameRecord(rs) Then
                Exit Do
            End If
        ElseIf mKey.RecInfoKubun = "2" Then
            If Mid$(rs("TokuNum_SyubetuCD"), 1, 4) _
                = Mid$(mKey.TokuNum_SyubetuCD, 1, 4) _
                And rs("RecInfoKubun") = mKey.RecInfoKubun Then
                Exit Do
            End If
        End If
        rs.MovePrevious
    Loop
    
    If rs.BOF Then
        Set GetPreviousRecordKey = Nothing
        Exit Function
    End If
    
    Set newKey = New clsKeyRC
    Call newKey.SetFromRS(rs)
    Set GetPreviousRecordKey = newKey
    Exit Function
ErrorHandler:
    gApp.Log "RCSearch::PreviousRecordKey Error"
    gApp.ErrLog
    Resume Next
End Function



'
'   機能: レコードのキーを返す
'
'   備考: 引き数 opt -true:現在のキーがあれば現在のキーを返す
'                    -false:現在のキーがあっても次回のキーを返す
'
Private Function GetNextRecordKey(opt As Boolean) As clsKeyRC
On Error GoTo ErrorHandler
    Dim newKey As clsKeyRC
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Call OpenTableDirect(rs, mCn, "Record")
    
    ' 先頭に移動
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    
    ' 過去分は読み飛ばす
    Do While Not rs.EOF
        If opt Then
            If rs("RecInfoKubun") & rs("Year") & rs("MonthDay") & rs("JyoCD") & _
            rs("Kaiji") & rs("Nichiji") & rs("RaceNum") & rs("TokuNum_SyubetuCD") & _
            rs("Kyori") & rs("TrackCD") >= mKey.Str Then
                Exit Do
            End If
        Else
            If rs("RecInfoKubun") & rs("Year") & rs("MonthDay") & rs("JyoCD") & _
            rs("Kaiji") & rs("Nichiji") & rs("RaceNum") & rs("TokuNum_SyubetuCD") & _
            rs("Kyori") & rs("TrackCD") > mKey.Str Then
                Exit Do
            End If
        End If
        rs.MoveNext
    Loop

    ' 同条件の最も近い未来のレコードを探す
    Do While Not rs.EOF
        If mKey.RecInfoKubun = "1" Then
            If SameRecord(rs) Then
                Exit Do
            End If
        ElseIf mKey.RecInfoKubun = "2" Then
            If Mid$(rs("TokuNum_SyubetuCD"), 1, 4) _
                = Mid$(mKey.TokuNum_SyubetuCD, 1, 4) _
                And rs("RecInfoKubun") = mKey.RecInfoKubun Then
                Exit Do
            End If
        End If
        rs.MoveNext
    Loop
    
    If rs.EOF Then
        Set GetNextRecordKey = Nothing
        Exit Function
    End If
    
    Set newKey = New clsKeyRC
    Call newKey.SetFromRS(rs)
    Set GetNextRecordKey = newKey
    Exit Function
ErrorHandler:
    gApp.Log "RCSearch::NextRecordKey Error"
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: キーとレコードセットを比較
'
'   備考: なし
'
Private Function SameRecord(rs As ADODB.Recordset) As Boolean
    If rs("RecInfoKubun") = mKey.RecInfoKubun _
    And rs("JyoCD") = mKey.JyoCD _
    And rs("Kyori") = mKey.KYORI _
    And rs("TrackCD") = mKey.TrackCD _
    And SyubetuType(rs("Year"), Mid$(rs("TokuNum_SyubetuCD"), 5, 2)) _
        = SyubetuType(mKey.Year, Mid$(mKey.TokuNum_SyubetuCD, 5, 2)) Then
        SameRecord = True
    Else
        SameRecord = False
    End If
    
End Function


'
'   機能: 種別タイプを返す
'
'   備考: なし
'
Private Function SyubetuType(ByVal Y As Long, SyubetuCD As String) As Long
    Select Case SyubetuCD
    Case "11"
        SyubetuType = 1
    Case "12" To "14"
        SyubetuType = 2
    Case "18" To "19"
        SyubetuType = 3
    Case Else
        SyubetuType = 0
    End Select
End Function


'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    Set mCn = gApp.GetCN_RECORD
End Sub


'
'   機能: クラス終了イベント
'
'   備考: なし
'
Private Sub Class_Terminate()
    Set mCn = Nothing
End Sub
