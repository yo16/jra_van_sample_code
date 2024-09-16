VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabaseMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "データベースマネージャ"
'
'   データベースマネージャ
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データベースに接続する。DBが無い場合、選択させる。
'
'   備考: なし
'
Public Function Connect() As Boolean
On Error GoTo Errorhandler

    Dim result As Long      '' MsgBoxの戻り値
    Dim DirRef As frmDirRef '' フォルダ参照ダイアログ
    
    Set DirRef = New frmDirRef
    DirRef.Message = "データベースフォルダを指定してください"
    
    ' データベースがみつかるまでループ
    Do While MissingDBNum <> 0
        DoEvents
        ' なければ、フォルダ選択
        DirRef.BeginingPath = gApp.R_DBPath
        DirRef.Show vbModal
        
        gApp.R_DBPath = DirRef.ReturnPath
        ' データベースがあるか調べる
        Select Case MissingDBNum
        Case -1
            result = MsgBox(gApp.R_DBPath & "に、データベースを新規作成しますか？" _
                            , vbYesNoCancel, "馬吉：新規データベース作成の確認")
            If result = vbYes Then
                Call createDB
            ElseIf result = vbCancel Then
                Connect = False
                Exit Function
            End If
        Case Is > 0
            MsgBox "データベースが壊れています。他のフォルダを指定してください。", vbExclamation, "馬吉:エラー"
            ' なければ、フォルダ選択
            DirRef.BeginingPath = gApp.R_DBPath
            DirRef.Show vbModal
            
            gApp.R_DBPath = DirRef.ReturnPath
        End Select
    Loop
    
    Connect = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    Connect = False
End Function


'
'   機能: 切断する
'
'   備考: なし
'
Public Function Disconnect() As Boolean
On Error GoTo Errorhandler
    
    Exit Function
Errorhandler:
    Disconnect = False
End Function


'
'   機能: 初期化
'
'   備考: なし
'
Private Sub Class_Initialize()
On Error GoTo Errorhandler

    Exit Sub
Errorhandler:
    Debug.Print Err.Description
    Resume Next
End Sub


'
'   機能: データベースの不足数を返す
'
'   備考: なし
'
Public Function MissingDBNum() As Long
    Dim SpecIDs()   As String
    Dim dbpath      As String
    Dim fso         As FileSystemObject
    Dim count       As Long
    Dim i           As Long

    gApp.Log "FSO New"
    Set fso = New FileSystemObject
    gApp.Log "FSO New finish"
    dbpath = gApp.R_DBPath

    SpecIDs = Split("BANUSI BATAIJYU CHOKYO CHOKYO_SEISEKI HANRO HANSYOKU HARAI KISHU KISHU_CHANGE KISHU_SEISEKI MINING ODDS_SANREN0 ODDS_SANREN1 ODDS_SANREN2 ODDS_SANREN3 ODDS_SANREN4 ODDS_SANREN5 ODDS_SANREN6 ODDS_SANREN7 ODDS_SANREN8 ODDS_SANREN9 ODDS_TANPUKUWAKU ODDS_UMAREN ODDS_UMATAN0 ODDS_UMATAN1 ODDS_UMATAN2 ODDS_UMATAN3 ODDS_UMATAN4 ODDS_UMATAN5 ODDS_UMATAN6 ODDS_UMATAN7 ODDS_UMATAN8 ODDS_UMATAN9 ODDS_WIDE RACE RECORD SANKU SCHEDULE SEISAN TENKO_BABA TOKU TOKU_RACE TORIKESI_JYOGAI UMA UMA_RACE_A UMA_RACE_B")

    For i = 0 To UBound(SpecIDs)
        If Not fso.FileExists(dbpath & "\sub" & SpecIDs(i) & ".mdb") Then
            count = count + 1
        End If
    Next i

    If count = UBound(SpecIDs) + 1 Then
        count = -1
    End If

    MissingDBNum = count
End Function


'
'   機能: データベースを作成する
'
'   備考: なし
'
Public Sub createDB()
On Error GoTo Errorhandler
    Dim cmdb As New clsCreateMDB
    cmdb.createMDB
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: コネクションを作成しオブジェクトを返す
'
'   備考: なし
'
Public Function GetCN(num As Long) As ADODB.Connection
On Error GoTo Errorhandler
    Dim cn As New ADODB.Connection
    Dim sngStartTimer As Single
    Dim lngNumLoop As Long
    
    sngStartTimer = Timer
    
    cn.Mode = adModeRead
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & gApp.R_DBPath & "\" & gstrMDBName(num)
    Set GetCN = cn
    
    If lngNumLoop > 0 Then
        gApp.Log "Get Connection Num Loop: " & lngNumLoop
    End If
    
    Exit Function
Errorhandler:
    DoEvents
    If sngStartTimer + 30 > Timer Then
        lngNumLoop = lngNumLoop + 1
        Resume
    End If
    gApp.ErrLog
    Resume Next
End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クラスターミネートイベント
'
'   備考: なし
'
Private Sub Class_Terminate()
On Error GoTo Errorhandler

    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub

