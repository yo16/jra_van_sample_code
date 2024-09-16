VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportODDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   JVData オッズ票数 データベース登録クラス
'   "H1", "O1", "O2", "O3", "O4", "O5"
'   オッズ票数は、MDBを使用せず、ファイルに書き出します。
'
Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrBasePath As String              '' 基準パス
Private fso          As FileSystemObject    '' FSO


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クラス初期化イベント
'
'   備考: ファイルシステムオブジェクトのインスタンス生成
'
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    ' FSO インスタンスの生成
    Set fso = New FileSystemObject
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: フォルダパス生成、フォルダ作成
'
'   備考: なし
'
Private Sub clsIImport_OpenDB()
On Error GoTo ErrorHandler        ' コネクションオープン
    
    
    ' 保存フォルダ基準パスの作成
    mstrBasePath = gApp.R_DBPath & "\ODDS\"
    
    ' 基準フォルダが存在しない場合は作成
    If Not fso.FolderExists(mstrBasePath) Then
        MkDir mstrBasePath
        gApp.Log "基準フォルダ作成 (" & mstrBasePath & ")"
    End If

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: ダミー
'
'   備考: なし
'
Private Sub clsIImport_CloseDB()
On Error GoTo ErrorHandler
    ' Empty
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: JVReadの返す１行をデータベースに登録する
'
'   備考: DBに追加を試みる
'
Private Function clsIImport_Add(lBuf() As Byte) As Boolean
On Error GoTo ErrorHandler

    If lBuf(2) = ASCII_ZERO Then
        Call DeleteRecord(StrConv(lBuf, vbUnicode))
        clsIImport_Add = True
    Else

        clsIImport_Add = ImportToFile(StrConv(lBuf, vbUnicode))

    End If

    Exit Function

ErrorHandler:
    gApp.ErrLog
    clsIImport_Add = False
End Function


'
'   機能: レコードを削除する
'
'   備考: なし
'
Private Sub DeleteRecord(lBuf As String)
On Error GoTo ErrorHandler
    Dim strFolder1  As String
    Dim strFileName As String
    
    'フォルダ名の作成
    ' 基準パス
    strFolder1 = mstrBasePath & "\" & Mid$(lBuf, 1, 2) & Mid$(lBuf, 12, 4)     ' \区分年
    'ファイル名の作成
    ' YYYYMMDDJJRR
    strFileName = Mid$(lBuf, 8, 14) & Mid$(lBuf, 26, 2) ' YearMonthDayJyoCDRaceNum をファイル名とする

    ' すでにフォルダが存在しているかどうか
    If fso.FileExists(strFolder1 & "\" & strFileName) Then
        ' 削除
        Call fso.DeleteFile(strFolder1 & "\" & strFileName)
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データベースに追加する
'
'   備考: なし
'
Private Function ImportToFile(ByRef lBuf As String) As Boolean
On Error GoTo ErrorHandler

    Dim strFolder1       As String
    Dim strFileName      As String
    Dim strExistFileName As String
    Dim intFileNumber    As Integer
    Dim buf              As String
    Dim MakeDateOld      As String
    Dim MakeDateNew      As String
    
    'フォルダ名の作成
    ' 基準パス
    strFolder1 = mstrBasePath & "\" & Mid$(lBuf, 1, 2) & Mid$(lBuf, 12, 4)     ' \区分年
    
    'ファイル名の作成
    ' IDKYYYYMMDDYYYYMMDDJJKKNNRR
    '    4   8   12            26
    strFileName = Mid$(lBuf, 12, 10) & Mid$(lBuf, 26, 2) ' YearMonthDayJyoCDRaceNum をファイル名とする
    
    ' すでにファイルが存在しているかどうか
    If fso.FileExists(strFolder1 & "\" & strFileName) Then
        ' ファイルが存在している場合
        
        ' MakeDateの読み込み
        intFileNumber = FreeFile   ' 未使用のファイル番号を取得
        Open strFolder1 & "\" & strFileName For Binary Access Read As #intFileNumber
        buf = String$(131071, " ")
        Get #intFileNumber, , buf
        MakeDateOld = Mid$(buf, 4, 8)
        Close #intFileNumber
                
        MakeDateNew = Mid$(lBuf, 4, 8)
        
        ' MakeDate の比較
        If MakeDateOld > MakeDateNew Then
            ' 更新の必要が無い場合
            ImportToFile = True
            Exit Function ' 成功で終了
        End If
    End If
    
    If Not fso.FolderExists(strFolder1) Then
        ' フォルダが存在していない場合
        
        'フォルダの作成
        Call fso.CreateFolder(strFolder1)
    End If
    
    ' 書き込み
    intFileNumber = FreeFile   ' 未使用のファイル番号を取得
    
    Open strFolder1 & "\" & strFileName For Binary Access Write As #intFileNumber
    Put #intFileNumber, , lBuf
    Close #intFileNumber
        
    ImportToFile = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    ImportToFile = False
End Function

