VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   集合体 クラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Dim mfso                                As FileSystemObject

Dim mblnBOF                             As Boolean
Dim mblnEOF                             As Boolean
Dim mstrBeginDate                       As String
Dim mstrEndDate                         As String
Dim mstrJyoCode                         As String
Dim mstrRecType                         As String

Private mFoldersYYYY                    As Collection
Dim mFileYYYYMMDDJJRR                   As String
Dim mContinue                           As Integer
Dim mHSearch                            As Long ' Search Handle
Dim mWFD                                As WIN32_FIND_DATA

Private mlngCIYYYY                      As Long
Private mlngCIYYYYMMDDJJRR              As Long

Private mblnCancelFlag                  As Boolean


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 開始日付をセット
'
'   備考: なし
'
Public Property Get BeginDate() As String
    BeginDate = mstrBeginDate
End Property

'
'   機能: 開始日付を取得
'
'   備考: なし
'
Public Property Let BeginDate(strParam As String)
    mstrBeginDate = strParam
End Property

'
'   機能: 終了日付をセット
'
'   備考: なし
'
Public Property Get EndDate() As String
    EndDate = mstrEndDate
End Property

'
'   機能: 終了日付をセット
'
'   備考: なし
'
Public Property Let EndDate(strParam As String)
    mstrEndDate = strParam
End Property

'
'   機能: 開催場コードを取得
'
'   備考: なし
'
Public Property Get JyoCode() As String
    JyoCode = mstrJyoCode
End Property

'
'   機能: 開催場コードをセット
'
'   備考: なし
'
Public Property Let JyoCode(strParam As String)
    mstrJyoCode = strParam
End Property

'
'   機能: レコードタイプを取得
'
'   備考: なし
'
Public Property Get RecordType() As String
    RecordType = mstrRecType
End Property

'
'   機能: レコードタイプをセット
'
'   備考: なし
'
Public Property Let RecordType(strParam As String)
    mstrRecType = strParam
End Property

'
'   機能: BOFの状態を取得
'
'   備考: なし
'
Public Property Get BOF() As Boolean
    BOF = mblnBOF
End Property

'
'   機能: EOFの状態を取得
'
'   備考: なし
'
Public Property Get EOF() As Boolean
    EOF = mblnEOF
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データ読み込み
'
'   備考: なし
'
Public Function GetData() As String
On Error GoTo ErrorHandler
    Dim intFileNum As Integer

    Dim YYYY          As String
    Dim YYYYMMDDJJRR    As String

    Dim Path     As String
    Dim filename As String
    Dim buf      As String

    YYYY = mFoldersYYYY.item(mlngCIYYYY).Name
    YYYYMMDDJJRR = mFileYYYYMMDDJJRR

    gApp.Log mFileYYYYMMDDJJRR

    Path = gApp.R_DBPath & "\ODDS\" & YYYY & "\" & YYYYMMDDJJRR

    intFileNum = FreeFile
    Open Path For Binary Access Read As intFileNum
    buf = String$(131071, " ") ' 2^17 bytes is enough for H6
    Get #intFileNum, , buf
    Close intFileNum

    GetData = buf
    Exit Function
ErrorHandler:
    gApp.ErrLog

End Function


'
'   機能: 次のファイルを探す
'
'   備考: なし
'
Public Sub MoveNext()
On Error GoTo ErrorHandler
    Dim strSearch As String

    Do
        DoEvents
        If mblnCancelFlag Then
            Exit Do
        End If
        mContinue = FindNextFile(mHSearch, mWFD)
        mFileYYYYMMDDJJRR = StripNulls(mWFD.cFileName)
        If mContinue <= 0 Then
            mContinue = FindClose(mHSearch)
            If mlngCIYYYY < mFoldersYYYY.count Then
                mlngCIYYYY = mlngCIYYYY + 1
                strSearch = gApp.R_DBPath & "\ODDS\" & mFoldersYYYY.item(mlngCIYYYY).Name & "\*.*"
                mHSearch = FindFirstFile(strSearch, mWFD)
                mFileYYYYMMDDJJRR = StripNulls(mWFD.cFileName)
                mContinue = True
            Else
                mblnEOF = True
                Exit Do
            End If
        End If
    Loop While mFileYYYYMMDDJJRR = "." Or mFileYYYYMMDDJJRR = ".." Or Match() = False

    mblnBOF = False

    Exit Sub

ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 最初のファイル
'
'   備考: なし
'
Public Sub MoveFirst()
On Error GoTo ErrorHandler
    Dim strSearch As String
    
    strSearch = gApp.R_DBPath & "\ODDS\" & mFoldersYYYY.item(mlngCIYYYY).Name & "\*.*"
    mHSearch = FindFirstFile(strSearch, mWFD)
    mFileYYYYMMDDJJRR = StripNulls(mWFD.cFileName)
    Do While mFileYYYYMMDDJJRR = "." Or mFileYYYYMMDDJJRR = ".." Or Match() = False

        DoEvents
        If mblnCancelFlag Then
            Exit Do
        End If

        mContinue = FindNextFile(mHSearch, mWFD)
        mFileYYYYMMDDJJRR = StripNulls(mWFD.cFileName)
        If mContinue <= 0 Then
            mContinue = FindClose(mHSearch)
            If mlngCIYYYY < mFoldersYYYY.count Then
                mlngCIYYYY = mlngCIYYYY + 1
                strSearch = gApp.R_DBPath & "\ODDS\" & mFoldersYYYY.item(mlngCIYYYY).Name & "\*.*"
                mHSearch = FindFirstFile(strSearch, mWFD)
                mFileYYYYMMDDJJRR = StripNulls(mWFD.cFileName)
                mContinue = True
            Else
                mblnEOF = True
                Exit Sub
            End If
        End If
    Loop

    mContinue = True

    mblnBOF = True

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 初期化
'
'   備考: なし
'
Public Sub Init()
On Error GoTo ErrorHandler
    Dim DBFolder As Folder
    Set DBFolder = mfso.GetFolder(gApp.R_DBPath & "\ODDS\")
    Set mFoldersYYYY = FoldersToCollection(DBFolder.SubFolders)

    mlngCIYYYY = 1
    mlngCIYYYYMMDDJJRR = 1

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: キャンセル
'
'   備考: なし
'
Public Sub Cancel()
    mblnCancelFlag = True
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
On Error GoTo ErrorHandler
    Set mfso = New FileSystemObject

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub

'
'   機能: クラス終了イベント
'
'   備考: なし
'
Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    Set mfso = Nothing

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub

'
'   機能: コレクションに加える
'
'   備考: なし
'
Private Function FoldersToCollection(ByRef f As Folders) As Collection
    Dim c As New Collection
    Dim tmp As Folder

    For Each tmp In f
        If Left$(tmp.Name, 2) = mstrRecType Then
            c.Add tmp
        End If
    Next tmp

    Set FoldersToCollection = c
End Function

'
'   機能: 開催情報の検査
'
'   備考: なし
'
Private Function Match() As Boolean

    Dim JyoRace As String
    Dim JyoCode As String

    Dim JyoCD As String

    Dim YYYYMMDDJJRR As String

    Dim i As Long
    
    Dim YearMonthDay As String

    Dim yyyymmdd1 As String
    Dim yyyymmdd2 As String

    yyyymmdd1 = Format$(mstrBeginDate, "YYYYMMDD")
    yyyymmdd2 = Format$(mstrEndDate, "YYYYMMDD")

    Match = False

    YYYYMMDDJJRR = mFileYYYYMMDDJJRR
    YearMonthDay = Mid$(YYYYMMDDJJRR, 1, 8)
    JyoRace = Mid$(YYYYMMDDJJRR, 9, 4)

    JyoCode = Mid$(JyoRace, 1, 2)

    If YearMonthDay >= yyyymmdd1 And YearMonthDay <= yyyymmdd2 Then
        Select Case mstrJyoCode
        Case 0:
            If 1 <= val(JyoCode) And val(JyoCode) <= 10 Then
                Match = True
            End If
        Case 1:
            If 11 <= val(JyoCode) And val(JyoCode) <= 99 Then
                Match = True
            End If
        Case 2:
            If "A0" <= JyoCode Then
                Match = True
            End If
        End Select
    End If

End Function


