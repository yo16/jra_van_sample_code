VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistoryMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   履歴管理クラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngCurrent As Long
Private mPathHistory As Collection      '' リンクパス履歴


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 現在番号の取得
'
'   備考: なし
'
Public Property Get CurrentNum() As Long
    CurrentNum = mlngCurrent
End Property

'
'   機能: 現在番号をセット
'
'   備考: なし
'
Public Property Get Current(Optional step As Long = 0) As clsHistoryItem
    Dim Index As Long
    Index = mlngCurrent + step
    If Index >= 1 And Index <= mPathHistory.count Then
        Set Current = mPathHistory(Index)
    End If
End Property


'
'   機能: 最古なら真
'
'   備考: なし
'
Public Property Get IsFirst() As Boolean
    IsFirst = (mlngCurrent = 1)
End Property


'
'   機能: 最新なら真
'
'   備考: なし
'
Public Property Get IsLast() As Boolean
    IsLast = (mlngCurrent = mPathHistory.count)
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    mlngCurrent = 0
    Set mPathHistory = New Collection
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: 履歴に追加する
'
'   備考: なし
'
Public Sub Add(ByRef newHistory As clsHistoryItem)
On Error GoTo ErrorHandler
    
    'Deleate feauturs
    Do While mPathHistory.count > mlngCurrent
        Call mPathHistory.Remove(mlngCurrent + 1)
    Loop
    
    Call mPathHistory.Add(newHistory)
    
    mlngCurrent = mPathHistory.count
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 履歴を移動する
'
'   備考: なし
'
Public Sub Move(step As Long)
    mlngCurrent = mlngCurrent + step
    If mlngCurrent < 1 Then
        mlngCurrent = 1
    End If
    If mlngCurrent > mPathHistory.count Then
        mlngCurrent = mPathHistory.count
    End If
End Sub

