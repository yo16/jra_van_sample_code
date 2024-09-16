VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistoryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   履歴アイテム クラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Key As String
Public ViewerName As String
Public Title As String
Public DateTime As String
Public ViewerState As clsIViewerState

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 履歴情報文字列の作成
'
'   備考: なし
'
Public Function ToString() As String
    ToString = "[" & Key & "]" _
            & ",[" & ViewerName & "]" _
            & ",[" & Title & "]" _
            & ",[" & DateTime & "]"

End Function

