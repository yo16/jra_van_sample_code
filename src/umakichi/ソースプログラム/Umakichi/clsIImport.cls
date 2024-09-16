VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   データベース登録クラス共通インターフェイス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 追加処理
'
'   備考: なし
'
Public Function Add(lBUf() As Byte) As Boolean
End Function


'
'   機能: 開始処理
'
'   備考: なし
'
Public Sub OpenDB()
End Sub



'
'   機能: 終了処理
'
'   備考: なし
'
Public Sub CloseDB()
End Sub

