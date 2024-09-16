VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSTabOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   カレントタブ情報のみ持つViewer用クラス
'

Option Explicit

Implements clsIViewerState

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngLastTabNumber   As Long         '' 最終に選択していたタブ番号保存
Private mblnIsNoTouch       As Boolean      '' 状態は変更されていない場合真



'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: 状態は変更されていない場合真
'
'   備考: なし
'
Public Property Get IsNoTouch() As Boolean
    IsNoTouch = mblnIsNoTouch
End Property


'
'   機能: 最終に選択していたタブ番号取得
'
'   備考: なし
'
Public Property Get LastTabNumber() As Long
    LastTabNumber = mlngLastTabNumber
End Property


'
'   機能: 最終に選択していたタブ番号保存
'
'   備考: なし
'
Public Property Let LastTabNumber(RHS As Long)
    mlngLastTabNumber = RHS
    mblnIsNoTouch = False
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
    mblnIsNoTouch = True
End Sub
