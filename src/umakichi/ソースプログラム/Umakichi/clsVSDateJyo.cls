VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSDateJyo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   日時場所行数の状態を持つViewer用クラス
'

Option Explicit

Implements clsIViewerState

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrYMD             As String       '' 年度
Private mstrJyo             As String       '' 場所
Private mlngRow             As Long         '' 行数
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
'   機能: YMDを得る
'
'   備考: なし
'
Public Property Get YMD() As String
    YMD = mstrYMD
End Property


'
'   機能: YMDをセットする
'
'   備考: なし
'
Public Property Let YMD(RHS As String)
    mstrYMD = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: Jyoを取得する
'
'   備考: なし
'
Public Property Get Jyo() As String
    Jyo = mstrJyo
End Property


'
'   機能: Jyoをセットする
'
'   備考: なし
'
Public Property Let Jyo(RHS As String)
    mstrJyo = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: ROWをセットする
'
'   備考: なし
'
Public Property Get row() As Long
    row = mlngRow
End Property


'
'   機能: ROWを得る
'
'   備考: なし
'
Public Property Let row(RHS As Long)
    mlngRow = RHS
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

