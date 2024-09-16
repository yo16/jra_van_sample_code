VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSYearJyo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   年度場所行数の状態を持つViewer用クラス
'   開催選択等
'

Option Explicit

Implements clsIViewerState

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrYear            As String       '' 年度
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
'   機能: Yearを得る
'
'   備考: なし
'
Public Property Get YEAR() As String
    YEAR = mstrYear
End Property


'
'   機能: Yearを保存
'
'   備考: なし
'
Public Property Let YEAR(RHS As String)
    mstrYear = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: Jyoを得る
'
'   備考: なし
'
Public Property Get Jyo() As String
    Jyo = mstrJyo
End Property


'
'   機能: Jyoを保存
'
'   備考: なし
'
Public Property Let Jyo(RHS As String)
    mstrJyo = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: Rowを得る
'
'   備考: なし
'
Public Property Get Row() As Long
    Row = mlngRow
End Property


'
'   機能: Rowを保存
'
'   備考: なし
'
Public Property Let Row(RHS As Long)
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

