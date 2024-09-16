VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSOdds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   OD Viewer用クラス
'
Option Explicit

Implements clsIViewerState

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngLastTabNumber   As Long         '' 最終に選択していたタブ番号保存
Private mlngAxisNum         As Long         '' 最終に選択していたAxisNum保存　3連複
Private mlngAxisNinki       As Long         '' 最終に選択していたAxisNinki保存　3連複
Private mlngAxisNumST       As Long         '' 最終に選択していたAxisNumST保存　3連単
Private mlngAxisNinkiST     As Long         '' 最終に選択していたAxisNinkiST保存　3連単
Private mblnOddsHyosuIsOdds As Boolean      '' ラジオボタン・オッズ票数が、オッズであれば真
Private mblnNumNinkiIsNum   As Boolean      '' ラジオボタン・番号人気が、番号であれば真
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


'
'   機能: 最終に選択していたAxisNum取得
'
'   備考: なし
'
Public Property Get AxisNum() As Long
    AxisNum = mlngAxisNum
End Property


'
'   機能: 最終に選択していたAxisNum保存
'
'   備考: なし
'
Public Property Let AxisNum(RHS As Long)
    mlngAxisNum = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: 最終に選択していたAxisNinki取得
'
'   備考: なし
'
Public Property Get AxisNinki() As Long
    AxisNinki = mlngAxisNinki
End Property


'
'   機能: 最終に選択していたAxisNinki保存
'
'   備考: なし
'
Public Property Let AxisNinki(RHS As Long)
    mlngAxisNinki = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: 最終に選択していたAxisNumST取得
'
'   備考: なし
'
Public Property Get AxisNumST() As Long
    AxisNumST = mlngAxisNumST
End Property


'
'   機能: 最終に選択していたAxisNumST保存
'
'   備考: なし
'
Public Property Let AxisNumST(RHS As Long)
    mlngAxisNumST = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: 最終に選択していたAxisNinkiST取得
'
'   備考: なし
'
Public Property Get AxisNinkiST() As Long
    AxisNinkiST = mlngAxisNinkiST
End Property


'
'   機能: 最終に選択していたAxisNinkiST保存
'
'   備考: なし
'
Public Property Let AxisNinkiST(RHS As Long)
    mlngAxisNinkiST = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: ラジオボタン・オッズ票数が、オッズであれば真をセット
'
'   備考: なし
'
Public Property Let OddsHyosuIsOdds(RHS As Boolean)
    mblnOddsHyosuIsOdds = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: ラジオボタン・オッズ票数が、オッズであれば真を返す
'
'   備考: なし
'
Public Property Get OddsHyosuIsOdds() As Boolean
    OddsHyosuIsOdds = mblnOddsHyosuIsOdds
End Property


'
'   機能: ラジオボタン・馬番人気が人気であれば真をセット
'
'   備考: なし
'
Public Property Let NumNinkiIsNum(RHS As Boolean)
    mblnNumNinkiIsNum = RHS
    mblnIsNoTouch = False
End Property


'
'   機能: ラジオボタン・馬番人気が、人気であれば真を返す
'
'   備考: なし
'
Public Property Get NumNinkiIsNum() As Boolean
    NumNinkiIsNum = mblnNumNinkiIsNum
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


