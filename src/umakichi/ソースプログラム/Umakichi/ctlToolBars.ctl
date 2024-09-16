VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlToolBars 
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   ScaleHeight     =   1365
   ScaleWidth      =   2265
   Begin MSComctlLib.Toolbar tbrInner 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlToolBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   ツールバー
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Public Event ButtonMenuClick(Index As Integer, ByVal ButtonMenu As MSComctlLib.ButtonMenu)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private msngMinWidth As Single

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ToolBarプロパティの取得
'
'   備考: なし
'
Public Property Get ToolBar(Index As Long) As ToolBar
    Set ToolBar = tbrInner(Index)
End Property

'
'   機能: numプロパティのセット
'
'   備考: なし
'
Public Property Let num(RHS As Long)
    Dim i As Long
    ' 足りない分をロードする
    For i = tbrInner.UBound + 1 To RHS - 1
        Load tbrInner(i)
        tbrInner(i).Visible = True
    Next i
End Property

'
'   機能: MinWidthプロパティの取得
'
'   備考: なし
'
Public Property Get MinWidth() As Single
    MinWidth = msngMinWidth
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ユーザコントロールの高さを調整
'
'   備考: なし
'
Public Sub fit()
    Call StandInLine
    UserControl.Height = tbrInner(0).ButtonHeight
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'
'   機能: MinWidthプロパティ値のセット
'
'   備考: なし
'
Private Sub StandInLine()
    Dim i As Long
    Dim p As Single
    
    For i = 0 To tbrInner.UBound
        tbrInner(i).Left = p
        tbrInner(i).width = CalcButtonsWidth(tbrInner(i))
        If tbrInner(i).Visible = True Then
            p = p + tbrInner(i).width
        End If
    Next i
    msngMinWidth = p
End Sub

'
'   機能: ボタンの幅を返す
'
'   備考: なし
'
Private Function CalcButtonsWidth(tbrTarget As ToolBar) As Single
    Dim p As Single
    Dim i As Long
    For i = 1 To tbrTarget.Buttons.count
        p = p + tbrTarget.Buttons(i).width
    Next i
    CalcButtonsWidth = p
End Function

'
'   機能: 内包するツールバーのボタンクリックイベント
'
'   備考: なし
'
Private Sub tbrInner_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler
    RaiseEvent ButtonClick(Index, Button)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: 内包するツールバーのボタンメニュークリックイベント
'
'   備考: なし
'
Private Sub tbrInner_ButtonMenuClick(Index As Integer, ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrorHandler
    RaiseEvent ButtonMenuClick(Index, ButtonMenu)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: ユーザコントロールのリサイズイベント
'
'   備考: なし
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Call StandInLine
    UserControl.Height = tbrInner(0).ButtonHeight
    UserControl.width = msngMinWidth * 10
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
