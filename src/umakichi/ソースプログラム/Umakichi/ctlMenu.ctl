VERSION 5.00
Begin VB.UserControl ctlMenu 
   Alignable       =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   ScaleHeight     =   5970
   ScaleWidth      =   3165
   Begin VB.Frame fraA1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'なし
      Caption         =   "0"
      Height          =   2115
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2985
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   0
         Left            =   0
         Tag             =   "RAKaiSel"
         Top             =   0
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   1
         Left            =   1005
         Tag             =   "TKKaiSel"
         Top             =   0
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   2
         Left            =   0
         Tag             =   "HCSel"
         Top             =   525
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   3
         Left            =   1005
         Tag             =   "RCSel"
         Top             =   525
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   4
         Left            =   0
         Tag             =   "Home"
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   5
         Left            =   1005
         Tag             =   "UPDT"
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   6
         Left            =   0
         Tag             =   "SOKU"
         Top             =   1575
         Width           =   2010
      End
   End
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2280
      Top             =   2370
   End
   Begin VB.Frame fraA2 
      BackColor       =   &H00C0C0C0&
      Height          =   1770
      Left            =   -90
      TabIndex        =   0
      Top             =   2040
      Width           =   3165
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   7
         Left            =   90
         Tag             =   "Find0"
         Top             =   150
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   8
         Left            =   1095
         Tag             =   "Find1"
         Top             =   150
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   9
         Left            =   90
         Tag             =   "Find2"
         Top             =   675
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   10
         Left            =   1095
         Tag             =   "Find3"
         Top             =   675
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   11
         Left            =   90
         Tag             =   "Find4"
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Image imgCmd 
         Height          =   525
         Index           =   12
         Left            =   1095
         Tag             =   "Find5"
         Top             =   1200
         Width           =   1005
      End
   End
   Begin VB.Image imgCmd 
      Height          =   525
      Index           =   14
      Left            =   1005
      Tag             =   "HELP"
      Top             =   3825
      Width           =   1005
   End
   Begin VB.Image imgCmd 
      Height          =   525
      Index           =   13
      Left            =   0
      Tag             =   "CONF"
      Top             =   3825
      Width           =   1005
   End
End
Attribute VB_Name = "ctlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   メニューパレット ユーザーコントロール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngHotNum  As Long  ' ホット状態のボタン
Private mlngPushNum As Long  ' プッシュ状態のボタン

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event Click(tag As String)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ウィンドウの座標を取得
'
'   備考: なし
'
Public Sub WindowRect(ByRef X1 As Long, ByRef Y1 As Long, ByRef X2 As Long, ByRef Y2 As Long)
    Call GetWindowRect(UserControl.hwnd, X1, Y1, X2, Y2)
End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: イメージのマウスダウンイベント
'
'   備考: なし
'
Private Sub imgCmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    tmrMouse.Enabled = False
    mlngPushNum = Index
    imgCmd(Index).Picture = LoadResPicture(101 + (Index * 3) + 2, vbResBitmap)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: イメージのマウスムーブイベント
'
'   備考: なし
'
Private Sub imgCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    mlngHotNum = Index
    Call SetHot
    tmrMouse.Enabled = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: イメージのマウスアップイベント
'
'   備考: なし
'
Private Sub imgCmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    If mlngPushNum < 0 Then Exit Sub 'ダブルクリック対処用
    RaiseEvent Click(imgCmd(mlngPushNum).tag)
    imgCmd(mlngPushNum).Picture = LoadResPicture(101 + (Index * 3) + 0, vbResBitmap)
    mlngPushNum = -1
    tmrMouse.Enabled = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: マウスタイマーイベント
'
'   備考: なし
'
Private Sub tmrMouse_Timer()
On Error GoTo ErrorHandler
    Call checkMouseOut
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: ユーザコントロールの初期化
'
'   備考: なし
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    mlngHotNum = -1
    mlngPushNum = -1
    Call SetHot
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: マウス座標の検査
'
'   備考: なし
'
Private Sub checkMouseOut()
On Error GoTo ErrorHandler
    Dim WX1 As Long
    Dim WY1 As Long
    Dim WX2 As Long
    Dim WY2 As Long
    
    Dim BX1 As Long
    Dim BY1 As Long
    Dim BX2 As Long
    Dim BY2 As Long
    
    Dim MX As Long
    Dim MY As Long
    
    If mlngHotNum <= imgCmd.UBound Then
        If mlngHotNum >= "7" And mlngHotNum <= "12" Then
            Call GetWindowRect(fraA2.hwnd, WX1, WY1, WX2, WY2)
        Else
            Call GetWindowRect(UserControl.hwnd, WX1, WY1, WX2, WY2)
        End If
        
        ' ボタンの座標範囲を取得
        With imgCmd(mlngHotNum)
            BX1 = WX1 + .Left / Screen.TwipsPerPixelX
            BY1 = WY1 + .Top / Screen.TwipsPerPixelY
            BX2 = WX1 + (.Left + .width) / Screen.TwipsPerPixelX
            BY2 = WY1 + (.Top + .Height) / Screen.TwipsPerPixelY
        End With
        
        ' マウスカーソル位置を取得
        Call GetCursorPos(MX, MY)
        
        ' カーソルがボタンの内側か外側かを判定
        If MX >= BX1 And MX <= BX2 And _
            MY >= BY1 And MY <= BY2 Then
        Else
            mlngHotNum = imgCmd.UBound + 1
            Call SetHot
            tmrMouse.Enabled = True
        End If
    Else
        tmrMouse.Enabled = False
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub

'
'   機能: イメージのピクチャを設定
'
'   備考: なし
'
Private Sub SetHot()
    Dim i           As Long
    Dim ButtonType  As Long
    
    
    For i = 0 To imgCmd.UBound
        If i = mlngPushNum Then
            ButtonType = 2
        ElseIf i = mlngHotNum Then
            ButtonType = 1
        Else
            ButtonType = 0
        End If
        imgCmd(i).Picture = LoadResPicture(101 + (i * 3) + ButtonType, vbResBitmap)
        If i = 14 Then Exit For ' no resource
    Next i
End Sub

'
'   機能: ユーザコントロールのリサイズイベント
'
'   備考: なし
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    With UserControl
        .width = imgCmd(14).Left + imgCmd(14).width
        .Height = imgCmd(14).Top + imgCmd(14).Height
    End With
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
