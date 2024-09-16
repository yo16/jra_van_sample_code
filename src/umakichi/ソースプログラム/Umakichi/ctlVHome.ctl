VERSION 5.00
Begin VB.UserControl ctlVHome 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   ScaleHeight     =   6480
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   8340
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   0
      Left            =   570
      TabIndex        =   6
      ToolTipText     =   "出馬表を選択、表示します。"
      Top             =   1500
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "Empty"
      ViewerName      =   "RAKaiSel"
      Caption         =   "出馬表"
   End
   Begin VB.Frame fraSub 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   345
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   2640
      Width           =   6915
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "マスター情報"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   60
         Width           =   1245
      End
   End
   Begin VB.Frame fraSub 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   345
      Index           =   0
      Left            =   330
      TabIndex        =   2
      Top             =   990
      Width           =   6195
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "レース情報"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   60
         Width           =   1065
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7800
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "総合メニュー"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   60
         Width           =   1950
      End
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   1
      Left            =   570
      TabIndex        =   7
      Top             =   1860
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "Empty"
      ViewerName      =   "TKKaiSel"
      Caption         =   "特別登録馬"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   2
      Left            =   2250
      TabIndex        =   8
      Top             =   1860
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "Empty"
      ViewerName      =   "HCSel"
      Caption         =   "坂路調教"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   3
      Left            =   570
      TabIndex        =   9
      Top             =   3150
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      ViewerName      =   "Find"
      Caption         =   "競走馬"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   4
      Left            =   570
      TabIndex        =   10
      Top             =   3510
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "1"
      ViewerName      =   "Find"
      Caption         =   "騎手"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   5
      Left            =   570
      TabIndex        =   11
      Top             =   3870
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "2"
      ViewerName      =   "Find"
      Caption         =   "調教師"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   6
      Left            =   2250
      TabIndex        =   12
      Top             =   3150
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "3"
      ViewerName      =   "Find"
      Caption         =   "馬主"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   7
      Left            =   2250
      TabIndex        =   13
      Top             =   3510
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "4"
      ViewerName      =   "Find"
      Caption         =   "生産者"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   8
      Left            =   2250
      TabIndex        =   14
      Top             =   3870
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "5"
      ViewerName      =   "Find"
      Caption         =   "繁殖馬"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   9
      Left            =   2280
      TabIndex        =   15
      Top             =   1530
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      ViewerName      =   "RCSel"
      Caption         =   "コースレコード"
   End
   Begin VB.Menu mnuRight 
      Caption         =   "右クリック"
      Visible         =   0   'False
      Begin VB.Menu mnuNewWin 
         Caption         =   "新しいウインドウで開く"
         Index           =   0
      End
   End
End
Attribute VB_Name = "ctlVHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   ホーム画面 ホームボタンを押すと、この画面が表示される
'   デフォルトの起動時画面もこの画面である
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)        '' Vierer変更イベント
Public Event LinkContextMenu(strViewerName As String, strKey As String) '' リンク右クリックイベント

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mViewerState    As clsVSNothing         '' 状態
Private mblnNoData As Boolean           '' データ無しフラグ

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キー設定プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Let Key(strKey As String)
    ' Empty
End Property


'
'   機能: キー取得プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Get Key() As String
    ' Empty
End Property


'
'   機能: タイトル取得プロパティ
'
'   備考: Viewer必須プロパティ、　Browser が呼びます
'
Public Property Get Title() As String
    Title = "ホーム"
End Property


'
'   機能: Viewer状態提供
'
'   備考: なし
'
Public Property Get ViewerState() As clsVSNothing
    Set ViewerState = mViewerState
End Property


'
'   機能: Viewer状態受け取り
'
'   備考: なし
'
Public Property Set ViewerState(RHS As clsVSNothing)
    Set mViewerState = RHS
End Property


'
'   機能: データ無しをブラウザに伝える
'
'   備考:　Viewer必須プロパティ
'
Public Property Get NoData() As Boolean
    NoData = mblnNoData
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: 画面変更イベント
'
'   備考: ブラウザにイベントをスルーする
'
Private Sub clblCmd_ChangeViewer(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent ChangeTo(clblCmd(Index).ViewerName, clblCmd(Index).Key)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 右クリック押下イベント
'
'   備考: ブラウザのポップアップイベントにスルーする
'
Private Sub clblCmd_RightMouseDown(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent LinkContextMenu(clblCmd(Index).ViewerName, clblCmd(Index).Key)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロール初期化
'
'   備考: なし
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    gApp.InitLog Me
    
    Dim i As Integer
    
    Set mViewerState = New clsVSNothing

    ' Color Assign
    BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
    For i = 0 To fraSub.UBound
        fraSub(i).BackColor = gApp.ColDarkBG
    Next i
    
    For i = lblFix.LBound To lblFix.UBound
        lblFix(i).BackColor = gApp.ColDarkBG
        lblFix(i).ForeColor = Contrast(gApp.ColDarkBG)
    Next i
    
    For i = clblCmd.LBound To clblCmd.UBound
        clblCmd(i).BackColor = gApp.ColBG
    Next i
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロールのマウスアップイベント
'
'   備考: なし
'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Call UserControl.Parent.ShowPopupMenu(Button, Shift, X, Y)
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
    With UserControl
        .width = Bigger(8000, .width)
        .Height = Bigger(5000, .Height)
    End With
    
    fraTop.width = ScaleWidth - fraTop.Left * 2
    fraSub(0).width = ScaleWidth - fraSub(0).Left * 2
    fraSub(1).width = ScaleWidth - fraSub(1).Left * 2
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロール終了イベント
'
'   備考: なし
'
Private Sub UserControl_Terminate()
On Error GoTo ErrorHandler
    gApp.TermLog Me
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 終了処理
'
'   備考: なし
'
Public Sub Free()
End Sub

