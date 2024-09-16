VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVTK 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ScaleHeight     =   6630
   ScaleWidth      =   9390
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  'なし
      Caption         =   "fraHeader"
      Height          =   585
      Left            =   240
      TabIndex        =   4
      Top             =   570
      Width           =   8655
      Begin VB.Timer tmrKako 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8130
         Top             =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   10
         Top             =   30
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   1500
         TabIndex        =   9
         Top             =   30
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   1500
         TabIndex        =   8
         Top             =   210
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   1500
         TabIndex        =   7
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   3570
         TabIndex        =   6
         Top             =   210
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8955
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   8370
         TabIndex        =   20
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "リソース不足です。不要な画面を閉じてください"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   90
         Width           =   5010
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   1170
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "基本情報"
      TabPicture(0)   =   "ctlVTK.ctx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "血統"
      TabPicture(1)   =   "ctlVTK.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "paneTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "過去走詳細"
      TabPicture(2)   =   "ctlVTK.ctx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "paneTab(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "条件別"
      TabPicture(3)   =   "ctlVTK.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "持ちタイム"
      TabPicture(4)   =   "ctlVTK.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "paneTab(4)"
      Tab(4).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2325
         Index           =   0
         Left            =   -74940
         TabIndex        =   11
         Top             =   360
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   4101
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1185
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2090
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2325
         Index           =   1
         Left            =   -74940
         TabIndex        =   12
         Top             =   360
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   4101
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1185
            Index           =   1
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2090
         End
         Begin VB.PictureBox Picture2 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   23
            Top             =   0
            Width           =   0
         End
         Begin VB.PictureBox Picture1 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   22
            Top             =   0
            Width           =   0
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2325
         Index           =   2
         Left            =   60
         TabIndex        =   13
         Top             =   360
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   4101
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1185
            Index           =   2
            Left            =   0
            TabIndex        =   25
            Top             =   270
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2090
         End
         Begin MSComCtl2.UpDown updKako 
            Height          =   270
            Left            =   795
            TabIndex        =   17
            Top             =   0
            Width           =   150
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtKako"
            BuddyDispid     =   196619
            OrigLeft        =   705
            OrigRight       =   885
            OrigBottom      =   270
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtKako 
            Alignment       =   1  '右揃え
            Height          =   270
            Left            =   360
            TabIndex        =   16
            Text            =   "5"
            Top             =   0
            Width           =   435
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "走"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   19
            Top             =   45
            Width           =   180
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "過去"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   45
            Width           =   360
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2325
         Index           =   3
         Left            =   -74940
         TabIndex        =   14
         Top             =   360
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   4101
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1185
            Index           =   3
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2090
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2325
         Index           =   4
         Left            =   -74940
         TabIndex        =   15
         Top             =   360
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   4101
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1185
            Index           =   4
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2090
         End
      End
   End
   Begin MSComctlLib.ImageList ilsTbrSmall 
      Left            =   2940
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lblMakeDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0EEEE&
      Caption         =   "データ作成日: 9999年99月99日"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4170
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVTK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   特別登録馬 表示コントロール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)    '' Vierer変更イベント
Public Event NewWindow(strViewerName As String, strKey As String)   '' Vierer変更イベント
Public Event WindowTitle(strKey As String)                          '' ウインドウタイトル変更イベント
Public Event LinkContextMenu(strViewerName As String, strKey As String)
Public Event Reload()                                               '' 再読み込み
Public Event Progression()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private WithEvents mToolBar As ToolBar
Attribute mToolBar.VB_VarHelpID = -1
Private mVB As clsViewerBase
Private mViewerState As clsVSTabOnly

Private mstrTitle As String              '' ウインドウタイトル
Private mKey As clsKeyRA                 '' キー
Private WithEvents mData As clsDataTK    '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1
Private mGridData(0 To 4) As clsGridData '' グリッドデータ
Private mblnNoData As Boolean           '' データ無しフラグ

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Const MINIMUMWIDTH  As Long = 7000
Const MINIMUMHEIGHT As Long = 5000


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キー設定プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Let Key(strKey As String)
    mKey.str = strKey
    Call Update
End Property


'
'   機能: タイトル取得プロパティ
'
'   備考: Viewer必須プロパティ、　Browser が参照
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   機能: タイトル設定プロパティ
'
'   備考: ブラウザに変更通知のイベント発生
'
Public Property Let Title(strTitle As String)
    mstrTitle = strTitle
    RaiseEvent WindowTitle(mstrTitle)
End Property


'
'   機能: Viewer状態提供
'
'   備考: なし
'
Public Property Get ViewerState() As clsVSTabOnly
    Set ViewerState = mViewerState
End Property


'
'   機能: Viewer状態受け取り
'
'   備考: なし
'
Public Property Set ViewerState(RHS As clsVSTabOnly)
    Set mViewerState = RHS
End Property


'
'   機能: データ無しをブラウザに伝える
'
'   備考: Viewer必須プロパティ
'
Public Property Get NoData() As Boolean
    NoData = mblnNoData
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クリックイベント
'
'   備考: リンク先画面へ切り替える
'
Private Sub flexTab_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim msrow As Long             '' マウスロウ
    Dim mscol As Long             '' マウスカラム
    Dim item As clsGridItem     '' グリッドアイテム
    
    ' マウス位置のグリッド座標を取得
    With flexTab(Index).Grid
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    
    ' グリッドアイテムをセルから取り出す
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    ' アイテムがリンクを持っている場合
    If item.HasAKey Then
        ' 画面切り替えイベント送信
        RaiseEvent ChangeTo(item.Link, item.Key)
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 右クリックで、コンテキストメニューを出す
'
'   備考: なし
'
Private Sub flexTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    Dim item As clsGridItem
    
    ' マウスの示すグリッド座標を取得
    With flexTab(Index).Grid
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
        
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    ' データがリンクキーを持っている
    If item.HasAKey Then
        ' かつ、右クリックである
        If Button = vbRightButton Then
            RaiseEvent LinkContextMenu(item.Link, item.Key)
        End If
    End If
    Exit Sub
    
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: マウスの下がリンク可能なグリッドならば反応する為のイベント
'
'   備考: 標準的な反応は、clsGridData.MouseMoveDrivenプロシージャに任せる
'
Private Sub flexTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Select Case Index
    Case 1
        Call flexTab(Index).ReflexiveMouseMoveDriven
    Case 4
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    Case Else
        Call flexTab(Index).MouseMoveDriven
    End Select
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: タブクリックイベント
'
'   備考: なし
'
Private Sub mstTab_Click(PreviousTab As Integer)
On Error GoTo ErrorHandler
    Dim i As Integer
    
    ' 選択されたタブに対応するfraTabのみを可視化
    For i = 0 To paneTab.count - 1
        paneTab(i).Visible = (i = mstTab.Tab)
    Next i
    
    mViewerState.LastTabNumber = mstTab.Tab

    Call UserControl_Resize
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 過去走タイマー
'
'   備考: なし
'
Private Sub tmrKako_Timer()
On Error GoTo ErrorHandler
    Call mData.CancelKakoFetching
    If Not mData.NowKakoFetching Then
        tmrKako.Enabled = False
        mData.FetchKako
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 過去走テキスト変更イベント
'
'   備考: なし
'
Private Sub txtKako_Change()
On Error GoTo ErrorHandler
    If (txtKako.Text) = 0 Then
        txtKako.Enabled = False
        txtKako.Text = "5"
        txtKako.Enabled = True
    End If
    If Len(txtKako.Text) > 2 Then
        txtKako.Enabled = False
        txtKako.Text = Right$(txtKako.Text, 2)
        txtKako.Enabled = True
    End If
    If txtKako.Enabled Then
        ' レジストリに記憶
        gApp.R_KakoNum = val(txtKako.Text)
        mstTab.TabCaption(2) = "過去" & gApp.R_KakoNum & "走"
        Call mData.CancelKakoFetching
        tmrKako.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 過去走テキストキー入力イベント
'
'   備考: なし
'
Private Sub txtKako_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
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
    Dim i As Long
    Set mKey = New clsKeyRA
    Set mData = New clsDataTK
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSTabOnly
    
    mstrTitle = "出馬表"
    
    mstTab.TabCaption(2) = "過去" & gApp.R_KakoNum & "走"
    With txtKako
        .Enabled = False ' イベントを発生させない
        .Text = gApp.R_KakoNum
        .Enabled = True
    End With
    
    ' 最小幅設定
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' 共通UI設定
    Call mVB.InitGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    ' FlexGrid設定
    For i = flexTab.LBound To flexTab.UBound
        Call mVB.FlexGridCommonSetting(flexTab(i).Grid)
    Next i
    
    ' アイコンイメージをリソースファイルから取得する
    With ilsTbrSmall
        .ListImages.Clear
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add 1, , LoadResPicture(201, vbResIcon)
    End With
    
    ' Color Assign
    BackColor = gApp.ColBG
    
    lblInfo(0).BackColor = gApp.ColDarkBG
    lblInfo(0).ForeColor = Contrast(gApp.ColDarkBG)
    ' skip lblInfo(1)
    For i = 2 To lblInfo.UBound
        lblInfo(i).BackColor = gApp.ColBG
        lblInfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
    
    
        ' すべてのペインを、データ取得中に設定する。
    For i = paneTab.LBound To paneTab.UBound
        paneTab(i).Mode = 0
    Next i
    
    ' すべてのタブを無効状態にする
    For i = 0 To mstTab.Tabs - 1
        mstTab.TabEnabled(i) = False
    Next i
    
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
    Dim i As Integer
    
    ' 最小幅設定
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' 共通UIリサイズ
    
    Call mVB.ResizeGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    ' Viewer特有UIリサイズ
    For i = 0 To 4
        With paneTab(i)
            .Top = mstTab.TabHeight + 60
            .Left = 60
            .width = Bigger(1, mstTab.width - .Left * 2)
            .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
        End With ' fraTab(mstTab.Index)
        
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
    
    Next i

    With lblInfo(1)
        .Left = fraTop.width - .width - 100
    End With
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: グリッドの幅・高さを調整する
'
'   備考: なし
'
Private Sub FitGrid(gd As MSFlexGrid)
    Dim i As Long
    Dim w As Long
    Dim h As Long
    
    For i = 0 To gd.Cols - 1
        w = w + gd.ColWidth(i)
    Next i
    For i = 0 To gd.Rows - 1
        h = h + gd.RowHeight(i)
    Next i
    
    gd.width = w + gd.GridLineWidth * (gd.Cols + 1)
    gd.Height = h + gd.GridLineWidth * (gd.Rows + 1)
End Sub


'
'   機能: ユーザコントロール終了イベント
'
'   備考: なし
'
Private Sub UserControl_Terminate()
On Error GoTo ErrorHandler
    gApp.TermLog Me
        
    Set mKey = Nothing
    Set mData = Nothing
    Set mVB = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データを取得する
'
'   備考: なし
'
Private Sub Update()
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    

    ' データを取得
    gApp.Log "TK Fetch Start"
    mblnNoData = Not mData.Fetch(mKey)
    gApp.Log "TK Fetch End"

    ' ウインドウタイトルの変更
    Me.Title = mData.Title
    
    ' ラベルを取得
    For i = 0 To 7
        lblInfo(i).Caption = ReplaceAmpersand(mData.Labels(i))
    Next i
    lblMakeDate = mData.Labels(8)

    '履歴用文字列追加
    mstrTitle = mstrTitle & " " & mData.Labels(0)
    
    ' ラベルを整列
    lblInfo(2).Left = 0
    lblInfo(2).Top = 30
    lblInfo(3).Left = lblInfo(2).Left + lblInfo(2).width
    lblInfo(3).Top = lblInfo(2).Top
    lblInfo(4).Left = lblInfo(3).Left
    lblInfo(4).Top = lblInfo(3).Top + lblInfo(3).Height
    lblInfo(5).Left = lblInfo(4).Left
    lblInfo(5).Top = lblInfo(4).Top + lblInfo(4).Height
    lblInfo(6).Left = lblInfo(4).Left + lblInfo(4).width
    lblInfo(6).Top = lblInfo(4).Top
    lblInfo(7).Left = lblInfo(6).Left
    lblInfo(7).Top = lblInfo(6).Top + lblInfo(6).Height

    ' 最初に表示するタブを設定する
    If mViewerState.IsNoTouch Then
        ' タブを基本情報に
        mstTab.Tab = 1
        mstTab.Tab = 0
    Else
        i = (mViewerState.LastTabNumber + 2) Mod mstTab.Tabs
        j = mViewerState.LastTabNumber
        mstTab.Tab = i
        mstTab.Tab = j
    End If
    
    gApp.Log "TK Update Finish"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 基本情報タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchedKihon(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch Kihon"
    Call flexTab(0).InsertGrid(GridData)
    With flexTab(0).Grid
        .FixedCols = 0
    End With
    Call flexTab(0).AutoSize(0, flexTab(0).Grid.Cols - 1)
    paneTab(0).Mode = 2
    mstTab.TabEnabled(0) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 血統タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchedKetto(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch Ketto"
    Call flexTab(1).InsertGrid(GridData)
    Call flexTab(1).AutoSize(0, flexTab(1).Grid.Cols - 1)
    
    ' 血統グリッド
    With flexTab(1).Grid
        .FixedCols = 0
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCells = flexMergeRestrictRows
    End With
    paneTab(1).Mode = 2
    mstTab.TabEnabled(1) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 過去走タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchedKako(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch Kako"
    Call flexTab(2).InsertGrid(GridData)
    
    Call flexTab(2).AutoSize(0, flexTab(2).Grid.Cols - 1, False, True)
    ' 過去N走タブ
    With flexTab(2).Grid
        .FixedCols = 0
        .FixedRows = 1
        .WordWrap = True
        
        ' セル幅を固定に
        Dim i As Integer
        .ColWidth(0) = 1800
        For i = 1 To .Cols - 1
            If LenB(.TextMatrix(1, 1)) > 10 Then
                .ColWidth(i) = 2880
            Else
                .ColWidth(i) = 700
            End If
        Next
    End With
    paneTab(2).Mode = 2
    mstTab.TabEnabled(2) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 条件別タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchedJokenBetu(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch JokenBetu"
    Call flexTab(3).InsertGrid(GridData)
    ' 条件別グリッド
    With flexTab(3).Grid
        Call flexTab(3).AutoSize(0, .Cols - 1, False, False, 1)
        
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeRestrictRows
        .FixedCols = 0
        .FixedRows = 2
    End With
    paneTab(3).Mode = 2
    mstTab.TabEnabled(3) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 持ちタイムタブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchedMotiTIme(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch MochiTime"
    
    Call flexTab(4).InsertGrid(GridData)
    Call flexTab(4).AutoSize(0, flexTab(4).Grid.Cols - 1, False, False, 1)
    ' 持ちタイム
    With flexTab(4).Grid
        
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .FixedRows = 2
        .FixedCols = 0
    End With
    paneTab(4).Mode = 2
    mstTab.TabEnabled(4) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データがない
'
'   備考: なし
'
Private Sub mData_NoUMARACE()
On Error GoTo ErrorHandler
    Dim i As Long
    
    For i = 0 To 4
        paneTab(i).Mode = 1
        mstTab.TabEnabled(i) = True
    Next i
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
    gApp.Log "Free"
    If Not mData Is Nothing Then
        mData.CancelKakoFetching
        mData.CancelFetching
    End If
    Dim i As Integer
    For i = flexTab.LBound To flexTab.UBound
        Call DestroyFlexGrid(flexTab(i))
    Next i
    Set mData = Nothing
End Sub

