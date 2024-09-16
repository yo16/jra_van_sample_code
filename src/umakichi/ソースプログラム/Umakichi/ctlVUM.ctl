VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVUM 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ScaleHeight     =   6885
   ScaleWidth      =   11280
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  'なし
      Caption         =   "fraHeader"
      Height          =   1665
      Left            =   330
      TabIndex        =   4
      Top             =   600
      Width           =   10185
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   2
         Left            =   3690
         TabIndex        =   11
         Top             =   675
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   318
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "."
      End
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   1
         Left            =   3690
         TabIndex        =   10
         Top             =   480
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   318
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "."
      End
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   0
         Left            =   3690
         TabIndex        =   12
         Top             =   270
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   318
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "."
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "ctlVUM.ctx":0000
         Top             =   300
         Width           =   825
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   5880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "ctlVUM.ctx":001F
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "ctlVUM.ctx":0054
         Top             =   900
         Width           =   2685
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "ctlVUM.ctx":0088
         Top             =   90
         Width           =   5265
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "ctlVUM.ctx":00EB
         Top             =   900
         Width           =   2685
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   5010
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   3615
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "血統"
      TabPicture(0)   =   "ctlVUM.ctx":011F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "過去成績"
      TabPicture(1)   =   "ctlVUM.ctx":013B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "paneTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "条件別成績"
      TabPicture(2)   =   "ctlVUM.ctx":0157
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vsbSE"
      Tab(2).Control(1)=   "hsbSE"
      Tab(2).Control(2)=   "paneTab(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "持ちタイム"
      TabPicture(3)   =   "ctlVUM.ctx":0173
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "坂路調教"
      TabPicture(4)   =   "ctlVUM.ctx":018F
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "paneTab(4)"
      Tab(4).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2205
         Index           =   2
         Left            =   -74940
         TabIndex        =   15
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3889
         Begin VB.PictureBox picIPane 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H8000000C&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   0
            ScaleHeight     =   1785
            ScaleWidth      =   6945
            TabIndex        =   16
            Top             =   0
            Width           =   6975
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1455
               Index           =   2
               Left            =   0
               TabIndex        =   21
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   2566
            End
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1455
               Index           =   5
               Left            =   1920
               TabIndex        =   26
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   2566
            End
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1455
               Index           =   6
               Left            =   3960
               TabIndex        =   27
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   2566
            End
         End
      End
      Begin VB.HScrollBar hsbSE 
         Height          =   285
         Left            =   -74340
         TabIndex        =   14
         Top             =   3060
         Width           =   5295
      End
      Begin VB.VScrollBar vsbSE 
         Height          =   1995
         Left            =   -67440
         TabIndex        =   13
         Top             =   360
         Width           =   285
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2370
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   4180
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2160
         Index           =   1
         Left            =   -75000
         TabIndex        =   19
         Top             =   840
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   3810
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   1
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '透明
            Caption         =   "障害レースについては、[後3ハロン]に""後3Fタイム""でなく、""当該レース走破タイムの1F平均タイム""を表示しています。"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   0
            TabIndex        =   28
            Top             =   1560
            Width           =   8190
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   3
         Left            =   -75000
         TabIndex        =   22
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   3
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   4
         Left            =   -75000
         TabIndex        =   24
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   4
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
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
      Left            =   6060
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   競走馬  表示コントロール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)        '' Vierer変更イベント
Public Event WindowTitle(strKey As String)                              '' ウインドウタイトル変更イベント
Public Event LinkContextMenu(strViewerName As String, strKey As String) '' 右クリックメニュー表示イベント

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB As clsViewerBase            '' Viewer Base
Private mViewerState As clsVSTabOnly    '' Viewer State

Private WithEvents mData As clsDataUM   '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' ウインドウタイトル
Private mKey As clsKeyUM                '' キー
Private mblnNoData As Boolean           '' データ無しフラグ

Private mSortAscending As Boolean       '' ソート方向フラグ

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 画面最小幅値
Const MINIMUMWIDTH  As Long = 4000
Const MINIMUMHEIGHT As Long = 4000

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キー設定プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Let Key(strKey As String)
    gApp.Log "UM: " & strKey
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
'   備考:　Viewer必須プロパティ
'
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
Private Sub clblinfo_ChangeViewer(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent ChangeTo(clblinfo(Index).ViewerName, clblinfo(Index).Key)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 右クリック押下イベント
'
'   備考: ブラウザのポップアップイベントにスルーする
'
Private Sub clblinfo_RightMouseDown(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent LinkContextMenu(clblinfo(Index).ViewerName, clblinfo(Index).Key)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


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
    msrow = flexTab(Index).Grid.MouseRow
    mscol = flexTab(Index).Grid.MouseCol
    
    ' グリッドアイテムをセルから取り出す
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
    
    If 0 = Index Then
        Call flexTab(Index).MouseMoveDriven
    ElseIf 1 = Index Or 3 = Index Then
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    Else
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ソート前イベント。特殊ソートを制御します｡
'
'   備考: グリッド特定カラムのsort禁止 & 隠しカラムでソート。
'
Private Sub flexTab_BeforeSort(Index As Integer, ByVal col As Long, Order As Integer)
On Error GoTo ErrorHandler
    If Index = 1 And col = 0 Then
        Order = 0
        With flexTab(Index).Grid
            If mSortAscending Then
                mSortAscending = False
            Else
                mSortAscending = True
            End If
        End With
    End If
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
    
    Call Tab_Resize
    
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

    Dim i As Long
    Set mKey = New clsKeyUM
    Set mData = New clsDataUM
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSTabOnly
    
    mstrTitle = "競走馬"

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
    
    ' スクロールバー初期化
    vsbSE.width = gApp.vsbWidth
    hsbSE.Height = gApp.hsbHeight
    
    ' Color Assign
    BackColor = gApp.ColBG
    
    lblInfo(0).BackColor = gApp.ColDarkBG
    lblInfo(0).ForeColor = Contrast(gApp.ColDarkBG)
    For i = txtInfo.LBound To txtInfo.UBound
        txtInfo(i).BackColor = gApp.ColBG
        txtInfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
    For i = clblinfo.LBound To clblinfo.UBound
        clblinfo(i).BackColor = gApp.ColBG
        clblinfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
        
    Call mstTab_Click(0)
    
    ' すべてのペインを、データ取得中に設定する。
    For i = paneTab.LBound To paneTab.UBound
        paneTab(i).Mode = 0
    Next i
    
    ' すべてのタブを、データ取得中に設定する。
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
    
    Call Tab_Resize
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
    

'
'   機能: タブのリサイズイベント
'
'   備考: なし
'
Private Sub Tab_Resize()
On Error GoTo ErrorHandler
    Dim i As Integer
    
    For i = 0 To 4
        With paneTab(i)
            .Top = mstTab.TabHeight + 60
            .Left = 60
            .width = Bigger(1, mstTab.width - .Left * 2)
            .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
        End With ' fraTab(mstTab.Index)

        Select Case i
        Case 0, 3, 4
        ' 過去成績, 条件別成績タブ意外は、グリッドを最大に
            With flexTab(i)
                .width = Bigger(1, paneTab(i).width - .Left)
                .Height = Bigger(1, paneTab(i).Height - .Top)
            End With ' flexTab(i)
        Case 1
        ' 過去成績タブ
            With flexTab(i)
                .Top = lblFix.Height
                .width = Bigger(1, paneTab(i).width - .Left)
                .Height = Bigger(1, paneTab(i).Height - .Top)
                lblFix.Top = 0
                lblFix.Left = 0
            End With
        Case 2
        ' 条件別成績タブ
            Call ScrollBarManage

            flexTab(2).Height = 2100
            flexTab(2).width = 6650
            flexTab(5).Height = 3400
            flexTab(5).width = 6650
            flexTab(6).Height = 2100
            flexTab(6).width = 6650
            
            flexTab(2).Top = 0
            flexTab(2).Left = 0
            flexTab(5).Top = 0
            flexTab(5).Left = flexTab(2).width + 300
            flexTab(6).Top = flexTab(2).Height + 300
            flexTab(6).Left = 0

            With picIPane
                .width = Bigger(MINIMUMWIDTH, flexTab(5).Left + flexTab(5).width + 200)
                .Height = Bigger(MINIMUMHEIGHT, flexTab(6).Top + flexTab(6).Height)
                .width = Bigger(.width, paneTab(2).width - .Left)
                .Height = Bigger(.Height, paneTab(2).Height - .Top)
            End With
        End Select
    Next i
 
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: スクロールバー
'
'   備考: なし
'
Private Sub ScrollBarManage()
    Dim hsbIsVisible As Boolean

    vsbSE.Visible = False
    hsbSE.Visible = False

    ' 水平
    hsbIsVisible = False
    If picIPane.width > paneTab(2).width + vsbSE.width Then
        paneTab(2).Height = paneTab(2).Height - hsbSE.Height
        hsbIsVisible = True
        hsbSE.Visible = (2 = mstTab.Tab)
    End If

    ' 垂直
    If picIPane.Height > paneTab(2).Height + hsbSE.Height Then
        paneTab(2).width = paneTab(2).width - vsbSE.width
        vsbSE.Visible = (2 = mstTab.Tab)
    End If

    ' 垂直スクロールバーにより水平スクロールバーが必要になったとき
    If hsbIsVisible = False And picIPane.width > paneTab(2).width + vsbSE.width Then
        paneTab(2).Height = paneTab(2).Height - hsbSE.Height
        hsbSE.Visible = (2 = mstTab.Tab)
    End If

    With hsbSE
        .Top = paneTab(2).Top + paneTab(2).Height
        .Left = paneTab(2).Left
        .width = paneTab(2).width
    End With

    With vsbSE
        .Top = paneTab(2).Top
        .Left = paneTab(2).Left + paneTab(2).width
        .Height = paneTab(2).Height
    End With

    hsbSE.max = picIPane.width - paneTab(2).width
    hsbSE.LargeChange = paneTab(2).width
    hsbSE.SmallChange = vsbSE.width

    vsbSE.max = picIPane.Height - paneTab(2).Height
    vsbSE.LargeChange = paneTab(2).Height
    vsbSE.SmallChange = hsbSE.Height
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
'   機能: データがない
'
'   備考: なし
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    gApp.Log "d:該当レコードはありません。UMの存在するkeyを指定してください。" & vbCrLf _
            & "呼び出し元をチェックしましょう＞＞開発者"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: タブデータがない
'
'   備考: なし
'
Private Sub mData_NoTabData(Index As Long)
On Error GoTo ErrorHandler
    paneTab(Index).Mode = 1
    mstTab.TabEnabled(Index) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: リンクラベルデータの取得完了
'
'   備考: なし
'
Private Sub mData_FetchedLinkLabels()
On Error GoTo ErrorHandler
    Dim i As Long

    gApp.Log "FetchedLinkLabels"
    Set clblinfo(0).LinkItem = mData.LinkLabels(0)
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    Set clblinfo(2).LinkItem = mData.LinkLabels(2)
    
    For i = 0 To 2
        With clblinfo(i)
            If .Key <> "" Then
                .ForeColor = ColorLinkExist
                .Font.Underline = True
            Else
                .ForeColor = Contrast(gApp.ColBG)
                .Font.Underline = False
            End If
        End With
    Next i
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 条件別データ取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_FetchedJokenBetu(gd2 As clsGridData, gd5 As clsGridData, gd6 As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedJokenBetu"
    Dim i As Long
    Dim r As Long, c As Long
    i = 2
    Call flexTab(i).InsertGrid(gd2)
    With flexTab(i).Grid
        .ScrollBars = flexScrollBarNone
        .BorderStyle = flexBorderNone
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 2
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .Enabled = False
                
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
    
    i = 5
    Call flexTab(i).InsertGrid(gd5)
    With flexTab(i).Grid
        .ScrollBars = flexScrollBarNone
        .BorderStyle = flexBorderNone
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 2
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .Enabled = False
        
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
    
    i = 6
    Call flexTab(i).InsertGrid(gd6)
    With flexTab(i).Grid
        .ScrollBars = flexScrollBarNone
        .BorderStyle = flexBorderNone
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 2
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .Enabled = False
        
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
        
    Call UserControl_Resize
    
    paneTab(2).Mode = 2
    paneTab(2).BorderStyle = ebscThin
    mstTab.TabEnabled(2) = True
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 血統データ取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_FetchedKetto(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedKetto"
    Dim i As Long
    
    i = 0
    Call flexTab(i).InsertGrid(gd)
    With flexTab(i).Grid
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .ColWidth(0) = 0
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCells = flexMergeRestrictColumns
    End With
    paneTab(i).Mode = 2
    mstTab.TabEnabled(i) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 坂路調教データ取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_FetchedHanro(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedHanro"
    Dim i As Long
    
    i = 4
    Call flexTab(i).InsertGrid(gd)
    If flexTab(i).Grid.Rows > 2 Then
        With flexTab(i).Grid
            flexTab(i).Grid.TextMatrix(0, 0) = " "
            flexTab(i).Grid.TextMatrix(0, 1) = " "
            flexTab(i).Grid.TextMatrix(0, 2) = " "
            
            Call flexTab(i).AutoSize(0, .Cols - 1, False, False, 0)
            
            .FixedRows = 2
            .FixedCols = 0
            .MergeRow(0) = True
            .MergeCells = flexMergeFree
            
            Dim r As Long, c As Long
            'For r = 0 To 1
            
        End With
        paneTab(i).Mode = 2
    End If
    mstTab.TabEnabled(i) = True
        
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 過去走データ取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_FetchedKako(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedKako"
    Dim i As Long
    
    i = 1
    Call flexTab(i).InsertGrid(gd)
    With flexTab(i).Grid
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 0
        .MergeCells = flexMergeRestrictColumns
        
        .col = 0
        .Sort = flexSortStringDescending
        Call SortFlexGrid(flexTab(i), .Cols - 1)
        
        .ColWidth(.Cols - 1) = 0
    End With
    paneTab(i).Mode = 2
    mstTab.TabEnabled(i) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 持ちタイムデータ取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_FetchedTime(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedTime"
    Dim i As Long
    Dim r As Long, c As Long
    Dim strTemp As String
    
    i = 3
    Call flexTab(i).InsertGrid(gd)
    With flexTab(i).Grid
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 0
        .MergeCells = flexMergeRestrictColumns
        .row = 0
        
        .col = .Cols - 1
        .Sort = flexSortStringAscending
        Call SortFlexGrid(flexTab(i), .Cols - 1)
        
        For r = 1 To .Rows - 1
            .row = r
            .col = .Cols - 1
            If strTemp <> Left(.Text, 5) Then
                strTemp = Left(.Text, 5)
            Else
                .RowHeight(r) = 0
            End If
        Next r
        
        .ColWidth(.Cols - 2) = 0
        .ColWidth(.Cols - 1) = 0
    End With
    paneTab(i).Mode = 2
    mstTab.TabEnabled(i) = True
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
    Dim i As Integer
    Dim j As Integer
    
    ' データを取得
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If
    
    ' ラベルを取得
    lblMakeDate = mData.Labels(0)
    lblInfo(0) = ReplaceAmpersand(mData.Labels(1))
    txtInfo(0) = mData.Labels(2)
    txtInfo(1) = mData.Labels(3)
    txtInfo(2) = mData.Labels(4)
    txtInfo(3) = mData.Labels(5)
    txtInfo(4) = mData.Labels(6)
    Set clblinfo(0).LinkItem = mData.LinkLabels(0)
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    Set clblinfo(2).LinkItem = mData.LinkLabels(2)
    
    '履歴用文字列追加
    mstrTitle = mstrTitle & " " & mData.Labels(1)
    
    ' 最初に表示するタブを設定する
    If mViewerState.IsNoTouch Then
        mstTab.Tab = 0
    Else
        mstTab.Tab = mViewerState.LastTabNumber
    End If
End Sub


'
'   機能: 水平スクロールバー変更イベント
'
'   備考: なし
'
Private Sub hsbSE_Change()
On Error GoTo ErrorHandler
    picIPane.Left = -hsbSE.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 水平スクロールバースクロールイベント
'
'   備考: なし
'
Private Sub hsbSE_Scroll()
On Error GoTo ErrorHandler
    picIPane.Left = -hsbSE.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 垂直スクロールバー変更イベント
'
'   備考: なし
'
Private Sub vsbSE_Change()
On Error GoTo ErrorHandler
    picIPane.Top = -vsbSE.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 垂直スクロールバースクロールイベント
'
'   備考: なし
'
Private Sub vsbSE_Scroll()
On Error GoTo ErrorHandler
    picIPane.Top = -vsbSE.value
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
    If Not (mData Is Nothing) Then mData.CancelFetching
    Dim i As Integer
    For i = flexTab.LBound To flexTab.UBound
        Call DestroyFlexGrid(flexTab(i))
    Next i
    Set mData = Nothing
End Sub

