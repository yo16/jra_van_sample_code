VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVRC 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ScaleHeight     =   6090
   ScaleWidth      =   7995
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  'なし
      Caption         =   "fraHeader"
      Height          =   825
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   6105
      Begin VB.PictureBox picXPTheme 
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         Height          =   645
         Left            =   240
         ScaleHeight     =   645
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   90
         Width           =   495
         Begin VB.CommandButton cmdNextRC 
            Caption         =   "次"
            Height          =   255
            Left            =   0
            Picture         =   "ctlVRC.ctx":0000
            TabIndex        =   12
            Top             =   0
            Width           =   420
         End
         Begin VB.CommandButton cmdPrevRC 
            Caption         =   "前"
            Height          =   255
            Left            =   0
            Picture         =   "ctlVRC.ctx":058A
            TabIndex        =   11
            Top             =   360
            Width           =   420
         End
      End
      Begin Umakichi.ctlClickLabel clblinfo 
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   6
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         Caption         =   "Next"
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "ctlVRC.ctx":0B14
         Top             =   270
         Width           =   4575
      End
      Begin Umakichi.ctlClickLabel clblinfo 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         Caption         =   "Prev"
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
      Width           =   7605
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
         Top             =   120
         Width           =   5010
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2745
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4842
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "レコード"
      TabPicture(0)   =   "ctlVRC.ctx":0B18
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   9
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
      Left            =   5400
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   レコード 表示コントロール
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
Private mViewerState    As clsVSNothing '' Viewer State

Private WithEvents mData As clsDataRC   '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' ウインドウタイトル
Private mKey As clsKeyRC                '' キー
Private mblnNoData As Boolean           '' データ無しフラグ

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

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
'   機能: 次ボタンクリック押下イベント
'
'   備考: なし
'
Private Sub cmdNextRC_Click()
On Error GoTo ErrorHandler
    Call clblinfo_ChangeViewer(0)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 前ボタンクリック押下イベント
'
'   備考: なし
'
Private Sub cmdPrevRC_Click()
On Error GoTo ErrorHandler
    Call clblinfo_ChangeViewer(1)
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
    
    Call flexTab(Index).MouseMoveDriven
    
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
    
    Call UserControl_Resize
    
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
    Set mKey = New clsKeyRC
    Set mData = New clsDataRC
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSNothing
    
    mstrTitle = "レコード"
    
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
    picXPTheme.BackColor = gApp.ColBG
    cmdNextRC.BackColor = gApp.ColBG
    cmdPrevRC.BackColor = gApp.ColBG
    
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
    
    
    With paneTab.item(mstTab.Tab)
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' paneTab.Item(mstTab.Index)
    
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
    Next i
    
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
    
    Set mVB = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: レコードタブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchComplete(gd As clsGridData, col As Long)
On Error GoTo ErrorHandler
    Call flexTab(0).InsertGrid(gd)
    With flexTab(0).Grid
        .FixedRows = 0
        Call flexTab(0).AutoSize(0, .Cols - 1, True)
        .CellAlignment = vbAlignNone
        .MergeCol(0) = True
        .MergeCells = flexMergeFree
    End With
    paneTab(0).Mode = 2
    mstTab.TabEnabled(0) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: レコードタブのデータがない
'
'   備考: なし
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    paneTab(0).Mode = 1
    mstTab.TabEnabled(0) = True
    gApp.Log "d:該当レコードはありません。存在するRCのkeyを指定してください"
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
    
    ' データを取得してもらう
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If

    ' ラベルをもらう
    lblMakeDate = mData.Labels(0)
    lblInfo(0) = ReplaceAmpersand(mData.Labels(1))
    txtInfo(0) = mData.Labels(2)
    
    '履歴用文字列追加
    mstrTitle = mstrTitle & " " & mData.Labels(1)
    
    ' データの有無でコマンドの可不可を設定
    If mData.LinkLabels(0).Text = "" Then
        cmdNextRC.Enabled = False
        clblinfo(0).Visible = False
    Else
        Set clblinfo(0).LinkItem = mData.LinkLabels(0)
    End If
    If mData.LinkLabels(1).Text = "" Then
        cmdPrevRC.Enabled = False
        clblinfo(1).Visible = False
    Else
        Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    End If
    
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
    
    Dim i As Integer
    For i = flexTab.LBound To flexTab.UBound
        Call DestroyFlexGrid(flexTab(i))
    Next i
    Set mData = Nothing
End Sub

