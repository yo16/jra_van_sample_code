VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVSK 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   ScaleHeight     =   6390
   ScaleWidth      =   9360
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  'なし
      Caption         =   "fraHeader"
      Height          =   1035
      Left            =   180
      TabIndex        =   3
      Top             =   810
      Width           =   7185
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   210
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
         Height          =   195
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   $"ctlVSK.ctx":0000
         Top             =   210
         Visible         =   0   'False
         Width           =   730
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
         Height          =   945
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "ctlVSK.ctx":000E
         Top             =   30
         Width           =   3795
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "鶴田任男"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   120
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   423
         AutoSize        =   -1  'True
         BackColor       =   12635340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "リソース不足です。不要な画面を閉じてください"
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2985
      Left            =   240
      TabIndex        =   1
      Top             =   1950
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "血統"
      TabPicture(0)   =   "ctlVSK.ctx":0046
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   8
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
      Left            =   6540
      TabIndex        =   2
      Top             =   30
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   産駒マスタ 表示コントロール
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

Private WithEvents mData As clsDataSK   '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' ウインドウタイトル
Private mKey As clsKeySK                '' キー
Private mblnNoData As Boolean           '' データ無しフラグ

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
    gApp.Log "SK: " & strKey
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
'   備考: Viewer必須プロパティ
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
    With paneTab.item(mstTab.Tab)
        .Top = mstTab.TabHeight + 60
        .Left = 60
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
'   機能: ユーザコントロール初期化
'
'   備考: なし
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    gApp.InitLog Me
    
    Dim i As Long
    Set mKey = New clsKeySK
    Set mData = New clsDataSK
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSNothing
    
    mstrTitle = "産駒"

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
    
    clblinfo(0).BackColor = gApp.ColDarkBG
    clblinfo(0).ForeColor = Contrast(gApp.ColBG)
    clblinfo(1).BackColor = gApp.ColBG
    clblinfo(1).ForeColor = Contrast(gApp.ColBG)
    txtInfo(0).BackColor = gApp.ColBG
    txtInfo(0).ForeColor = Contrast(gApp.ColBG)
    txtInfo(1).BackColor = gApp.ColBG
    txtInfo(1).ForeColor = Contrast(gApp.ColBG)
    
    ' ctlPane の初期化
    paneTab(0).Mode = 0 ' データ取得中に
    mstTab.TabEnabled(0) = False
    
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
'   機能: データがない
'
'   備考: なし
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    gApp.Log "d:該当レコードはありません。SKの存在するkeyを指定してください。" & vbCrLf _
            & "呼び出し元をチェックしましょう＞＞開発者"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: リンクラベルデータ取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_FetchCompleteLinklabels()
On Error GoTo ErrorHandler
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    With clblinfo(1)
        If .Key <> "" Then
            .ForeColor = ColorLinkExist
            .Font.Underline = True
        Else
            .ForeColor = Contrast(gApp.ColBG)
            .Font.Underline = False
        End If
    End With
    txtInfo(1).Visible = True
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: 血統タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteKetto(gd As clsGridData)
On Error GoTo ErrorHandler
    
    Call flexTab(0).InsertGrid(gd)
    
    With flexTab(0).Grid
        .FixedCols = 0
        Call flexTab(0).AutoSize(0, .Cols - 1)
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCells = flexMergeRestrictColumns
    End With
    paneTab(0).Mode = 2
    mstTab.TabEnabled(0) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能:　血統タブのデータがない
'
'   備考: なし
'
Private Sub mData_NoDataKetto()
On Error GoTo ErrorHandler
    paneTab(0).Mode = 1
    mstTab.TabEnabled(0) = True
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

    ' データを取得してもらう
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If

    ' ラベルを取得
    lblMakeDate = mData.Labels(0)
    txtInfo(0) = mData.Labels(1)
    Set clblinfo(0).LinkItem = mData.LinkLabels(0)
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    For i = 0 To 1
        Set clblinfo(i).LinkItem = mData.LinkLabels(i)
        With clblinfo(i)
            If .Key <> "" Then
                .ForeColor = ColorLinkExist
                .Font.Underline = True
            Else
                .ForeColor = Contrast(IIf(i = 0, gApp.ColDarkBG, gApp.ColBG))
                .Font.Underline = False
            End If
        End With
    Next i
    
    '履歴用文字列追加
    mstrTitle = mstrTitle & " " & clblinfo(0).Caption
    
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

