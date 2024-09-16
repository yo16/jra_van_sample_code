VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVRASel 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   ScaleHeight     =   6855
   ScaleWidth      =   9495
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "出馬表選択画面"
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
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   1800
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2745
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4842
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "開催情報"
      TabPicture(0)   =   "ctlVRASel.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   1815
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   3201
         Begin VB.PictureBox fraScroll 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H80000003&
            ForeColor       =   &H80000008&
            Height          =   1035
            Left            =   0
            ScaleHeight     =   1005
            ScaleWidth      =   2265
            TabIndex        =   6
            Top             =   0
            Width           =   2295
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   405
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   714
            End
         End
         Begin VB.VScrollBar vsbSel 
            Height          =   1395
            Left            =   4320
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.HScrollBar hsbSel 
            Height          =   255
            Left            =   1170
            TabIndex        =   4
            Top             =   1440
            Visible         =   0   'False
            Width           =   3105
         End
      End
   End
End
Attribute VB_Name = "ctlVRASel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   出馬表選択 表示コントロール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)
Public Event WindowTitle(strKey As String)
Public Event LinkContextMenu(strViewerName As String, strKey As String)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrTitle       As String
Private mKey            As clsKeyRASel
Private mVB             As clsViewerBase
Private mViewerState    As clsVSNothing         '' 状態
Private mblnNoData As Boolean           '' データ無しフラグ

Private WithEvents mData As clsDataRASel
Attribute mData.VB_VarHelpID = -1

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
'   機能: 水平スクロールバー変更イベント
'
'   備考: なし
'
Private Sub hsbSel_Change()
On Error GoTo ErrorHandler
    flexTab.Left = -hsbSel.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 水平スクロールバーフォーカス取得イベント
'
'   備考: なし
'
Private Sub hsbSel_GotFocus()
On Error GoTo ErrorHandler
    paneTab.SetFocus
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 水平スクロールバースクロールイベント
'
'   備考: なし
'
Private Sub hsbSel_Scroll()
On Error GoTo ErrorHandler
    flexTab.Left = -hsbSel.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 垂直スクロールバー変更イベント
'
'   備考: なし
'
Private Sub vsbSel_Change()
On Error GoTo ErrorHandler
    flexTab.Top = -vsbSel.value
    paneTab.SetFocus
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 垂直スクロールバーフォーカス取得イベント
'
'   備考: なし
'
Private Sub vsbSel_GotFocus()
On Error GoTo ErrorHandler
    paneTab.SetFocus
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 垂直スクロールバーフォーカス取得イベント
'
'   備考: なし
'
Private Sub vsbSel_Scroll()
On Error GoTo ErrorHandler
    flexTab.Top = -vsbSel.value
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
    
    mstrTitle = "出馬表選択"
    Set mData = New clsDataRASel
    Set mKey = New clsKeyRASel
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSNothing
    
    Call mVB.FlexGridCommonSetting(flexTab.Grid)

    ' Color Assign
    BackColor = gApp.ColBG
    mstTab.BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
    
    lblInfo.BackColor = gApp.ColDarkBG
    lblInfo.ForeColor = Contrast(gApp.ColDarkBG)
    
    flexTab.Grid.ScrollBars = flexScrollBarNone

    paneTab.Mode = 0
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


'
'   機能: マウスの下がリンク可能なグリッドならば反応する為のイベント
'
'   備考: 標準的な反応は、clsGridData.MouseMoveDrivenプロシージャに任せる
'
Private Sub flexTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    
    Call flexTab.MouseMoveDriven
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: クリックイベント
'
'   備考: リンク先画面へ切り替える
'
Private Sub flexTab_Click()
On Error GoTo ErrorHandler
    Dim msrow As Long             '' マウスロウ
    Dim mscol As Long             '' マウスカラム
    Dim item As clsGridItem     '' グリッドアイテム
    
    ' マウス位置のグリッド座標を取得
    With flexTab.Grid
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    
    ' グリッドアイテムをセルから取り出す
    Call SetItem(item, flexTab, msrow, mscol)
    
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
Private Sub flexTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    Dim item As clsGridItem
    
    ' マウスの示すグリッド座標を取得
    msrow = flexTab.Grid.MouseRow
    mscol = flexTab.Grid.MouseCol
    
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
    Call SetItem(item, flexTab, msrow, mscol)
    
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
'   機能: ユーザコントロールのリサイズイベント
'
'   備考: なし
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim blnVSBVisible As Boolean
    Dim blnHSBVisible As Boolean
    
    ' ユーザーコントロールの最低幅を決める
    With UserControl
        .width = Bigger(8000, .width)
        .Height = Bigger(5000, .Height)
    End With
    
    
    fraTop.width = ScaleWidth - fraTop.Left * 2
    With mstTab
        .width = Bigger(1, ScaleWidth - .Left * 2)
        .Height = Bigger(1, ScaleHeight - .Top - .Left)
    End With ' mstTab
    
    With paneTab
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' fraTab.Item(mstTab.Index)
    
    With fraScroll
        .width = paneTab.width
        .Height = paneTab.Height
    End With
    
    With flexTab
        .width = .Grid.ColPos(.Grid.Cols - 1) + .Grid.ColWidth(.Grid.Cols - 1) + Screen.TwipsPerPixelX
        .Height = .Grid.RowPos(.Grid.Rows - 1) + .Grid.RowHeight(.Grid.Rows - 1) + Screen.TwipsPerPixelY
    End With ' flexTab(i)
    
    If flexTab.width > fraScroll.width Then
        blnHSBVisible = True
        fraScroll.Height = paneTab.Height - gApp.hsbHeight
    End If
    
    If flexTab.Height > fraScroll.Height Then
        blnVSBVisible = True
        fraScroll.width = paneTab.width - gApp.vsbWidth
    End If
    
    If flexTab.width > fraScroll.width Then
        blnHSBVisible = True
        fraScroll.Height = paneTab.Height - gApp.hsbHeight
    End If
    
    With fraScroll
        hsbSel.Move .Left, .Top + .Height, .width, gApp.hsbHeight
        vsbSel.Move .Left + .width, .Top, gApp.vsbWidth, .Height
    End With
    
    flexTab.Grid.BorderStyle = flexBorderNone
    hsbSel.Min = 0
    hsbSel.max = flexTab.width - fraScroll.width
    hsbSel.LargeChange = flexTab.width
    hsbSel.Visible = blnHSBVisible
    vsbSel.Min = 0
    vsbSel.max = flexTab.Height - fraScroll.Height
    vsbSel.LargeChange = flexTab.Height
    vsbSel.Visible = blnVSBVisible
    
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
    Dim sc As New clsStringConverter
    
    mblnNoData = Not mData.Fetch(mKey)
    
    lblInfo.Caption = mData.FraTopStr
    
    '履歴用文字列追加
    mstrTitle = mstrTitle & " " & lblInfo.Caption
End Sub


'
'   機能: データ取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_FetchComplete(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab.InsertGrid(gd)
    
    If gd.Cols = 3 Then
        
        With flexTab.Grid
            .FixedCols = 0
            .WordWrap = True
            .Visible = True
        End With
        Call flexTab.AutoSize(0, flexTab.Grid.Cols - 1)
        
        ' セル幅を固定値に
        Dim i As Integer
        For i = 0 To flexTab.Grid.Cols - 1
            flexTab.Grid.ColWidth(i) = 3960
        Next

    Else
        flexTab.Grid.Visible = False
    End If
    paneTab.Mode = 2
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
    gApp.Log "Free VRASel"
    
    Call DestroyFlexGrid(flexTab)
    
    Set mKey = Nothing
    Set mData = Nothing
End Sub

