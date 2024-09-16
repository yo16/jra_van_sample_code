VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVRCSel 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   ScaleHeight     =   6120
   ScaleWidth      =   11880
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  'なし
      Caption         =   "条件"
      Height          =   1515
      Left            =   180
      TabIndex        =   17
      Top             =   600
      Width           =   8175
      Begin VB.OptionButton optButton 
         Caption         =   "ダミー"
         Height          =   315
         Index           =   0
         Left            =   3060
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optButton 
         Caption         =   "小倉競馬場"
         Height          =   315
         Index           =   10
         Left            =   4920
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   13
         Top             =   1110
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "阪神競馬場"
         Height          =   315
         Index           =   9
         Left            =   3720
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   12
         Top             =   1110
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "京都競馬場"
         Height          =   315
         Index           =   8
         Left            =   2520
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   11
         Top             =   1110
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "中京競馬場"
         Height          =   315
         Index           =   7
         Left            =   1320
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   10
         Top             =   1110
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "中山競馬場"
         Height          =   315
         Index           =   6
         Left            =   120
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   9
         Top             =   1110
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "東京競馬場"
         Height          =   315
         Index           =   5
         Left            =   4920
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   8
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "新潟競馬場"
         Height          =   315
         Index           =   4
         Left            =   3720
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   7
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "福島競馬場"
         Height          =   315
         Index           =   3
         Left            =   2520
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   6
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "函館競馬場"
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   5
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "札幌競馬場"
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   4
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton optButton 
         Caption         =   "GⅠ競走"
         Height          =   315
         Index           =   13
         Left            =   2040
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   3
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton optButton 
         Caption         =   "ダート"
         Height          =   315
         Index           =   12
         Left            =   1080
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   2
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton optButton 
         Caption         =   "芝"
         Height          =   315
         Index           =   11
         Left            =   120
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   1
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lblRed 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   $"ctlVRCSel.ctx":0000
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3360
         TabIndex        =   22
         Top             =   0
         Width           =   4755
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "各競馬場コースレコード"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "中央競馬レコード"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   20
         Top             =   0
         Width           =   1320
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   405
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7305
      Begin VB.Timer tmrUpdateTrigger 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   6480
         Top             =   0
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "レコード表"
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
         Left            =   180
         TabIndex        =   16
         Top             =   90
         Width           =   1110
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   3945
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2130
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   6959
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "レコード選択"
      TabPicture(0)   =   "ctlVRCSel.ctx":00AE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2925
         Left            =   60
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   5159
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1215
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2143
         End
      End
   End
End
Attribute VB_Name = "ctlVRCSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   レコード選択 表示コントロール
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

Private WithEvents mData As clsDataRCSel        '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1

Private mstrTitle       As String               '' ウインドウタイトル
Private mVB             As clsViewerBase
Private mKey            As clsKeyRCSel
Private mViewerState    As clsVSTabOnly
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
    mKey.Key = strKey
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
Public Property Get NoData() As Boolean
    NoData = mblnNoData
End Property


'
'   機能: オプションボタン押下イベント
'
'   備考: なし
'
Private Sub optButton_Click(Index As Integer)
On Error GoTo ErrorHandler
    If mData.NowFetching Then mData.CancelFetching
    mViewerState.LastTabNumber = Index
    Call Update
    tmrUpdateTrigger.Enabled = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: flexTab　クリックイベント
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
    
    '範囲外なら終了
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
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
    
    ' グリッドアイテムをセルから取り出す
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
'   機能: flexTab マウス移動イベント
'
'   備考: リンクハイライト、カーソル形状、ツールチップテキストの切り替え
'
Private Sub flexTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    
    Call flexTab.MouseMoveDriven
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データを取得する
'
'   備考: なし
'
Private Sub mData_FetchComplete(gd As clsGridData)
On Error GoTo ErrorHandler
    Dim i As Long
    Dim strPrevKyori As String
    
    Call flexTab.InsertGrid(gd)
    
    With flexTab.Grid
        'セルを連結
        For i = 0 To mKey.TitleRowUb
            .MergeRow(mKey.TitleRow(i)) = True
        Next i
        mKey.TitleRowClr
        .MergeCells = flexMergeFree
        paneTab.Mode = 2
        
        'セルのサイズ調整
        .WordWrap = True
        Call flexTab.AutoSize(0, .Cols - 1)
        
        ' セル幅を固定に
        flexTab.Grid.ColWidth(0) = 885
        flexTab.Grid.ColWidth(1) = 795
        flexTab.Grid.ColWidth(2) = 1845
        flexTab.Grid.ColWidth(3) = 780
        flexTab.Grid.ColWidth(4) = 1800
        flexTab.Grid.ColWidth(5) = 600
        flexTab.Grid.ColWidth(6) = 1400
        flexTab.Grid.ColWidth(7) = 1140
        flexTab.Grid.ColWidth(8) = 645
        flexTab.Grid.ColWidth(9) = 555
        flexTab.Grid.ColWidth(10) = 855
        
    End With
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データを取得する
'
'   備考: なし
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    paneTab.Mode = 1
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
    Dim strKey  As String
    Dim i   As Integer
    
    Refresh
    If mViewerState.IsNoTouch Then
        For i = 1 To optButton.UBound
            If optButton(i).value Then
                strKey = CStr(optButton.item(i).Index)
                Exit For
            End If
        Next i
    Else
        strKey = mViewerState.LastTabNumber
        optButton(mViewerState.LastTabNumber) = True
    End If
    If strKey <> "" Then
        If Len(strKey) < 2 Then strKey = String$(2 - Len(strKey), "0") & strKey
        mKey.Key = strKey
        tmrUpdateTrigger = True
    End If
End Sub


'
'   機能: アップデートトリガータイマー
'
'   備考: なし
'
Private Sub tmrUpdateTrigger_Timer()
On Error GoTo ErrorHandler
    paneTab.Mode = 0
    mData.CancelFetching
    If Not mData.NowFetching Then
        tmrUpdateTrigger.Enabled = False
        Call mData.Fetch(mKey)
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
    gApp.InitLog Me

    Set mVB = New clsViewerBase     '' ViewerBase オブジェクト
    Set mData = New clsDataRCSel    '' データ取得オブジェクト
    Set mKey = New clsKeyRCSel
    Set mViewerState = New clsVSTabOnly

    gApp.InitLog Me
    mstrTitle = "レコード一覧"
    
    ' FlexGrid設定
    Call mVB.FlexGridCommonSetting(flexTab.Grid)
    With flexTab.Grid
        .FixedCols = 0
        .FixedRows = 1
    End With

    ' Color Assign
    BackColor = gApp.ColBG
    mstTab.BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
    lblFix(0).BackColor = gApp.ColDarkBG
    lblFix(0).ForeColor = Contrast(gApp.ColDarkBG)
    lblFix(1).BackColor = gApp.ColBG
    lblFix(1).ForeColor = Contrast(gApp.ColBG)
    lblFix(2).BackColor = gApp.ColBG
    lblFix(2).ForeColor = Contrast(gApp.ColBG)
    fraHeader.BackColor = gApp.ColBG
    
    '初期状態ではデータを取得，表示しない
    paneTab.Mode = 3
    
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
    
    ' ユーザーコントロールの最低幅を決める
    With UserControl
        .width = Bigger(8000, .width)
        .Height = Bigger(5000, .Height)
    End With
    
    fraTop.width = ScaleWidth - fraTop.Left * 2
    With mstTab
        .width = Bigger(1, ScaleWidth - .Left * 2)
        .Height = Bigger(1, ScaleHeight - .Top - .Left)
    End With
    
    With paneTab
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With
    
    With flexTab
        .width = Bigger(1, paneTab.width - .Left)
        .Height = Bigger(1, paneTab.Height - .Top)
    End With
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
    
    Dim i As Integer
    Dim tmp As Long
    For i = 0 To flexTab.Grid.Cols - 1
        tmp = flexTab.Grid.ColWidth(i)
    Next
    
    Set mVB = Nothing
    Set mKey = Nothing
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
    gApp.Log "Free VRCSel"
    
    Call DestroyFlexGrid(flexTab)
    
    Set mKey = Nothing
    Set mData = Nothing
End Sub

