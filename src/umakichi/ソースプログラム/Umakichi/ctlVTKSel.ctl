VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVTKSel 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   ScaleHeight     =   4815
   ScaleWidth      =   6990
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
         Caption         =   "特別登録馬 選択"
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
         Width           =   1875
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2745
      Left            =   150
      TabIndex        =   2
      Top             =   780
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4842
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "特別登録馬開催情報"
      TabPicture(0)   =   "ctlVTKSel.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2205
         Left            =   30
         TabIndex        =   3
         Top             =   360
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3889
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   945
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   1667
         End
      End
   End
End
Attribute VB_Name = "ctlVTKSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   特別登録馬  選択 表示コントロール
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

Private mstrTitle As String
Private mKey As clsKeyRASel
Private mVB As clsViewerBase
Private mViewerState As clsVSNothing
Private mblnNoData As Boolean           '' データ無しフラグ

Private WithEvents mData As clsDataTKSel
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
'
Public Property Get NoData() As Boolean
    NoData = mblnNoData
End Property


'
'   機能: ユーザコントロール初期化
'
'   備考: なし
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    gApp.InitLog Me
    
    mstrTitle = "特別登録馬選択"
    Set mData = New clsDataTKSel
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
    
    With flexTab
        .width = Bigger(1, paneTab.width - .Left)
        .Height = Bigger(1, paneTab.Height - .Top)
    End With ' flexTab(i)
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
    
    Call mData.Fetch(mKey)
    
    
    lblInfo.Caption = mData.FraTopStr
    
    '履歴用文字列追加
    mstrTitle = mstrTitle & " " & lblInfo.Caption
End Sub


'
'   機能: データがそろった
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
        
        ' セル幅を固定に
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
    gApp.Log "Free VSel"
    
    Call DestroyFlexGrid(flexTab)
    
    Set mKey = Nothing
    Set mData = Nothing
End Sub

