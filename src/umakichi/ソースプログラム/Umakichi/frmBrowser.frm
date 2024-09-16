VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "馬吉"
   ClientHeight    =   8130
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11775
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows の既定値
   Begin MSComDlg.CommonDialog dlgHelpFile 
      Left            =   5940
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilsTbrCmd 
      Left            =   1710
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.HScrollBar hsbPane 
      Height          =   285
      Left            =   540
      TabIndex        =   4
      Top             =   4350
      Width           =   4155
   End
   Begin VB.VScrollBar vsbPane 
      Height          =   2175
      Left            =   4890
      TabIndex        =   3
      Top             =   1920
      Width           =   255
   End
   Begin VB.Frame fraScrollPane 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   600
      TabIndex        =   2
      Top             =   1860
      Width           =   4065
   End
   Begin VB.Timer tmrToolbarBug 
      Interval        =   1
      Left            =   810
      Top             =   1140
   End
   Begin MSComctlLib.ImageList ilsSmallIcons 
      Left            =   90
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar staStatusBar 
      Align           =   2  '下揃え
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   7845
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20258
            MinWidth        =   176
            Text            =   "　"
            TextSave        =   "　"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Width           =   706
            MinWidth        =   706
            Text            =   "｜／－＼－"
            TextSave        =   "｜／－＼－"
            Object.ToolTipText     =   "アニメ"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTop 
      Align           =   1  '上揃え
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1614
      _CBWidth        =   11775
      _CBHeight       =   915
      _Version        =   "6.7.9782"
      BandBackColor1  =   16711935
      Child1          =   "tbsBrowser"
      MinWidth1       =   1005
      MinHeight1      =   330
      Width1          =   2970
      NewRow1         =   0   'False
      MinWidth2       =   1005
      MinHeight2      =   525
      Width2          =   1410
      NewRow2         =   0   'False
      Visible2        =   0   'False
      Child3          =   "tbdTitleBand"
      MinHeight3      =   300
      Width3          =   495
      NewRow3         =   -1  'True
      Begin Umakichi.ctlToolBars tbsBrowser 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   120
         Width           =   2775
         _ExtentX        =   0
         _ExtentY        =   582
      End
      Begin Umakichi.ctlTitleBand tbdTitleBand 
         Height          =   300
         Left            =   165
         TabIndex        =   5
         Top             =   585
         Width           =   11520
         _ExtentX        =   8916
         _ExtentY        =   529
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ファイル(&F)"
      Begin VB.Menu mnuFileSub 
         Caption         =   "新規ウインドウ(&N)"
         Index           =   0
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "ホームメニュー(&H)"
            Index           =   0
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "出馬表(&D)"
            Index           =   1
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "特別登録馬(&T)"
            Index           =   2
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "坂路調教(&C)"
            Index           =   3
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "レコード(&R)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "閉じる(&C)"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "馬吉の終了(&X)"
         Index           =   3
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "設定(&V)"
      Begin VB.Menu mnuConfig 
         Caption         =   "馬吉設定ダイアログ(&C)"
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "標準のボタン(&S)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "メニューパレット(&P)"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuDB 
      Caption         =   "データベース(&D)"
      Begin VB.Menu mnuDBSub 
         Caption         =   "更新(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuDBSub 
         Caption         =   "最適化(&O)"
         Index           =   1
      End
      Begin VB.Menu mnuDBSub 
         Caption         =   "データセットアップ(&S)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuBrowser 
      Caption         =   "BrowserContextメニュー"
      Visible         =   0   'False
      Begin VB.Menu mnuBrowserSub 
         Caption         =   "戻る(&B)"
         Index           =   0
      End
      Begin VB.Menu mnuBrowserSub 
         Caption         =   "進む(&N)"
         Index           =   1
      End
      Begin VB.Menu mnuBrowserSub 
         Caption         =   "ホーム(&H)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuToolBar 
      Caption         =   "ブラウザツールバーContextメニュ－"
      Visible         =   0   'False
      Begin VB.Menu mnuToolBarSubText 
         Caption         =   "テキストの表示"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuLink 
      Caption         =   "LinkContextメニュー"
      Visible         =   0   'False
      Begin VB.Menu mnuLinkSub 
         Caption         =   "新しいウインドウで開く(&W)"
         Index           =   0
      End
      Begin VB.Menu mnuLinkSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLinkSub 
         Caption         =   "戻る(&B)"
         Index           =   2
      End
      Begin VB.Menu mnuLinkSub 
         Caption         =   "進む(&N)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "ヘルプ(&H)"
      Begin VB.Menu mnuHelpFile 
         Caption         =   "馬吉ヘルプ"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUma 
         Caption         =   "馬吉について(&U)"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   ブラウザーフォーム
'
'   Viewerを乗せるコンテナ
'   WebBrowserに似たインターフェイスを持つ。
'   Viewerがはみ出す場合スクロールバーで制御。
'
Option Explicit

Private WithEvents mextViewer As VBControlExtender ' Viewerコントロール参照　イベント取得用
Attribute mextViewer.VB_VarHelpID = -1

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mobjViewer As Object                       ' Viewerコントロール参照　メソッドコール用
Attribute mobjViewer.VB_VarHelpID = -1

Private mblnDoubleGameFlag As Boolean
Private mHistoryMgr As clsHistoryMgr               ' 履歴管理オブジェクト

Private mstrViewerContextMenuViewerName As String  ' Viewerコンテキストメニューのリンク先Viewer名
Private mstrViewerContextMenuKey As String         ' Viewerコンテキストメニューのリンク先キー

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: メニューパレットの表示のチェックを設定する
'
'   備考: なし
'
Public Property Let ShowMenuPalette(RHS As Boolean)
    mnuViewSub(1).Checked = RHS
End Property

'
'   機能: メニューパレットの表示のチェック状態を返す
'
'   備考: なし
'
Public Property Get ShowMenuPalette() As Boolean
    ShowMenuPalette = mnuViewSub(1).Checked
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 初期表示画面の設定
'
'   備考: 引数 strViewerName - Viewer名, strKey - Viewerがデータを特定する為のキー
'
Public Sub FirstPage(strViewerName As String, strKey As String)
    Dim newHistory As clsHistoryItem
    
    Set newHistory = New clsHistoryItem
    
    Set mextViewer = Controls.Add("Umakichi.ctlV" & strViewerName, VName)
    Set mobjViewer = mextViewer
    mobjViewer.key = strKey
    
    With newHistory
        .key = strKey
        .ViewerName = strViewerName
        .Title = mobjViewer.Title
    End With
    mHistoryMgr.Add newHistory
    
    Call FitViewer
    
    mextViewer.Visible = True
        
    Call TitleChange(mobjViewer.Title, strViewerName)
    Call SetHistoryToToolbar
    Call ChangeToolBar(strViewerName, strKey)

End Sub


'
'   機能: ポップアップメニュー表示処理
'
'   備考: なし
'
Public Sub ShowPopupMenu(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuBrowser, vbPopupMenuRightButton
    End If
End Sub


'
'   機能: Viewerのリロード
'
'   備考: なし
'
Public Sub Reload()
    With mHistoryMgr.Current
        Call ChangeViewer(.ViewerName, .key)
    End With
End Sub


'
'   機能: Viewerにホームを表示
'
'   備考: なし
'
Public Sub GotoHome()
    Call GoToNextViewer("Home", "Empty")
End Sub


'
'   機能: Viewerを１つ前に戻す
'
'   備考: なし
'
Public Sub BackOne()
    Call historyBack(1)
End Sub


'
'   機能: Viewerの履歴を開放する
'
'   備考: なし
'
Public Sub FreeViewer()
    Call mobjViewer.Free
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クールバーの高さ変更イベント
'
'   備考: なし
'
Private Sub cbrTop_HeightChanged(ByVal newHeight As Single)
On Error GoTo Errorhandler

    Call FitViewer(newHeight)
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: クールバーマウスダウンイベント
'
'   備考: なし
'
Private Sub cbrTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Errorhandler
    If Button = vbRightButton Then
    
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォーム初期化イベント
'
'   備考: なし
'
Private Sub Form_Initialize()
On Error GoTo Errorhandler

    Set mHistoryMgr = New clsHistoryMgr
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームロードイベント
'
'   備考: なし
'
Private Sub Form_Load()
On Error GoTo Errorhandler

    mnuViewSub(1).Checked = gApp.R_MenuVisible
    
    With ilsSmallIcons
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add 1, , LoadResPicture(101, vbResIcon)
        .ListImages.Add 2, , LoadResPicture(102, vbResIcon)
        .ListImages.Add 3, , LoadResPicture(103, vbResIcon)
        .ListImages.Add 4, , LoadResPicture(104, vbResIcon)
        .ListImages.Add 5, , LoadResPicture(105, vbResIcon)
        .ListImages.Add 6, , LoadResPicture(106, vbResIcon)
        .ListImages.Add 7, , LoadResPicture(107, vbResIcon)
        Me.Icon = LoadResPicture(100, vbResIcon)
    End With
    With tbsBrowser
        .num = 3
        With .ToolBar(0)
            .ImageList = ilsSmallIcons
            .Buttons.Add 1, "BACK", "戻る", tbrDropdown
            .Buttons.item(1).Image = 1
            .Buttons.Add 2, "NEXT", "進む", tbrDropdown
            .Buttons.item(2).Image = 2
            .Buttons.Add 3, "HOME", "ホーム"
            .Buttons.item(3).Image = 3
            .Buttons.Add 4, "UPDT", "更新"
            .Buttons.item(4).Image = 4
            .Buttons.Add 5, "CONF", "設定"
            .Buttons.item(5).Image = 5
            .width = .Buttons.item(1).width + .Buttons.item(2).width + .Buttons.item(3).width + _
                    .Buttons.item(4).width + .Buttons.item(5).width
        End With
        With .ToolBar(2)
            .ImageList = ilsSmallIcons
            .Buttons.Add 1, "RTOPEN", "速報取得"
            .Buttons.item(1).Image = 7
        End With
        Call .fit
    End With
    cbrTop.Bands(1).MinWidth = tbsBrowser.ToolBar(0).width
    cbrTop.Bands(1).width = tbsBrowser.ToolBar(0).width
    
    vsbPane.width = gApp.vsbWidth
    hsbPane.Height = gApp.hsbHeight

    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームアンロードイベント
'
'   備考: なし
'
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errorhandler
    Call mobjViewer.Free
    If Err.Number <> 0 Then
        gApp.ErrLog
        gApp.Log "frmBrowser.Form_Unload() " & TypeName(mobjViewer) & "はFree()を実装してないかもしれません>開発者"
    End If
    gApp.BrowserUnregist Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: メニューの設定イベント
'
'   備考: なし
'
Private Sub mnuConfig_Click()
    gApp.Configulation
End Sub


'
'   機能: メニューのヘルプイベント
'
'   備考: なし
'
Private Sub mnuHelpFile_Click()
        Call ShowHtmlHelp
End Sub


'
'   機能: メニューのヘルプ－馬吉についてイベント
'
'   備考: なし
'
Private Sub mnuHelpUma_Click()
On Error GoTo Errorhandler

    Dim aboutWindow As New frmAbout
    aboutWindow.Show vbModal, Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ツールバー選択イベント
'
'   備考: なし
'
Private Sub mnuToolBarSubText_Click()
On Error GoTo Errorhandler

    Dim i As Long
    mnuToolBarSubText.Checked = Not mnuToolBarSubText.Checked
    If mnuToolBarSubText.Checked Then
        For i = 1 To tbsBrowser.ToolBar(0).Buttons.count
            tbsBrowser.ToolBar(0).Buttons(i).Caption = tbsBrowser.ToolBar(0).Buttons(i).Description
        Next i
    Else
        For i = 1 To tbsBrowser.ToolBar(0).Buttons.count
            tbsBrowser.ToolBar(0).Buttons(i).Caption = ""
        Next i
    End If
    cbrTop.Bands(1).MinHeight = tbsBrowser.ToolBar(0).ButtonHeight
    cbrTop.Bands(1).MinWidth = tbsBrowser.ToolBar(0).width
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームリサイズイベント
'
'   備考: オブジェクトが最初に表示されたときに発生するほか、
'         最大化、最小化、元のサイズに戻すなど、
'         オブジェクトのウィンドウ状態が変化したときにも発生。
'
Private Sub Form_Resize()
On Error GoTo Errorhandler
    Call FitViewer
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: Viewerをフォーム全体にフィットさせる
'
'   備考: Viewerの大きさから、ScrollPaneよりも大きい場合、スクロールバーを表示する
'
Private Sub FitViewer(Optional newHeight As Single)
    Dim blnHSBVisible As Boolean
    Dim blnVSBVisible As Boolean
    
    blnHSBVisible = False
    blnVSBVisible = False
    
    cbrTop.Height = IIf(newHeight = 0, cbrTop.Height, newHeight)
    
    ' スクロールペインを表示領域最大限に合わせる
    With fraScrollPane
        .Left = 0                                                               ' 左端
        .Top = cbrTop.Height                                                    ' クールバーの下
        .width = Bigger(ScaleWidth, 1)                                          ' 枠まで幅いっぱい
        .Height = Bigger(ScaleHeight - cbrTop.Height - staStatusBar.Height, 1)  ' 枠まで高さぴったり
    End With
    
    Set mextViewer.Container = fraScrollPane
    
    ' Viewer をスクロールペインの大きさに合わせる
    With mextViewer
        .Left = 0                       ' スクロールペインの左端
        .Top = 0                        ' スクロールペインの右端
        .width = fraScrollPane.width    ' 最大幅
        .Height = fraScrollPane.Height  ' 最大高
    End With
    
    ' Viewerの幅がスクロールペインの幅より大きい場合
    If mextViewer.width > fraScrollPane.width Then
        ' 水平SB 可視に決定
        blnHSBVisible = True
        ' 水平SBの分、スクロールペインを狭くする
        fraScrollPane.Height = Bigger(fraScrollPane.Height - hsbPane.Height, 1)
        ' スクロールペインの高さにViewerを合わせる
        mextViewer.Height = fraScrollPane.Height
    End If
    
    ' Viewerの高さがスクロールペインの高さより大きい場合
    If mextViewer.Height > fraScrollPane.Height Then
        ' 垂直 SB 可視に決定
        blnVSBVisible = True
        ' 垂直スクロールバーの分、スクロールペインを狭くする
        fraScrollPane.width = Bigger(fraScrollPane.width - vsbPane.width, 1)
        ' スクロールペインの幅にViewerを合わせる
        mextViewer.width = fraScrollPane.width
    End If
    
    ' 垂直スクロールバーの出現で幅が狭くなった結果
    ' Viewerの幅がスクロールペインの幅より大きい場合
    If blnHSBVisible = False And mextViewer.width > fraScrollPane.width Then
        ' 水平SB 可視に決定
        blnHSBVisible = True
        ' 水平スクロールバーの分、スクロールペインを狭くする
        fraScrollPane.Height = Bigger(fraScrollPane.Height - hsbPane.Height, 1)
        ' スクロールペインの高さにViewerを合わせる
        mextViewer.Height = fraScrollPane.Height
    End If
    
    ' 水平SB を配置
    If blnHSBVisible Then
        With hsbPane
            .Left = fraScrollPane.Left
            .Top = fraScrollPane.Top + fraScrollPane.Height
            .width = fraScrollPane.width
        End With
    End If
    hsbPane.Visible = blnHSBVisible
    
    ' 垂直SB を配置
    If blnVSBVisible Then
        With vsbPane
            .Left = fraScrollPane.Left + fraScrollPane.width
            .Top = fraScrollPane.Top
            .Height = fraScrollPane.Height
        End With
    End If
    vsbPane.Visible = blnVSBVisible
    
    hsbPane.max = mextViewer.width - fraScrollPane.width
    hsbPane.LargeChange = mextViewer.width
    hsbPane.SmallChange = mextViewer.width / 10
    vsbPane.max = mextViewer.Height - fraScrollPane.Height
    vsbPane.LargeChange = mextViewer.Height
    vsbPane.SmallChange = mextViewer.Height / 10

End Sub


'
'   機能: 水平スクロールバー変更イベント
'
'   備考: なし
'
Private Sub hsbPane_Change()
On Error GoTo Errorhandler
    mextViewer.Left = -hsbPane.value
    mextViewer.SetFocus                 ' Viewerにフォーカスをセット
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 水平スクロールバードラック中イベント
'
'   備考: なし
'
Private Sub hsbPane_Scroll()
On Error GoTo Errorhandler
    mextViewer.Left = -hsbPane.value
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: レースチェンジャー選択イベント
'
'   備考: なし
'
Private Sub tbdTitleBand_Change(key As clsKeyRA)
On Error GoTo EH
    With mHistoryMgr.Current
        Call GoToNextViewer(.ViewerName, key.str)
    End With
    Exit Sub
EH:
    gApp.ErrLog
End Sub


'
'   機能: ツールバー表示タイマーイベント
'
'   備考: なし
'
Private Sub tmrToolbarBug_Timer()
On Error GoTo Errorhandler:
    If Me.Visible = True Then
        tmrToolbarBug.Enabled = False
        cbrTop.Bands(1).Child.fit
    Else
        '
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 垂直スクロールバー変更イベント
'
'   備考: なし
'
Private Sub vsbPane_Change()
On Error GoTo Errorhandler
    mextViewer.Top = -vsbPane.value
    mextViewer.SetFocus             ' Viewerにフォーカスをセット
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 垂直スクロールバードラック中イベント
'
'   備考: なし
'
Private Sub vsbPane_Scroll()
On Error GoTo Errorhandler
    mextViewer.Top = -vsbPane.value
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ファイルサブメニュー選択イベント
'
'   備考: なし
'
Private Sub mnuFileSub_Click(Index As Integer)
On Error GoTo Errorhandler
Select Case Index
    Case 0  ' 新規ウインドウ
    Case 1 ' ボーダー
    Case 2 ' 閉じる
        Unload Me
    Case 3 ' 馬吉全体の終了
        If vbYes = MsgBox("馬吉を終了しますか？", vbYesNo + vbQuestion, "馬吉：終了の確認") Then
            gApp.ExitApp
        End If
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ファイル新規サブメニュー選択イベント
'
'   備考: なし
'
Private Sub mnuFileNewSub_Click(Index As Integer)
On Error GoTo Errorhandler
    Select Case Index
    Case 0  ' ホーム
        Call gApp.NewWindow("Home", "Empty")
    Case 1  ' 出馬表
        Call gApp.NewWindow("RAKaiSel", "Empty")
    Case 2  ' 特別登録馬
        Call gApp.NewWindow("TKKaiSel", "Empty")
    Case 3  ' 販路調教
        Call gApp.NewWindow("HCSel", "Empty")
    Case 4  ' レコード
        Call gApp.NewWindow("RCSel", "Empty")
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ＤＢサブメニュー選択イベント
'
'   備考: なし
'
Private Sub mnuDBSub_Click(Index As Integer)
On Error GoTo Errorhandler
    Dim MsgResult       As VbMsgBoxResult
    
    Select Case Index
    Case 0  ' 更新
        If MsgBox("データの更新をしますか", vbYesNo + vbQuestion, "馬吉：データ更新処理開始の確認") = vbYes Then
            Call gApp.DBUpdate
            With mHistoryMgr.Current
                Call ChangeViewer(.ViewerName, .key)
            End With
            
        End If
    ' 最適化
    Case 1
        Call gApp.DBCompact
    ' セットアップ
    Case 2
        Call setup

    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: データセットアップ
'
'   備考: なし
'
Private Sub setup()
On Error GoTo Errorhandler
    Dim LastTimeBackup  As String
    Dim ConfigFirst     As frmConfigFirst
    Dim DBUpdateForm    As frmDBUpdate
    Dim MsgResult       As VbMsgBoxResult
    
    ' 再セットアップかどうか
    If gApp.R_JVDLastTime <> String$(14, "0") Then
        MsgResult = MsgBox(gApp.R_DBPath & _
            "データベースはすでにセットアップされています。" & vbCrLf & _
            "セットアップをやり直しますか？", vbExclamation + vbYesNo + vbDefaultButton2, "馬吉：再セットアップの確認")
        If MsgResult = vbYes Then
            ' 最終取得時刻を退避
            LastTimeBackup = gApp.R_JVDLastTime
            ' 最終取得時刻をリセット
            gApp.R_JVDLastTime = String$(14, "0")
            ' データセットアップの設定画面を出す
            Set ConfigFirst = New frmConfigFirst
            ConfigFirst.Show vbModal
            If ConfigFirst.ButtonType <> "OK" Then
                ' 最終取得時刻を復帰
                gApp.R_JVDLastTime = LastTimeBackup
                MsgBox "再セットアップは行われませんでした。", vbInformation, "馬吉：再セットアップキャンセル"
                Exit Sub
            End If
            ' 取得を行う
            Set DBUpdateForm = New frmDBUpdate
            DBUpdateForm.Show vbModal
            If DBUpdateForm.AfterJVOpen = False Then
                gApp.R_JVDLastTime = LastTimeBackup
            End If
            gApp.AllReload
        End If
    Else
        Call gApp.DBUpdate
    End If
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: リンクサブメニュー選択イベント
'
'   備考: 右クリックメニューです。
'
Private Sub mnuLinkSub_Click(Index As Integer)
On Error GoTo Errorhandler
Select Case Index
    Case 0
        Call gApp.NewWindow(mstrViewerContextMenuViewerName _
                            , mstrViewerContextMenuKey)
    Case 1 ' ボーダー
    
    Case 2 ' 戻る
        Call historyBack(1)
    Case 3 ' 進む
        Call historyNext(1)
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 表示サブメニュー選択イベント
'
'   備考: なし
'
Private Sub mnuViewSub_Click(Index As Integer)
On Error GoTo Errorhandler
    Select Case Index
    Case 0 ' 標準のボタン
        With mnuViewSub(Index)
            .Checked = Not .Checked
            cbrTop.Bands(1).Visible = .Checked
            tbsBrowser.fit
        End With
    Case 1 ' メニューパレット
        Call gApp.ShowMenuPalette(Not mnuViewSub(1).Checked)
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ブラウザメニュー選択イベント
'
'   備考: なし
'
Private Sub mnuBrowserSub_Click(Index As Integer)
On Error GoTo Errorhandler
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub

'
'   機能: Viewerのイベント処理
'
'   備考: タイトル変更、画面切り替え等
'
Private Sub mextViewer_ObjectEvent(Info As EventInfo)
On Error GoTo Errorhandler
    Dim newViewer As Control
    Dim obj As Object
    
    Dim ViewerName As String
    Dim key As String

    Select Case Info.Name
        Case "WindowTitle"
            ' ウインドウタイトル変更
            ViewerName = Mid(TypeName(mextViewer), 5)
            Call TitleChange(Info.EventParameters.item(0), ViewerName)
        
        Case "ChangeTo"
            ' 画面変更イベント
            
            Call GoToNextViewer(Info.EventParameters.item(0), Info.EventParameters.item(1))
            
        Case "NewWindow"
            ' 新規ウインドウ
            
            ' イベント引き数を受け取る
            With Info.EventParameters
                ViewerName = .item(0)
                key = .item(1)
            End With
            
            Call gApp.NewWindow(ViewerName, key)
        
        Case "LinkContextMenu"
            ' 右クリックメニュー
            
            ' リンク先情報をモジュール変数に保存、メニューイベントで拾う
            With Info.EventParameters
                mstrViewerContextMenuViewerName = .item(0)
                mstrViewerContextMenuKey = .item(1)
            End With
                        
            Me.PopupMenu mnuLink, vbPopupMenuRightButton
            
        Case "Reload"
            Call Reload
            
        
        Case "StatusBarTextChange"
            ' 進歩
            
            staStatusBar.Panels(1).Text = Info.EventParameters(0)
            
        Case "Progression"
            
        Case Else
            gApp.Log "Unknown Viewer Event (" & Info.Name & ")"
    
    End Select ' Info.Name
    
    Exit Sub

Errorhandler:
    gApp.ErrLog

End Sub


'
'   機能: Viewerを切り替えヒストリを進めツールバーの更新
'
'   備考: なし
'
Private Sub GoToNextViewer(ViewerName As String, key As String)
On Error GoTo Errorhandler
    Dim newHistory As clsHistoryItem
    
    ' 最終状態データを取得
    On Error Resume Next
        Set mHistoryMgr.Current.ViewerState = mobjViewer.ViewerState
        If Err.Number <> 0 Then
            gApp.ErrLog
            gApp.Log TypeName(mobjViewer) & "にはViewerStatus が実装されていないと思われます。"
        End If
    On Error GoTo Errorhandler
    
    
    ' 切り替え
    If ChangeViewer(ViewerName, key) Then
        ' 成功なら後処理
        ' 履歴保存
        Set newHistory = New clsHistoryItem
        With newHistory
            .ViewerName = ViewerName
            .key = key
            .Title = mobjViewer.Title
            .DateTime = Timer
        End With
        mHistoryMgr.Add newHistory
        
        ' 履歴をツールバーに反映
        Call SetHistoryToToolbar

    Else
        ' 切り替え失敗の場合
        MsgBox "開けません。", vbExclamation, "馬吉：エラー"
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ウインドウタイトルを設定
'
'   備考: 表通部分と引き数を連結して設定する
'
Private Sub TitleChange(strTitle As String, ViewerName As String)
On Error GoTo Errorhandler

    Me.Caption = strTitle & " : " & cAppName
    
    tbdTitleBand.Caption = strTitle

    cbrTop.Bands(3).MinWidth = tbdTitleBand.width
    Exit Sub

Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 表Viewer名を返す
'
'   備考: ダブルバッファリングのような処理の為
'
Private Function VName() As String
    VName = "Viewer" & IIf(mblnDoubleGameFlag, "1", "0")
End Function


'
'   機能: 裏Viewer名を返す
'
'
'   備考: 返り値 裏Viewer名, ダブルバッファリングのような処理の為
'
Private Function VNameBack() As String
    VNameBack = "Viewer" & IIf(mblnDoubleGameFlag, "0", "1")
End Function


'
'   機能: Viewerの切り替え
'
'   備考: 引き数 strViewerName - Viewer名
'                strKey        - Viewerが表示データを特定する為のキー
'                viewerState   - Viewer状態
'         裏Viewerとして生成した後､表Viewerを破棄､そして裏を表としている｡
'         これは、生成失敗した時のため旧Viewerを保持しておく間
'         コントロール名が重複出来ないため。
'         あくまで、Viewerを張り替えるだけで、ヒストリの進行等その他の作業は行わない
'         次の画面へリンクを飛ぶ場合は、GoToNextViewerをよぶ
'
Private Function ChangeViewer(strViewerName As String, strKey As String, _
                              Optional ViewerState As clsIViewerState) As Boolean
On Error GoTo Errorhandler
    Dim mp           As New clsPointer '' マウスポインタ
    Dim objNewViewer As Object         '' 新規Viewer
    Dim faze         As Long
    
    
    faze = 1
    
    Call mp.SetBusyPointer(Me)
    staStatusBar.Panels(1).Text = "読み込み中..."
    
    faze = 2
    
    ' 裏Viewer作成
    gApp.Log "CreateControl: " & VNameBack
    Set objNewViewer = Controls.Add("Umakichi.ctlV" & strViewerName, VNameBack)
    
    faze = 3
    
    ' 最終状態データをセット
    On Error Resume Next
        If Not ViewerState Is Nothing Then
            Set objNewViewer.ViewerState = ViewerState
            If Err.Number <> 0 Then
                gApp.ErrLog
                gApp.Log TypeName(objNewViewer) & "にはViewerStatus が実装されていないと思われます。"
            End If
        End If
    On Error GoTo Errorhandler
    
    ' 裏Viewerにキーの設定 = 取得開始
    gApp.Log "ChangeViewer SetKey"
    objNewViewer.key = strKey
    
    ' 表Viewerのイベントを拾わないようにする
    Set mextViewer = Nothing
    
    ' データが無い場合、ホームへ移動させる
    If objNewViewer.NoData Then
        Call Controls.Remove(VNameBack)
        Set objNewViewer = Controls.Add("Umakichi.ctlVHome", VNameBack)
        strViewerName = "Home"
        strKey = Empty
    End If
    
    ' 表Viewerの終了処理プロシージャを呼ぶ
    gApp.Log "ChangeViewer Free"
    Call mobjViewer.Free
    
    ' 表Viewer削除
    gApp.Log "Unload: " & VName
    Call Controls.Remove(VName)
    
    ' ウインドウタイトルを裏Viewerから取得
    gApp.Log "ChangeViewer TitleCange"
    Call TitleChange(objNewViewer.Title, strViewerName)
    
    ' イベントを受け取る為コントロールエクステンダとして保持
    gApp.Log "ChangeViewer Set mextViewer"
    Set mextViewer = objNewViewer
    
    ' アクセスの為オブジェクトとして保持
    gApp.Log "ChangeViewer Set mobjViewer"
    Set mobjViewer = objNewViewer
    
    ' 裏Viewerのサイズをウインドウに密着するよう変更
    gApp.Log "ChangeViewer Fit Viewer"
    FitViewer
    
    ' 裏Viewerの表示
    mextViewer.Visible = True
    
    ' ツールバー変更
    gApp.Log "ChangeViewer ToolBar Setting"
    Call ChangeToolBar(strViewerName, strKey)
    
    ' 裏表切り替え
    gApp.Log "DoubleGameFlag Before: " & mblnDoubleGameFlag
    mblnDoubleGameFlag = Not mblnDoubleGameFlag
    gApp.Log "DoubleGameFlag After : " & mblnDoubleGameFlag
    
    staStatusBar.Panels(1).Text = ""
    
    ChangeViewer = True
    Exit Function

Errorhandler:
    gApp.ErrLog
    If faze < 3 Then
    Call Controls.Remove(VNameBack)
    ChangeViewer = False
    Else
    Resume Next
    End If
End Function


'
'   機能: Viewerに応じて、Viewer用ツールバーを表示
'
'   備考: 引き数 strViewerName - Viewer名, strKey - 呼び出しキー(RaceChanger初期化用)
'
Private Sub ChangeToolBar(strViewerName As String, strKey As String)
On Error GoTo Errorhandler
    Dim i As Long
    Dim max As Long
    Dim raceChnagerKey As New clsKeyRASel
    
    Select Case strViewerName
    Case "RA", "OD"
        Set mobjViewer.ToolBar = tbsBrowser  ' 独自ボタン設定
        tbsBrowser.ToolBar(1).Visible = True ' 独自ボタン表示
        tbsBrowser.ToolBar(2).Visible = True ' 速報取得ボタン表示
        tbsBrowser.fit
        raceChnagerKey.str = strKey
        
        Call tbdTitleBand.ShowRaceChanger(raceChnagerKey)
        cbrTop.Bands(1).MinWidth = tbsBrowser.MinWidth
        cbrTop.Bands(1).width = tbsBrowser.MinWidth
    
    Case "HK"
        Set mobjViewer.ToolBar = tbsBrowser
        With tbsBrowser
            .ToolBar(1).Visible = False  ' 独自ボタン非表示
            .ToolBar(2).Visible = True   ' 速報取得ボタン表示
            .fit
        End With
        Call tbdTitleBand.ShowRaceChanger(Nothing)
    Case "TK"
        With tbsBrowser
            .ToolBar(1).Visible = False   ' 独自ボタン非表示
            .ToolBar(2).Visible = False   ' 速報取得ボタン非表示
            .fit
        End With
        raceChnagerKey.str = strKey
        Call tbdTitleBand.ShowRaceChanger(raceChnagerKey, True) ' TKmode ON
    Case Else
        With tbsBrowser
            .ToolBar(1).Visible = False   ' 独自ボタン非表示
            .ToolBar(2).Visible = False   ' 速報取得ボタン非表示
            .fit
        End With
        Call tbdTitleBand.ShowRaceChanger(Nothing)
    End Select
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 戻る、進むボタンのドロップダウンメニューに履歴を登録する
'
'   備考: 最大10件までに制限してある
'
Private Sub SetHistoryToToolbar()
    Dim i As Long
    Dim tmpHistory As clsHistoryItem
    
    tbsBrowser.ToolBar(0).Buttons(1).ButtonMenus.Clear
    tbsBrowser.ToolBar(0).Buttons(2).ButtonMenus.Clear
    
    ' 戻るボタン
    For i = 1 To 10
        Set tmpHistory = mHistoryMgr.Current(-i)
        If tmpHistory Is Nothing Then
            Exit For
        End If
        tbsBrowser.ToolBar(0).Buttons(1).ButtonMenus.Add , , mHistoryMgr.CurrentNum - i & tmpHistory.Title
    Next i
    ' 先頭なら無効にする
    tbsBrowser.ToolBar(0).Buttons(1).Enabled = Not mHistoryMgr.IsFirst
    
'    ' 進むボタン
    For i = 1 To 10
        Set tmpHistory = mHistoryMgr.Current(i)
        If tmpHistory Is Nothing Then
            Exit For
        End If
        tbsBrowser.ToolBar(0).Buttons(2).ButtonMenus.Add , , mHistoryMgr.CurrentNum + i & tmpHistory.Title
    Next i
    ' 最終なら無効にする
    tbsBrowser.ToolBar(0).Buttons(2).Enabled = Not mHistoryMgr.IsLast

End Sub


'
'   機能: ブラウザ用ツールバーのクリックイベント
'
'   備考: 履歴を戻る、進むの処理
'
Private Sub tbsBrowser_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
On Error GoTo Errorhandler
    gApp.Log "Browser Catch TBSClick"
    Select Case Button.key
    ' 戻るボタン
    Case "BACK"
        Call historyBack(1)
    ' 進むボタン
    Case "NEXT"
        Call historyNext(1)
    ' ホームボタン
    Case "HOME"
        Call GoToNextViewer("Home", "Empty")
    ' 更新ボタン
    Case "UPDT"
        If MsgBox("データの更新をしますか", vbYesNo + vbQuestion, "馬吉：データ更新処理開始の確認") = vbYes Then
            Call gApp.DBUpdate
            With mHistoryMgr.Current
                Call ChangeViewer(.ViewerName, .key)
            End With
        End If
    ' 設定ボタン
    Case "CONF"
        Call gApp.Configulation
        Call Reload
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ブラウザ用ツールバーのドロップダウンメニューのクリックイベント
'
'   備考: 履歴を一度に何段階か戻る、進むの処理
'
Private Sub tbsBrowser_ButtonMenuClick(Index As Integer, ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo Errorhandler
Select Case ButtonMenu.Parent.key
    Case "BACK"
        gApp.Log ButtonMenu.Index
        Call historyBack(ButtonMenu.Index)
    Case "NEXT"
        Call historyNext(ButtonMenu.Index)
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 履歴を１段階または何段階か戻りViewerを変更する
'
'   備考: ツールバーの変更も行う
'
Private Sub historyBack(lngStep As Long)
On Error GoTo Errorhandler
    gApp.Log "戻る"
    
    ' 最終状態データを取得
    On Error Resume Next
        Set mHistoryMgr.Current.ViewerState = mobjViewer.ViewerState
        If Err.Number <> 0 Then
            gApp.ErrLog
            gApp.Log TypeName(mobjViewer) & "にはViewerStatus が実装されていないと思われます。"
        End If
    On Error GoTo Errorhandler
    
    If Not mHistoryMgr.IsFirst Then
'        ' ヒストリポインタの二つ前が前画面
        Call mHistoryMgr.Move(-lngStep)
        With mHistoryMgr.Current
            Call ChangeViewer(.ViewerName, .key, .ViewerState)
        End With
        Call SetHistoryToToolbar
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 履歴を１段階または何段階か進みViewerを変更する
'
'   備考: ツールバーの変更も行う
'
Private Sub historyNext(lngStep As Long)
On Error GoTo Errorhandler

    gApp.Log "進む"
        
    ' 最終状態データを取得
    On Error Resume Next
        Set mHistoryMgr.Current.ViewerState = mobjViewer.ViewerState
        If Err.Number <> 0 Then
            gApp.ErrLog
            gApp.Log TypeName(mobjViewer) & "にはViewerStatus が実装されていないと思われます。"
        End If
    On Error GoTo Errorhandler
    
    Call mHistoryMgr.Move(lngStep)
    With mHistoryMgr.Current
        Call ChangeViewer(.ViewerName, .key, .ViewerState)
    End With
    
    ' ツールバーを更新
    Call SetHistoryToToolbar
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: ブラウザ用ツールバーのマウスダウンイベント
'
'   備考: ツールバーの変更も行う
'
Private Sub tbrBrowser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Errorhandler
    If Button = vbRightButton Then
        PopupMenu mnuToolBar, vbPopupMenuRightButton
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub
