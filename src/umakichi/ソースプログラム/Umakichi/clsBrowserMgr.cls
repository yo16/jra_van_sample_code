VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrowserMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ブラウザマネージャ　すべてのブラウザを管理する\r\n"
'
'   ブラウザマネージャ　Browser Manager
'
'   すべてのブラウザを管理する
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBrowsers As Collection     '' ブラウザフォームを格納するコレクション


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    Set mBrowsers = New Collection
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: 新規にブラウザを生成する
'
'   備考: 引き数 strViewerName - Viewer名
'                strKey - キー
'
Public Sub NewBrowser(strViewerName As String, strKey As String)
    Dim i As Long
    With mBrowsers
        mBrowsers.Add New frmBrowser
        i = mBrowsers.count
        Call mBrowsers.item(i).FirstPage(strViewerName, strKey)
        mBrowsers.item(i).Show
        mBrowsers.item(i).Refresh
    End With
End Sub


'
'   機能: ブラウザをコレクションから削除する
'
'   備考: なし
'
Public Sub Unregist(objForm As Form)
    Dim i As Long
    With mBrowsers
    For i = 1 To .count
        If .item(i) Is objForm Then
            .Remove i
            Exit For
        End If
    Next i
    End With
End Sub


'
'   機能: すべてのブラウザに画面を更新させる
'
'   備考: なし
'
Public Sub NowFetching(value As Boolean)
    Dim i As Long
    Dim Browser As frmBrowser
    
    For i = 1 To mBrowsers.count
        If TypeName(mBrowsers(i)) = "frmBrowser" Then
            Set Browser = mBrowsers(i)
            Browser.Reload
        Else
            gApp.Log "ブラウザコレクションが不正です。"
        End If
    Next i
End Sub



'
'   機能: すべてのブラウザにビューワの再表示を命じる
'
'   備考: なし
'
Public Sub AllReload()
    Dim i As Long
    Dim Browser As frmBrowser
    
    For i = 1 To mBrowsers.count
        If TypeName(mBrowsers(i)) = "frmBrowser" Then
            Set Browser = mBrowsers(i)
            Browser.Reload
        Else
            gApp.Log "ブラウザコレクションが不正です。"
        End If
    Next i
End Sub


'
'   機能: すべてのブラウザにViewerの開放をさせる
'
'   備考: なし
'
Public Sub AllFreeViewer()
    Dim i As Long
    Dim Browser As frmBrowser
    
    For i = 1 To mBrowsers.count
        If TypeName(mBrowsers(i)) = "frmBrowser" Then
            Set Browser = mBrowsers(i)
            Browser.FreeViewer
        Else
            gApp.Log "ブラウザコレクションが不正です。"
        End If
    Next i
End Sub


'
'   機能: すべてのブラウザにメニューパレットの表示状態を伝える
'
'   備考: なし
'
Public Sub AllShowMenuPalette(value As Boolean)
    Dim i As Long
    Dim Browser As frmBrowser
    
    For i = 1 To mBrowsers.count
        If TypeName(mBrowsers(i)) = "frmBrowser" Then
            Set Browser = mBrowsers(i)
            Browser.ShowMenuPalette = value
        Else
            gApp.Log "ブラウザコレクションが不正です。"
        End If
    Next i
End Sub


'
'   機能: ブラウザがすべて閉じられているかどうかの確認
'
'   備考: なし
'
Public Function BrowserExist() As Boolean
    Dim i As Long
    Dim Browser As frmBrowser
    Dim BrowserVisible As Boolean
    
    For i = 1 To mBrowsers.count
        If TypeName(mBrowsers(i)) = "frmBrowser" Then
            Set Browser = mBrowsers(i)
            BrowserVisible = BrowserVisible Or Browser.Visible
        Else
            gApp.Log "ブラウザコレクションが不正です。"
        End If
    Next i
    BrowserExist = BrowserVisible
End Function
