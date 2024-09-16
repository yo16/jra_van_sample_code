VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Viewer基底クラス
'
'   ctlV* はすべて、このインスタンスを持つこと
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: GUI の共通設定　初期化
'
'   備考: なし
'
Public Sub InitGUI(lngWidth As Long _
                , lngHeight As Long _
                , fraTop As Frame _
                , lblMakeDate As Label _
                , fraHeader As Frame _
                , mstTab As SSTab _
                    )
    Dim lngColorDark As Long
    Dim lngColorLight As Long
    
    lngColorDark = gApp.ColDarkBG
    lngColorLight = gApp.ColBG
    
    ' ポジショニング＆リサイズ
    Call ResizeGUI(lngWidth, lngHeight, fraTop, lblMakeDate, fraHeader, mstTab)
    
    With fraTop
        .BackColor = lngColorDark
        .BorderStyle = vbBSNone
    End With 'fraTop
    
    With lblMakeDate
        .BackColor = lngColorLight
        .FontBold = False
        .FontItalic = False
        .FontName = "MS UI Gothic"
        .FontSize = 8
        .FontStrikethru = False
        .FontUnderline = False
        .ForeColor = Contrast(lngColorLight)
        .Alignment = vbRightJustify
    End With
    
    With fraHeader
        .BackColor = lngColorLight
        .BorderStyle = vbBSNone
    End With

    With mstTab
        .BackColor = lngColorLight
        .ZOrder 0
    End With
End Sub



'
'   機能: GUI の共通設定　リサイズ
'
'   備考: なし
'
Public Sub ResizeGUI(lngWidth As Long _
                , lngHeight As Long _
                , fraTop As Frame _
                , lblMakeDate As Label _
                , fraHeader As Frame _
                , mstTab As SSTab _
                    )
    fraTop.Visible = False
    lblMakeDate.Visible = False
    fraHeader.Visible = False
    mstTab.Visible = False
    
    With fraTop
        .Left = 120
        .Top = 180
        .width = Bigger(1, lngWidth - (.Left * 2))
        ' .Height は、手動設定
    End With 'fraTop
    
    With lblMakeDate
        .Top = 30
        .Left = fraTop.Left + fraTop.width - .width
        .AutoSize = True
    End With
    
    With fraHeader
        .Left = fraTop.Left + 120
        .Top = fraTop.Top + fraTop.Height
        .width = Bigger(1, lngWidth - (.Left * 2))
        ' .Height は、手動設定
    End With

    With mstTab
        .Left = fraTop.Left
        .Top = fraHeader.Top + fraHeader.Height
        .width = Bigger(1, lngWidth - (.Left * 2))
        .Height = Bigger(1, lngHeight - .Top - .Left)
    End With
    
    fraTop.Visible = True
    lblMakeDate.Visible = True
    fraHeader.Visible = True
    mstTab.Visible = True

End Sub


'
'   機能: トップフレーム用のフォント設定　タイプ１
'
'   備考: なし
'
Public Sub FraTopFontType1(ByRef target As Font)
    With target
        .Bold = True
        .Italic = False
        .Name = "ＭＳ Ｐゴシック"
        .Size = 12
        .Strikethrough = False
        .Underline = False
        
    End With
End Sub


'
'   機能: FlexGrid の共通設定
'
'   備考: なし
'
Public Sub FlexGridCommonSetting(ByRef target As MSFlexGrid)
    With target
        '行（列）のヘッダのクリックにより行（列）全体を選択できるようにします。
        .AllowBigSelection = False
        'マウスやキーボードを使ってセルを範囲選択できるようにします。
        '.AllowSelection = False
        'マウスを使って静止行（列）数を変更できるかどうかを設定します。
        '.AllowUserFreezing = flexFreezeNone
        'マウスを使って行（列）の大きさを変更できるようにします。
        .AllowUserResizing = flexResizeBoth
        
        'MDIフォームまたはフォームオブジェクト上のコントロールの描画スタイルを設定します。
        .Appearance = flexFlat
        'ユーザーが行ヘッダをダブルクリックすると、列内のデータに合わせて列幅が自動調節されるかどうかを設定します。
        '.AutoSizeMouse = True
        
        'スクロール可能領域の背景色
        .BackColor = &H80000005
        'グリッドの最後の列（最右列）または行（最下行）を越えた通常グレー表示されている領域の背景色
        .BackColorBkg = &H8000000C
        '固定列（行）の背景色
        .BackColorFixed = &H8000000F
        '静止セルの背景色
        '.BackColorFrozen = &H0&
        '選択されたセルのテキスト色
        .BackColorSel = &H80000005
        '１行おきの背景色を設定します。
        '.BackColorAlternate = &H80000005
        
        'コントロールを編集可能にするかどうかを設定します。
        '.Editable = flexEDNone
        'セル内の文字列がすべて表示できないときに、省略符号（...）を表示するようにするかどうかを設定します。
        '.Ellipsis = flexNoEllipsis
        'ヘッダを使ってソートや移動ができるようにするかどうかを設定します。
        '.ExplorerBar = flexExSortShow
        '最後（右端）の列の幅をコントロール幅に合わせて自動的に調節するかどうかを設定します。
        '.ExtendLastCol = False
        'カレントセルの周囲に描画するフォーカス枠のスタイルを設定します。
        .FocusRect = flexFocusNone
        
        'スクロール可能領域のテキスト色
        .ForeColor = &H0
        '　固定列（行）のテキスト色
        .ForeColorFixed = &H0
        ' 静止セルのテキスト色
        '.ForeColorFrozen = &H0
        '選択されたセルのテキスト色
        .ForeColorSel = &H0
        
        .GridColor = &HC0C0C0
        .GridColorFixed = &H0
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .HighLight = flexHighlightNever
        
        'ユーザーがスクロールのつまみを動かしたときに、それに合わせてFlexGridの内容をスクロールするかどうかを設定します。
        .ScrollTrack = True
        
        .Font.Bold = False
        .FontName = "ＭＳ ゴシック"
        End With
End Sub

