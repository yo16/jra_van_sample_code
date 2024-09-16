VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "コード変換クラス"
'
'   コード変換クラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 競馬場コード の 場名 を可読文字列にする
'
'   備考: なし
'
Public Function KIBJ1(strCD As String) As String
    Select Case strCD
    Case "00"
        KIBJ1 = ""
    Case "01"
        KIBJ1 = "札幌競馬場"
    Case "02"
        KIBJ1 = "函館競馬場"
    Case "03"
        KIBJ1 = "福島競馬場"
    Case "04"
        KIBJ1 = "新潟競馬場"
    Case "05"
        KIBJ1 = "東京競馬場"
    Case "06"
        KIBJ1 = "中山競馬場"
    Case "07"
        KIBJ1 = "中京競馬場"
    Case "08"
        KIBJ1 = "京都競馬場"
    Case "09"
        KIBJ1 = "阪神競馬場"
    Case "10"
        KIBJ1 = "小倉競馬場"
    Case "30"
        KIBJ1 = "門別競馬場"
    Case "31"
        KIBJ1 = "北見競馬場"
    Case "32"
        KIBJ1 = "岩見沢競馬場"
    Case "33"
        KIBJ1 = "帯広競馬場"
    Case "34"
        KIBJ1 = "旭川競馬場"
    Case "35"
        KIBJ1 = "盛岡競馬場"
    Case "36"
        KIBJ1 = "水沢競馬場"
    Case "37"
        KIBJ1 = "上山競馬場"
    Case "38"
        KIBJ1 = "三条競馬場"
    Case "39"
        KIBJ1 = "足利競馬場"
    Case "40"
        KIBJ1 = "宇都宮競馬場"
    Case "41"
        KIBJ1 = "高崎競馬場"
    Case "42"
        KIBJ1 = "浦和競馬場"
    Case "43"
        KIBJ1 = "船橋競馬場"
    Case "44"
        KIBJ1 = "大井競馬場"
    Case "45"
        KIBJ1 = "川崎競馬場"
    Case "46"
        KIBJ1 = "金沢競馬場"
    Case "47"
        KIBJ1 = "笠松競馬場"
    Case "48"
        KIBJ1 = "名古屋競馬場"
    Case "49"
        KIBJ1 = "紀三井寺競馬場"
    Case "50"
        KIBJ1 = "園田競馬場"
    Case "51"
        KIBJ1 = "姫路競馬場"
    Case "52"
        KIBJ1 = "益田競馬場"
    Case "53"
        KIBJ1 = "福山競馬場"
    Case "54"
        KIBJ1 = "高知競馬場"
    Case "55"
        KIBJ1 = "佐賀競馬場"
    Case "56"
        KIBJ1 = "荒尾競馬場"
    Case "57"
        KIBJ1 = "中津競馬場"
    Case "58"
        KIBJ1 = "札幌競馬場（地方競馬）"
    Case "59"
        KIBJ1 = "函館競馬場（地方競馬）"
    Case "60"
        KIBJ1 = "新潟競馬場（地方競馬）"
    Case "61"
        KIBJ1 = "中京競馬場（地方競馬）"
    Case "A0"
        KIBJ1 = "その他の外国"
    Case "A2"
        KIBJ1 = "日本"
    Case "A4"
        KIBJ1 = "アメリカ"
    Case "A6"
        KIBJ1 = "イギリス"
    Case "A8"
        KIBJ1 = "フランス"
    Case "B0"
        KIBJ1 = "インド"
    Case "B2"
        KIBJ1 = "アイルランド"
    Case "B4"
        KIBJ1 = "ニュージーランド"
    Case "B6"
        KIBJ1 = "オーストラリア"
    Case "B8"
        KIBJ1 = "カナダ"
    Case "C0"
        KIBJ1 = "イタリア"
    Case "C2"
        KIBJ1 = "ドイツ"
    Case "C4"
        KIBJ1 = "(未使用)"
    Case "C5"
        KIBJ1 = "オマーン"
    Case "C6"
        KIBJ1 = "イラク"
    Case "C7"
        KIBJ1 = "アラブ首長国連邦"
    Case "C8"
        KIBJ1 = "シリア"
    Case "D0"
        KIBJ1 = "スウェーデン"
    Case "D2"
        KIBJ1 = "ハンガリー"
    Case "D4"
        KIBJ1 = "ポルトガル"
    Case "D6"
        KIBJ1 = "ロシア"
    Case "D8"
        KIBJ1 = "ウルグアイ"
    Case "E0"
        KIBJ1 = "ペルー"
    Case "E2"
        KIBJ1 = "アルゼンチン"
    Case "E4"
        KIBJ1 = "ブラジル"
    Case "E6"
        KIBJ1 = "ベルギー"
    Case "E8"
        KIBJ1 = "トルコ"
    Case "F0"
        KIBJ1 = "韓国"
    Case "F1"
        KIBJ1 = "中国"
    Case "F2"
        KIBJ1 = "チリ"
    Case "F4"
        KIBJ1 = "(未使用)"
    Case "F6"
        KIBJ1 = "(未使用)"
    Case "F8"
        KIBJ1 = "パナマ"
    Case "G0"
        KIBJ1 = "香港"
    Case "G2"
        KIBJ1 = "スペイン"
    Case "G4"
        KIBJ1 = "(未使用)"
    Case "G6"
        KIBJ1 = "(未使用)"
    Case "G8"
        KIBJ1 = "(未使用)"
    Case "H0"
        KIBJ1 = "西ドイツ"
    Case "H2"
        KIBJ1 = "南アフリカ"
    Case "H4"
        KIBJ1 = "スイス"
    Case "H6"
        KIBJ1 = "モナコ"
    Case "H8"
        KIBJ1 = "フィリピン"
    Case "I0"
        KIBJ1 = "プエルトリコ"
    Case "I2"
        KIBJ1 = "コロンビア"
    Case "I4"
        KIBJ1 = "チェコスロバキア"
    Case "I6"
        KIBJ1 = "チェコ"
    Case "I8"
        KIBJ1 = "スロバキア"
    Case "J0"
        KIBJ1 = "エクアドル"
    Case "J2"
        KIBJ1 = "ギリシャ"
    Case "J4"
        KIBJ1 = "マレーシア"
    Case "J6"
        KIBJ1 = "メキシコ"
    Case "J8"
        KIBJ1 = "モロッコ"
    Case "K0"
        KIBJ1 = "パキスタン"
    Case "K2"
        KIBJ1 = "ポーランド"
    Case "K4"
        KIBJ1 = "パラグアイ"
    Case "K6"
        KIBJ1 = "サウジアラビア"
    Case "K8"
        KIBJ1 = "キプロス"
    Case "L0"
        KIBJ1 = "タイ"
    Case "L2"
        KIBJ1 = "ウクライナ"
    Case "L4"
        KIBJ1 = "ベネズエラ"
    Case "L6"
        KIBJ1 = "ユーゴスラビア"
    Case "L8"
        KIBJ1 = "デンマーク"
    Case "M0"
        KIBJ1 = "シンガポール"
    Case "M2"
        KIBJ1 = "マカオ"
    End Select
End Function


'
'   機能: 競馬場コード の 場略名(1文字) を可読文字列にする
'
'   備考: なし
'
Public Function KIBJ2(strCD As String) As String
    Select Case strCD
    Case "00"
        KIBJ2 = ""
    Case "01"
        KIBJ2 = "札"
    Case "02"
        KIBJ2 = "函"
    Case "03"
        KIBJ2 = "福"
    Case "04"
        KIBJ2 = "新"
    Case "05"
        KIBJ2 = "東"
    Case "06"
        KIBJ2 = "中"
    Case "07"
        KIBJ2 = "名"
    Case "08"
        KIBJ2 = "京"
    Case "09"
        KIBJ2 = "阪"
    Case "10"
        KIBJ2 = "小"
    Case "30"
        KIBJ2 = "門"
    Case "31"
        KIBJ2 = "北"
    Case "32"
        KIBJ2 = "岩"
    Case "33"
        KIBJ2 = "帯"
    Case "34"
        KIBJ2 = "旭"
    Case "35"
        KIBJ2 = "盛"
    Case "36"
        KIBJ2 = "水"
    Case "37"
        KIBJ2 = "上"
    Case "38"
        KIBJ2 = "三"
    Case "39"
        KIBJ2 = "足"
    Case "40"
        KIBJ2 = "宇"
    Case "41"
        KIBJ2 = "高"
    Case "42"
        KIBJ2 = "浦"
    Case "43"
        KIBJ2 = "船"
    Case "44"
        KIBJ2 = "大"
    Case "45"
        KIBJ2 = "川"
    Case "46"
        KIBJ2 = "金"
    Case "47"
        KIBJ2 = "笠"
    Case "48"
        KIBJ2 = "古"
    Case "49"
        KIBJ2 = "紀"
    Case "50"
        KIBJ2 = "園"
    Case "51"
        KIBJ2 = "姫"
    Case "52"
        KIBJ2 = "益"
    Case "53"
        KIBJ2 = "福"
    Case "54"
        KIBJ2 = "高"
    Case "55"
        KIBJ2 = "佐"
    Case "56"
        KIBJ2 = "荒"
    Case "57"
        KIBJ2 = "中"
    Case "58"
        KIBJ2 = "札"
    Case "59"
        KIBJ2 = "函"
    Case "60"
        KIBJ2 = "新"
    Case "61"
        KIBJ2 = "名"
    Case "A0"
        KIBJ2 = "外"
    Case "A2"
        KIBJ2 = "日"
    Case "A4"
        KIBJ2 = "米"
    Case "A6"
        KIBJ2 = "英"
    Case "A8"
        KIBJ2 = "仏"
    Case "B0"
        KIBJ2 = "印"
    Case "B2"
        KIBJ2 = "愛"
    Case "B4"
        KIBJ2 = "新"
    Case "B6"
        KIBJ2 = "豪"
    Case "B8"
        KIBJ2 = "加"
    Case "C0"
        KIBJ2 = "伊"
    Case "C2"
        KIBJ2 = "独"
    Case "C4"
        KIBJ2 = ""
    Case "C5"
        KIBJ2 = "オ"
    Case "C6"
        KIBJ2 = "イ"
    Case "C7"
        KIBJ2 = "首"
    Case "C8"
        KIBJ2 = "叙"
    Case "D0"
        KIBJ2 = "瑞"
    Case "D2"
        KIBJ2 = "洪"
    Case "D4"
        KIBJ2 = "葡"
    Case "D6"
        KIBJ2 = "露"
    Case "D8"
        KIBJ2 = "宇"
    Case "E0"
        KIBJ2 = "秘"
    Case "E2"
        KIBJ2 = "亜"
    Case "E4"
        KIBJ2 = "伯"
    Case "E6"
        KIBJ2 = "白"
    Case "E8"
        KIBJ2 = "土"
    Case "F0"
        KIBJ2 = "韓"
    Case "F1"
        KIBJ2 = "中"
    Case "F2"
        KIBJ2 = "智"
    Case "F4"
        KIBJ2 = ""
    Case "F6"
        KIBJ2 = ""
    Case "F8"
        KIBJ2 = "巴"
    Case "G0"
        KIBJ2 = "香"
    Case "G2"
        KIBJ2 = "西"
    Case "G4"
        KIBJ2 = ""
    Case "G6"
        KIBJ2 = ""
    Case "G8"
        KIBJ2 = ""
    Case "H0"
        KIBJ2 = "独"
    Case "H2"
        KIBJ2 = "南"
    Case "H4"
        KIBJ2 = "ス"
    Case "H6"
        KIBJ2 = "モ"
    Case "H8"
        KIBJ2 = "比"
    Case "I0"
        KIBJ2 = "プ"
    Case "I2"
        KIBJ2 = "コ"
    Case "I4"
        KIBJ2 = "チ"
    Case "I6"
        KIBJ2 = "チ"
    Case "I8"
        KIBJ2 = "ス"
    Case "J0"
        KIBJ2 = "エ"
    Case "J2"
        KIBJ2 = "ギ"
    Case "J4"
        KIBJ2 = "馬"
    Case "J6"
        KIBJ2 = "墨"
    Case "J8"
        KIBJ2 = "摩"
    Case "K0"
        KIBJ2 = "基"
    Case "K2"
        KIBJ2 = "波"
    Case "K4"
        KIBJ2 = "拉"
    Case "K6"
        KIBJ2 = "サ"
    Case "K8"
        KIBJ2 = "キ"
    Case "L0"
        KIBJ2 = "泰"
    Case "L2"
        KIBJ2 = "烏"
    Case "L4"
        KIBJ2 = "ベ"
    Case "L6"
        KIBJ2 = "ユ"
    Case "L8"
        KIBJ2 = "丁"
    Case "M0"
        KIBJ2 = "嘉"
    Case "M2"
        KIBJ2 = "澳"
    End Select
End Function


'
'   機能: 競馬場コード の 場略名(2文字) を可読文字列にする
'
'   備考: なし
'
Public Function KIBJ3(strCD As String) As String
    Select Case strCD
    Case "00"
        KIBJ3 = ""
    Case "01"
        KIBJ3 = "札幌"
    Case "02"
        KIBJ3 = "函館"
    Case "03"
        KIBJ3 = "福島"
    Case "04"
        KIBJ3 = "新潟"
    Case "05"
        KIBJ3 = "東京"
    Case "06"
        KIBJ3 = "中山"
    Case "07"
        KIBJ3 = "中京"
    Case "08"
        KIBJ3 = "京都"
    Case "09"
        KIBJ3 = "阪神"
    Case "10"
        KIBJ3 = "小倉"
    Case "30"
        KIBJ3 = "門別"
    Case "31"
        KIBJ3 = "北見"
    Case "32"
        KIBJ3 = "岩見"
    Case "33"
        KIBJ3 = "帯広"
    Case "34"
        KIBJ3 = "旭川"
    Case "35"
        KIBJ3 = "盛岡"
    Case "36"
        KIBJ3 = "水沢"
    Case "37"
        KIBJ3 = "上山"
    Case "38"
        KIBJ3 = "三条"
    Case "39"
        KIBJ3 = "足利"
    Case "40"
        KIBJ3 = "宇都"
    Case "41"
        KIBJ3 = "高崎"
    Case "42"
        KIBJ3 = "浦和"
    Case "43"
        KIBJ3 = "船橋"
    Case "44"
        KIBJ3 = "大井"
    Case "45"
        KIBJ3 = "川崎"
    Case "46"
        KIBJ3 = "金沢"
    Case "47"
        KIBJ3 = "笠松"
    Case "48"
        KIBJ3 = "名古"
    Case "49"
        KIBJ3 = "紀三"
    Case "50"
        KIBJ3 = "園田"
    Case "51"
        KIBJ3 = "姫路"
    Case "52"
        KIBJ3 = "益田"
    Case "53"
        KIBJ3 = "福山"
    Case "54"
        KIBJ3 = "高知"
    Case "55"
        KIBJ3 = "佐賀"
    Case "56"
        KIBJ3 = "荒尾"
    Case "57"
        KIBJ3 = "中津"
    Case "58"
        KIBJ3 = "N-札"
    Case "59"
        KIBJ3 = "N-函"
    Case "60"
        KIBJ3 = "N-新"
    Case "61"
        KIBJ3 = "N-中"
    Case "A0"
        KIBJ3 = "他外"
    Case "A2"
        KIBJ3 = "日本"
    Case "A4"
        KIBJ3 = "アメ"
    Case "A6"
        KIBJ3 = "イギ"
    Case "A8"
        KIBJ3 = "フラ"
    Case "B0"
        KIBJ3 = "イン"
    Case "B2"
        KIBJ3 = "アイ"
    Case "B4"
        KIBJ3 = "ニュ"
    Case "B6"
        KIBJ3 = "オー"
    Case "B8"
        KIBJ3 = "カナ"
    Case "C0"
        KIBJ3 = "イタ"
    Case "C2"
        KIBJ3 = "ドイ"
    Case "C4"
        KIBJ3 = ""
    Case "C5"
        KIBJ3 = "オマ"
    Case "C6"
        KIBJ3 = "イラ"
    Case "C7"
        KIBJ3 = "ア首"
    Case "C8"
        KIBJ3 = "シリ"
    Case "D0"
        KIBJ3 = "スウ"
    Case "D2"
        KIBJ3 = "ハン"
    Case "D4"
        KIBJ3 = "ポル"
    Case "D6"
        KIBJ3 = "ロシ"
    Case "D8"
        KIBJ3 = "ウル"
    Case "E0"
        KIBJ3 = "ペル"
    Case "E2"
        KIBJ3 = "アル"
    Case "E4"
        KIBJ3 = "ブラ"
    Case "E6"
        KIBJ3 = "ベル"
    Case "E8"
        KIBJ3 = "トル"
    Case "F0"
        KIBJ3 = "韓国"
    Case "F1"
        KIBJ3 = "中国"
    Case "F2"
        KIBJ3 = "チリ"
    Case "F4"
        KIBJ3 = ""
    Case "F6"
        KIBJ3 = ""
    Case "F8"
        KIBJ3 = "パナ"
    Case "G0"
        KIBJ3 = "香港"
    Case "G2"
        KIBJ3 = "スペ"
    Case "G4"
        KIBJ3 = ""
    Case "G6"
        KIBJ3 = ""
    Case "G8"
        KIBJ3 = ""
    Case "H0"
        KIBJ3 = "西独"
    Case "H2"
        KIBJ3 = "南ア"
    Case "H4"
        KIBJ3 = "スイ"
    Case "H6"
        KIBJ3 = "モナ"
    Case "H8"
        KIBJ3 = "フィ"
    Case "I0"
        KIBJ3 = "プエ"
    Case "I2"
        KIBJ3 = "コロ"
    Case "I4"
        KIBJ3 = "チェ"
    Case "I6"
        KIBJ3 = "チェ"
    Case "I8"
        KIBJ3 = "スロ"
    Case "J0"
        KIBJ3 = "エク"
    Case "J2"
        KIBJ3 = "ギリ"
    Case "J4"
        KIBJ3 = "マレ"
    Case "J6"
        KIBJ3 = "メキ"
    Case "J8"
        KIBJ3 = "モロ"
    Case "K0"
        KIBJ3 = "パキ"
    Case "K2"
        KIBJ3 = "ポー"
    Case "K4"
        KIBJ3 = "パラ"
    Case "K6"
        KIBJ3 = "サウ"
    Case "K8"
        KIBJ3 = "キプ"
    Case "L0"
        KIBJ3 = "タイ"
    Case "L2"
        KIBJ3 = "ウク"
    Case "L4"
        KIBJ3 = "ベネ"
    Case "L6"
        KIBJ3 = "ユー"
    Case "L8"
        KIBJ3 = "デン"
    Case "M0"
        KIBJ3 = "シン"
    Case "M2"
        KIBJ3 = "マカ"
    End Select
End Function


'
'   機能: 競馬場コード の 場略名(3文字) を可読文字列にする
'
'   備考: なし
'
Public Function KIBJ4(strCD As String) As String
    Select Case strCD
    Case "00"
        KIBJ4 = ""
    Case "01"
        KIBJ4 = "札幌"
    Case "02"
        KIBJ4 = "函館"
    Case "03"
        KIBJ4 = "福島"
    Case "04"
        KIBJ4 = "新潟"
    Case "05"
        KIBJ4 = "東京"
    Case "06"
        KIBJ4 = "中山"
    Case "07"
        KIBJ4 = "中京"
    Case "08"
        KIBJ4 = "京都"
    Case "09"
        KIBJ4 = "阪神"
    Case "10"
        KIBJ4 = "小倉"
    Case "30"
        KIBJ4 = "門別"
    Case "31"
        KIBJ4 = "北見"
    Case "32"
        KIBJ4 = "岩見沢"
    Case "33"
        KIBJ4 = "帯広"
    Case "34"
        KIBJ4 = "旭川"
    Case "35"
        KIBJ4 = "盛岡"
    Case "36"
        KIBJ4 = "水沢"
    Case "37"
        KIBJ4 = "上山"
    Case "38"
        KIBJ4 = "三条"
    Case "39"
        KIBJ4 = "足利"
    Case "40"
        KIBJ4 = "宇都宮"
    Case "41"
        KIBJ4 = "高崎"
    Case "42"
        KIBJ4 = "浦和"
    Case "43"
        KIBJ4 = "船橋"
    Case "44"
        KIBJ4 = "大井"
    Case "45"
        KIBJ4 = "川崎"
    Case "46"
        KIBJ4 = "金沢"
    Case "47"
        KIBJ4 = "笠松"
    Case "48"
        KIBJ4 = "名古屋"
    Case "49"
        KIBJ4 = "紀三寺"
    Case "50"
        KIBJ4 = "園田"
    Case "51"
        KIBJ4 = "姫路"
    Case "52"
        KIBJ4 = "益田"
    Case "53"
        KIBJ4 = "福山"
    Case "54"
        KIBJ4 = "高知"
    Case "55"
        KIBJ4 = "佐賀"
    Case "56"
        KIBJ4 = "荒尾"
    Case "57"
        KIBJ4 = "中津"
    Case "58"
        KIBJ4 = "N-札幌"
    Case "59"
        KIBJ4 = "N-函館"
    Case "60"
        KIBJ4 = "N-新潟"
    Case "61"
        KIBJ4 = "N-中京"
    Case "A0"
        KIBJ4 = "他外国"
    Case "A2"
        KIBJ4 = "日本"
    Case "A4"
        KIBJ4 = "アメリ"
    Case "A6"
        KIBJ4 = "イギリ"
    Case "A8"
        KIBJ4 = "フラン"
    Case "B0"
        KIBJ4 = "インド"
    Case "B2"
        KIBJ4 = "アイル"
    Case "B4"
        KIBJ4 = "ニュー"
    Case "B6"
        KIBJ4 = "オース"
    Case "B8"
        KIBJ4 = "カナダ"
    Case "C0"
        KIBJ4 = "イタリ"
    Case "C2"
        KIBJ4 = "ドイツ"
    Case "C4"
        KIBJ4 = ""
    Case "C5"
        KIBJ4 = "オマー"
    Case "C6"
        KIBJ4 = "イラク"
    Case "C7"
        KIBJ4 = "アラブ"
    Case "C8"
        KIBJ4 = "シリア"
    Case "D0"
        KIBJ4 = "スウェ"
    Case "D2"
        KIBJ4 = "ハンガ"
    Case "D4"
        KIBJ4 = "ポルト"
    Case "D6"
        KIBJ4 = "ロシア"
    Case "D8"
        KIBJ4 = "ウルグ"
    Case "E0"
        KIBJ4 = "ペルー"
    Case "E2"
        KIBJ4 = "アルゼ"
    Case "E4"
        KIBJ4 = "ブラジ"
    Case "E6"
        KIBJ4 = "ベルギ"
    Case "E8"
        KIBJ4 = "トルコ"
    Case "F0"
        KIBJ4 = "韓国"
    Case "F1"
        KIBJ4 = "中国"
    Case "F2"
        KIBJ4 = "チリ"
    Case "F4"
        KIBJ4 = ""
    Case "F6"
        KIBJ4 = ""
    Case "F8"
        KIBJ4 = "パナマ"
    Case "G0"
        KIBJ4 = "香港"
    Case "G2"
        KIBJ4 = "スペイ"
    Case "G4"
        KIBJ4 = ""
    Case "G6"
        KIBJ4 = ""
    Case "G8"
        KIBJ4 = ""
    Case "H0"
        KIBJ4 = "西独"
    Case "H2"
        KIBJ4 = "南アフ"
    Case "H4"
        KIBJ4 = "スイス"
    Case "H6"
        KIBJ4 = "モナコ"
    Case "H8"
        KIBJ4 = "フィリ"
    Case "I0"
        KIBJ4 = "プエル"
    Case "I2"
        KIBJ4 = "コロン"
    Case "I4"
        KIBJ4 = "チェコ"
    Case "I6"
        KIBJ4 = "チェコ"
    Case "I8"
        KIBJ4 = "スロバ"
    Case "J0"
        KIBJ4 = "エクア"
    Case "J2"
        KIBJ4 = "ギリシ"
    Case "J4"
        KIBJ4 = "マレー"
    Case "J6"
        KIBJ4 = "メキシ"
    Case "J8"
        KIBJ4 = "モロッ"
    Case "K0"
        KIBJ4 = "パキス"
    Case "K2"
        KIBJ4 = "ポーラ"
    Case "K4"
        KIBJ4 = "パラグ"
    Case "K6"
        KIBJ4 = "サウジ"
    Case "K8"
        KIBJ4 = "キプロ"
    Case "L0"
        KIBJ4 = "タイ"
    Case "L2"
        KIBJ4 = "ウクラ"
    Case "L4"
        KIBJ4 = "ベネゼ"
    Case "L6"
        KIBJ4 = "ユーゴ"
    Case "L8"
        KIBJ4 = "デンマ"
    Case "M0"
        KIBJ4 = "シンガ"
    Case "M2"
        KIBJ4 = "マカオ"
    End Select
End Function


'
'   機能: 競馬場コード の 場欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function KIBJ5(strCD As String) As String
    Select Case strCD
    Case "00"
        KIBJ5 = ""
    Case "01"
        KIBJ5 = "SAPPORO"
    Case "02"
        KIBJ5 = "HAKODATE"
    Case "03"
        KIBJ5 = "FUKUSHIMA"
    Case "04"
        KIBJ5 = "NIIGATA"
    Case "05"
        KIBJ5 = "TOKYO"
    Case "06"
        KIBJ5 = "NAKAYAMA"
    Case "07"
        KIBJ5 = "CHUKYO"
    Case "08"
        KIBJ5 = "KYOTO"
    Case "09"
        KIBJ5 = "HANSHIN"
    Case "10"
        KIBJ5 = "KOKURA"
    Case "30"
        KIBJ5 = "MONBETSU"
    Case "31"
        KIBJ5 = "KITAMI"
    Case "32"
        KIBJ5 = "IWAMIZAWA"
    Case "33"
        KIBJ5 = "OBIHIRO"
    Case "34"
        KIBJ5 = "ASAHIKAWA"
    Case "35"
        KIBJ5 = "MORIOKA"
    Case "36"
        KIBJ5 = "MIZUSAWA"
    Case "37"
        KIBJ5 = "KAMINOYAMA"
    Case "38"
        KIBJ5 = "SANJYO"
    Case "39"
        KIBJ5 = "ASHIKAGA"
    Case "40"
        KIBJ5 = "UTSUNOMIYA"
    Case "41"
        KIBJ5 = "TAKASAKI"
    Case "42"
        KIBJ5 = "URAWA"
    Case "43"
        KIBJ5 = "FUNABASHI"
    Case "44"
        KIBJ5 = "OOI"
    Case "45"
        KIBJ5 = "KAWASAKI"
    Case "46"
        KIBJ5 = "KANAZAWA"
    Case "47"
        KIBJ5 = "KASAMATSU"
    Case "48"
        KIBJ5 = "NAGOYA"
    Case "49"
        KIBJ5 = "KIMIIDERA"
    Case "50"
        KIBJ5 = "SONODA"
    Case "51"
        KIBJ5 = "HIMEJI"
    Case "52"
        KIBJ5 = "MASUDA"
    Case "53"
        KIBJ5 = "FUKUYAMA"
    Case "54"
        KIBJ5 = "KOCHI"
    Case "55"
        KIBJ5 = "SAGA"
    Case "56"
        KIBJ5 = "ARAO"
    Case "57"
        KIBJ5 = "NAKATSU"
    Case "58"
        KIBJ5 = "SAPPORO(NAR)"
    Case "59"
        KIBJ5 = "HAKODATE(NAR)"
    Case "60"
        KIBJ5 = "NIIGATA(NAR)"
    Case "61"
        KIBJ5 = "CHUKYO(NAR)"
    Case "A0"
        KIBJ5 = ""
    Case "A2"
        KIBJ5 = "Japan"
    Case "A4"
        KIBJ5 = "United States of America"
    Case "A6"
        KIBJ5 = "Great Britain"
    Case "A8"
        KIBJ5 = "France"
    Case "B0"
        KIBJ5 = "India"
    Case "B2"
        KIBJ5 = "Ireland"
    Case "B4"
        KIBJ5 = "New Zealand"
    Case "B6"
        KIBJ5 = "Australia"
    Case "B8"
        KIBJ5 = "Canada"
    Case "C0"
        KIBJ5 = "Italy"
    Case "C2"
        KIBJ5 = "Germany"
    Case "C4"
        KIBJ5 = ""
    Case "C5"
        KIBJ5 = "Oman"
    Case "C6"
        KIBJ5 = "Iraq"
    Case "C7"
        KIBJ5 = "United Arab Emirates"
    Case "C8"
        KIBJ5 = "Syrian"
    Case "D0"
        KIBJ5 = "Sweden"
    Case "D2"
        KIBJ5 = "Hungary"
    Case "D4"
        KIBJ5 = "Portugal"
    Case "D6"
        KIBJ5 = "Russia"
    Case "D8"
        KIBJ5 = "Uruguay"
    Case "E0"
        KIBJ5 = "Peru"
    Case "E2"
        KIBJ5 = "Argentina"
    Case "E4"
        KIBJ5 = "Brazil"
    Case "E6"
        KIBJ5 = "Belgium"
    Case "E8"
        KIBJ5 = "Turkey"
    Case "F0"
        KIBJ5 = "Korea"
    Case "F1"
        KIBJ5 = "China"
    Case "F2"
        KIBJ5 = "Chile"
    Case "F4"
        KIBJ5 = ""
    Case "F6"
        KIBJ5 = ""
    Case "F8"
        KIBJ5 = "Panama"
    Case "G0"
        KIBJ5 = "Hong Kong"
    Case "G2"
        KIBJ5 = "Spain"
    Case "G4"
        KIBJ5 = ""
    Case "G6"
        KIBJ5 = ""
    Case "G8"
        KIBJ5 = ""
    Case "H0"
        KIBJ5 = "West Germany"
    Case "H2"
        KIBJ5 = "South Africa"
    Case "H4"
        KIBJ5 = "Switzerland"
    Case "H6"
        KIBJ5 = "Monaco"
    Case "H8"
        KIBJ5 = "Philippines"
    Case "I0"
        KIBJ5 = "Puerto Rico"
    Case "I2"
        KIBJ5 = "Colombia"
    Case "I4"
        KIBJ5 = "Czechoslovakia"
    Case "I6"
        KIBJ5 = "Czech Republic"
    Case "I8"
        KIBJ5 = "Slovakia"
    Case "J0"
        KIBJ5 = "Ecuador"
    Case "J2"
        KIBJ5 = "Greece"
    Case "J4"
        KIBJ5 = "Malaysia"
    Case "J6"
        KIBJ5 = "Mexico"
    Case "J8"
        KIBJ5 = "Morocco"
    Case "K0"
        KIBJ5 = "Pakistan"
    Case "K2"
        KIBJ5 = "Poland"
    Case "K4"
        KIBJ5 = "Paraguay"
    Case "K6"
        KIBJ5 = "Saudi Arabia"
    Case "K8"
        KIBJ5 = "Cyprus"
    Case "L0"
        KIBJ5 = "Thailand"
    Case "L2"
        KIBJ5 = "Ukraine"
    Case "L4"
        KIBJ5 = "Venezuela"
    Case "L6"
        KIBJ5 = "Yugoslavia"
    Case "L8"
        KIBJ5 = "Denmark"
    Case "M0"
        KIBJ5 = "Singapore"
    Case "M2"
        KIBJ5 = "Macau"
    End Select
End Function


'
'   機能: 競馬場コード の 場欧字略名(3文字) を可読文字列にする
'
'   備考: なし
'
Public Function KIBJ6(strCD As String) As String
    Select Case strCD
    Case "00"
        KIBJ6 = ""
    Case "01"
        KIBJ6 = ""
    Case "02"
        KIBJ6 = ""
    Case "03"
        KIBJ6 = ""
    Case "04"
        KIBJ6 = ""
    Case "05"
        KIBJ6 = ""
    Case "06"
        KIBJ6 = ""
    Case "07"
        KIBJ6 = ""
    Case "08"
        KIBJ6 = ""
    Case "09"
        KIBJ6 = ""
    Case "10"
        KIBJ6 = ""
    Case "30"
        KIBJ6 = ""
    Case "31"
        KIBJ6 = ""
    Case "32"
        KIBJ6 = ""
    Case "33"
        KIBJ6 = ""
    Case "34"
        KIBJ6 = ""
    Case "35"
        KIBJ6 = ""
    Case "36"
        KIBJ6 = ""
    Case "37"
        KIBJ6 = ""
    Case "38"
        KIBJ6 = ""
    Case "39"
        KIBJ6 = ""
    Case "40"
        KIBJ6 = ""
    Case "41"
        KIBJ6 = ""
    Case "42"
        KIBJ6 = ""
    Case "43"
        KIBJ6 = ""
    Case "44"
        KIBJ6 = ""
    Case "45"
        KIBJ6 = ""
    Case "46"
        KIBJ6 = ""
    Case "47"
        KIBJ6 = ""
    Case "48"
        KIBJ6 = ""
    Case "49"
        KIBJ6 = ""
    Case "50"
        KIBJ6 = ""
    Case "51"
        KIBJ6 = ""
    Case "52"
        KIBJ6 = ""
    Case "53"
        KIBJ6 = ""
    Case "54"
        KIBJ6 = ""
    Case "55"
        KIBJ6 = ""
    Case "56"
        KIBJ6 = ""
    Case "57"
        KIBJ6 = ""
    Case "58"
        KIBJ6 = ""
    Case "59"
        KIBJ6 = ""
    Case "60"
        KIBJ6 = ""
    Case "61"
        KIBJ6 = ""
    Case "A0"
        KIBJ6 = ""
    Case "A2"
        KIBJ6 = "JPN"
    Case "A4"
        KIBJ6 = "USA"
    Case "A6"
        KIBJ6 = "GB"
    Case "A8"
        KIBJ6 = "FR"
    Case "B0"
        KIBJ6 = "IND"
    Case "B2"
        KIBJ6 = "IRE"
    Case "B4"
        KIBJ6 = "NZ"
    Case "B6"
        KIBJ6 = "AUS"
    Case "B8"
        KIBJ6 = "CAN"
    Case "C0"
        KIBJ6 = "ITY"
    Case "C2"
        KIBJ6 = "GER"
    Case "C4"
        KIBJ6 = ""
    Case "C5"
        KIBJ6 = "OMN"
    Case "C6"
        KIBJ6 = "IRQ"
    Case "C7"
        KIBJ6 = "UAE"
    Case "C8"
        KIBJ6 = "SYR"
    Case "D0"
        KIBJ6 = "SWE"
    Case "D2"
        KIBJ6 = "HUN"
    Case "D4"
        KIBJ6 = "POR"
    Case "D6"
        KIBJ6 = "RUS"
    Case "D8"
        KIBJ6 = "URU"
    Case "E0"
        KIBJ6 = "PER"
    Case "E2"
        KIBJ6 = "ARG"
    Case "E4"
        KIBJ6 = "BRZ"
    Case "E6"
        KIBJ6 = "BEL"
    Case "E8"
        KIBJ6 = "TUR"
    Case "F0"
        KIBJ6 = "KOR"
    Case "F1"
        KIBJ6 = "CHN"
    Case "F2"
        KIBJ6 = "CHI"
    Case "F4"
        KIBJ6 = ""
    Case "F6"
        KIBJ6 = ""
    Case "F8"
        KIBJ6 = "PAN"
    Case "G0"
        KIBJ6 = "HK"
    Case "G2"
        KIBJ6 = "SPA"
    Case "G4"
        KIBJ6 = ""
    Case "G6"
        KIBJ6 = ""
    Case "G8"
        KIBJ6 = ""
    Case "H0"
        KIBJ6 = "GER"
    Case "H2"
        KIBJ6 = "SAF"
    Case "H4"
        KIBJ6 = "SWI"
    Case "H6"
        KIBJ6 = "MCO"
    Case "H8"
        KIBJ6 = "PHI"
    Case "I0"
        KIBJ6 = "PUE"
    Case "I2"
        KIBJ6 = "COL"
    Case "I4"
        KIBJ6 = "CZE"
    Case "I6"
        KIBJ6 = "CZE"
    Case "I8"
        KIBJ6 = "SLO"
    Case "J0"
        KIBJ6 = "ECU"
    Case "J2"
        KIBJ6 = "GR"
    Case "J4"
        KIBJ6 = "MAL"
    Case "J6"
        KIBJ6 = "MEX"
    Case "J8"
        KIBJ6 = "MOR"
    Case "K0"
        KIBJ6 = "PAK"
    Case "K2"
        KIBJ6 = "POL"
    Case "K4"
        KIBJ6 = "PRY"
    Case "K6"
        KIBJ6 = "SDA"
    Case "K8"
        KIBJ6 = "CYP"
    Case "L0"
        KIBJ6 = "THA"
    Case "L2"
        KIBJ6 = "UKR"
    Case "L4"
        KIBJ6 = "VEN"
    Case "L6"
        KIBJ6 = "YUG"
    Case "L8"
        KIBJ6 = "DEN"
    Case "M0"
        KIBJ6 = "SIN"
    Case "M2"
        KIBJ6 = "MAC"
    End Select
End Function


'
'   機能: 曜日コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function YOBI1(strCD As String) As String
    Select Case strCD
    Case "0"
        YOBI1 = ""
    Case "1"
        YOBI1 = "土曜日"
    Case "2"
        YOBI1 = "日曜日"
    Case "3"
        YOBI1 = "祝日"
    Case "4"
        YOBI1 = "月曜日"
    Case "5"
        YOBI1 = "火曜日"
    Case "6"
        YOBI1 = "水曜日"
    Case "7"
        YOBI1 = "木曜日"
    Case "8"
        YOBI1 = "金曜日"
    End Select
End Function


'
'   機能: 曜日コード の 略名(1文字) を可読文字列にする
'
'   備考: なし
'
Public Function YOBI2(strCD As String) As String
    Select Case strCD
    Case "0"
        YOBI2 = ""
    Case "1"
        YOBI2 = "土"
    Case "2"
        YOBI2 = "日"
    Case "3"
        YOBI2 = "祝"
    Case "4"
        YOBI2 = "月"
    Case "5"
        YOBI2 = "火"
    Case "6"
        YOBI2 = "水"
    Case "7"
        YOBI2 = "木"
    Case "8"
        YOBI2 = "金"
    End Select
End Function


'
'   機能: 曜日コード の 略名(2文字) を可読文字列にする
'
'   備考: なし
'
Public Function YOBI3(strCD As String) As String
    Select Case strCD
    Case "0"
        YOBI3 = ""
    Case "1"
        YOBI3 = "土曜"
    Case "2"
        YOBI3 = "日曜"
    Case "3"
        YOBI3 = "祝日"
    Case "4"
        YOBI3 = "月曜"
    Case "5"
        YOBI3 = "火曜"
    Case "6"
        YOBI3 = "水曜"
    Case "7"
        YOBI3 = "木曜"
    Case "8"
        YOBI3 = "金曜"
    End Select
End Function


'
'   機能: 曜日コード の 略名(3文字) を可読文字列にする
'
'   備考: なし
'
Public Function YOBI4(strCD As String) As String
    Select Case strCD
    Case "0"
        YOBI4 = ""
    Case "1"
        YOBI4 = "土曜日"
    Case "2"
        YOBI4 = "日曜日"
    Case "3"
        YOBI4 = "祝日"
    Case "4"
        YOBI4 = "月曜日"
    Case "5"
        YOBI4 = "火曜日"
    Case "6"
        YOBI4 = "水曜日"
    Case "7"
        YOBI4 = "木曜日"
    Case "8"
        YOBI4 = "金曜日"
    End Select
End Function


'
'   機能: 曜日コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function YOBI5(strCD As String) As String
    Select Case strCD
    Case "0"
        YOBI5 = ""
    Case "1"
        YOBI5 = "SATURDAY"
    Case "2"
        YOBI5 = "SUNDAY"
    Case "3"
        YOBI5 = "NATIONAL HOLIDAY"
    Case "4"
        YOBI5 = "MONDAY"
    Case "5"
        YOBI5 = "TUESDAY"
    Case "6"
        YOBI5 = "WEDNESDAY"
    Case "7"
        YOBI5 = "THURSDAY"
    Case "8"
        YOBI5 = "FRIDAY"
    End Select
End Function


'
'   機能: 曜日コード の 欧字略名(3文字) を可読文字列にする
'
'   備考: なし
'
Public Function YOBI6(strCD As String) As String
    Select Case strCD
    Case "0"
        YOBI6 = ""
    Case "1"
        YOBI6 = "SAT"
    Case "2"
        YOBI6 = "SUN"
    Case "3"
        YOBI6 = "HOL"
    Case "4"
        YOBI6 = "MON"
    Case "5"
        YOBI6 = "TUE"
    Case "6"
        YOBI6 = "WED"
    Case "7"
        YOBI6 = "THU"
    Case "8"
        YOBI6 = "FRI"
    End Select
End Function


'
'   機能: 曜日コード をカッコ付の可読文字列にする
'
'   備考: なし
'
Public Function YOBI7(strCD As String) As String
    Select Case strCD
    Case "0"
        YOBI7 = ""
    Case "1"
        YOBI7 = "(土)"
    Case "2"
        YOBI7 = "(日)"
    Case "3"
        YOBI7 = "(祝)"
    Case "4"
        YOBI7 = "(月)"
    Case "5"
        YOBI7 = "(火)"
    Case "6"
        YOBI7 = "(水)"
    Case "7"
        YOBI7 = "(木)"
    Case "8"
        YOBI7 = "(金)"
    End Select
End Function


'
'   機能: グレードコード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function GRAD1(strCD As String) As String
    Select Case strCD
    Case "A"
        GRAD1 = "G1（平地競走）"
    Case "B"
        GRAD1 = "G2（平地競走）"
    Case "C"
        GRAD1 = "G3（平地競走）"
    Case "D"
        GRAD1 = "グレードのない重賞"
    Case "E"
        GRAD1 = "重賞以外の特別競走"
    Case "F"
        GRAD1 = "J･G1（障害競走）"
    Case "G"
        GRAD1 = "J･G2（障害競走）"
    Case "H"
        GRAD1 = "J･G3（障害競走）"
    Case " "
        GRAD1 = ""
    End Select
End Function


'
'   機能: グレードコード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function GRAD2(strCD As String) As String
    Select Case strCD
    Case "A"
        GRAD2 = "(ＧⅠ)"
    Case "B"
        GRAD2 = "(ＧⅡ)"
    Case "C"
        GRAD2 = "(ＧⅢ)"
    Case "D"
        GRAD2 = ""
    Case "E"
        GRAD2 = ""
    Case "F"
        GRAD2 = "(J･ＧⅠ)"
    Case "G"
        GRAD2 = "(J･ＧⅡ)"
    Case "H"
        GRAD2 = "(J･ＧⅢ)"
    Case " "
        GRAD2 = ""
    End Select
End Function

'
'   機能: グレードコード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function GRAD3(strCD As String) As String
    Select Case strCD
    Case "A"
        GRAD3 = "(ＧⅠ)"
    Case "B"
        GRAD3 = "(ＧⅡ)"
    Case "C"
        GRAD3 = "(ＧⅢ)"
    Case "D"
        GRAD3 = ""
    Case "E"
        GRAD3 = ""
    Case "F"
        GRAD3 = "(JGⅠ)"
    Case "G"
        GRAD3 = "(JGⅡ)"
    Case "H"
        GRAD3 = "(JGⅢ)"
    Case " "
        GRAD3 = ""
    End Select
End Function


'
'   機能: グレードコード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function GRAD4(strCD As String) As String
    Select Case strCD
    Case "A"
        GRAD4 = " ＧⅠ "
    Case "B"
        GRAD4 = " ＧⅡ "
    Case "C"
        GRAD4 = " ＧⅢ "
    Case "D"
        GRAD4 = ""
    Case "E"
        GRAD4 = ""
    Case "F"
        GRAD4 = " JGⅠ "
    Case "G"
        GRAD4 = " JGⅡ "
    Case "H"
        GRAD4 = " JGⅢ "
    Case " "
        GRAD4 = ""
    End Select
End Function


'
'   機能: 競走種別コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function KSSB1(strCD As String) As String
    Select Case strCD
    Case "0"
        KSSB1 = ""
    Case "11"
        KSSB1 = "サラブレッド系2歳"
    Case "12"
        KSSB1 = "サラブレッド系3歳"
    Case "13"
        KSSB1 = "サラブレッド系3歳以上"
    Case "14"
        KSSB1 = "サラブレッド系4歳以上"
    Case "18"
        KSSB1 = "サラブレッド系障害3歳以上"
    Case "19"
        KSSB1 = "サラブレッド系障害4歳以上"
    Case "21"
        KSSB1 = "アラブ系2歳"
    Case "22"
        KSSB1 = "アラブ系3歳"
    Case "23"
        KSSB1 = "アラブ系3歳以上"
    Case "24"
        KSSB1 = "アラブ系4歳以上"
    End Select
End Function


'
'   機能: 競走種別コード の 略名(4文字) を可読文字列にする
'
'   備考: なし
'
Public Function KSSB2(strCD As String) As String
    Select Case strCD
    Case "0"
        KSSB2 = ""
    Case "11"
        KSSB2 = "サラ２才"
    Case "12"
        KSSB2 = "サラ３才"
    Case "13"
        KSSB2 = "サラ３上"
    Case "14"
        KSSB2 = "サラ４上"
    Case "18"
        KSSB2 = "障害３上"
    Case "19"
        KSSB2 = "障害４上"
    Case "21"
        KSSB2 = "アラ２才"
    Case "22"
        KSSB2 = "アラ３才"
    Case "23"
        KSSB2 = "アラ３上"
    Case "24"
        KSSB2 = "アラ４上"
    End Select
End Function


'
'   機能: 競走種別コード の 略名(6文字) を可読文字列にする
'
'   備考: なし
'
Public Function KSSB3(strCD As String) As String
    Select Case strCD
    Case "0"
        KSSB3 = ""
    Case "11"
        KSSB3 = "サラ系２歳"
    Case "12"
        KSSB3 = "サラ系３歳"
    Case "13"
        KSSB3 = "サラ系３歳上"
    Case "14"
        KSSB3 = "サラ系４歳上"
    Case "18"
        KSSB3 = "障害３歳上"
    Case "19"
        KSSB3 = "障害４歳上"
    Case "21"
        KSSB3 = "アラ系２歳"
    Case "22"
        KSSB3 = "アラ系３歳"
    Case "23"
        KSSB3 = "アラ系３歳上"
    Case "24"
        KSSB3 = "アラ系４歳上"
    End Select
End Function


'
'   機能: 競走種別コード の 略名(8文字) を可読文字列にする
'
'   備考: なし
'
Public Function KSSB4(strCD As String) As String
    Select Case strCD
    Case "0"
        KSSB4 = ""
    Case "11"
        KSSB4 = "サラ系２歳"
    Case "12"
        KSSB4 = "サラ系３歳"
    Case "13"
        KSSB4 = "サラ系３歳以上"
    Case "14"
        KSSB4 = "サラ系４歳以上"
    Case "18"
        KSSB4 = "サラ障害３歳以上"
    Case "19"
        KSSB4 = "サラ障害４歳以上"
    Case "21"
        KSSB4 = "アラブ系２歳"
    Case "22"
        KSSB4 = "アラブ系３歳"
    Case "23"
        KSSB4 = "アラブ系３歳以上"
    Case "24"
        KSSB4 = "アラブ系４歳以上"
    End Select
End Function


'
'   機能: 競走種別コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function KSSB5(strCD As String) As String
    Select Case strCD
    Case "0"
        KSSB5 = ""
    Case "11"
        KSSB5 = "TWO-YEAR-OLDS"
    Case "12"
        KSSB5 = "THREE-YEAR-OLDS"
    Case "13"
        KSSB5 = "THREE-YEAR-OLDS & UP"
    Case "14"
        KSSB5 = "FOUR-YEAR-OLDS & UP"
    Case "18"
        KSSB5 = "THREE-YEAR-OLDS & UP STEEPLE-CHASE"
    Case "19"
        KSSB5 = "FOUR-YEAR-OLDS & UP STEEPLE-CHASE"
    Case "21"
        KSSB5 = "TWO-YEAR-OLDS ANGLO-ARABS"
    Case "22"
        KSSB5 = "THREE-YEAR-OLDS ANGLO-ARABS"
    Case "23"
        KSSB5 = "THREE-YEAR-OLDS & UP ANGLO-ARABS"
    Case "24"
        KSSB5 = "FOUR-YEAR-OLDS & UP ANGLO-ARABS"
    End Select
End Function


'
'   機能: 競走種別コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function KSSB6(strCD As String) As String
    Select Case strCD
    Case "0"
        KSSB6 = ""
    Case "11"
        KSSB6 = "２歳" '"サラブレッド系2歳"
    Case "12"
        KSSB6 = "３歳" '"サラブレッド系3歳"
    Case "13"
        KSSB6 = "３歳上" '"サラブレッド系3歳以上"
    Case "14"
        KSSB6 = "４歳上" '"サラブレッド系4歳以上"
    Case "18"
        KSSB6 = "３歳上" '"サラブレッド系障害3歳以上"
    Case "19"
        KSSB6 = "４歳上" '"サラブレッド系障害4歳以上"
    Case "21"
        KSSB6 = "２歳" '"アラブ系2歳"
    Case "22"
        KSSB6 = "３歳" '"アラブ系3歳"
    Case "23"
        KSSB6 = "３歳上" '"アラブ系3歳以上"
    Case "24"
        KSSB6 = "４歳上" '"アラブ系4歳以上"
    End Select
End Function


'
'   機能: 競走記号コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function KSKG1(strCD As String, Year As String) As String
    Select Case strCD
    Case "000"
        KSKG1 = ""
    Case "001"
        KSKG1 = "(指定)"
    Case "002"
        If Year < 2004 Then
            KSKG1 = "見習騎手"
        Else
            KSKG1 = "若手騎手"
        End If
    Case "003"
        KSKG1 = "[指定]"
    Case "004"
        KSKG1 = "(特指)"
    Case "020"
        KSKG1 = "牝"
    Case "021"
        KSKG1 = "牝 (指定)"
    Case "023"
        KSKG1 = "牝 [指定]"
    Case "024"
        KSKG1 = "牝 (特指)"
    Case "030"
        KSKG1 = "牡・ｾﾝ"
    Case "031"
        KSKG1 = "牡・ｾﾝ (指定)"
    Case "033"
        KSKG1 = "牡・ｾﾝ [指定]"
    Case "034"
        KSKG1 = "牡・ｾﾝ (特指)"
    Case "040"
        KSKG1 = "牡・牝"
    Case "041"
        KSKG1 = "牡・牝 (指定)"
    Case "043"
        KSKG1 = "牡・牝 [指定]"
    Case "044"
        KSKG1 = "牡・牝 (特指)"
    Case "A00"
        KSKG1 = "(混合)"
    Case "A01"
        KSKG1 = "(混合)(指定)"
    Case "A02"
        If Year < 2004 Then
            KSKG1 = "(混合) 見習騎手"
        Else
            KSKG1 = "(混合) 若手騎手"
        End If
    Case "A03"
        KSKG1 = "(混合)[指定]"
    Case "A04"
        KSKG1 = "(混合)(特指)"
    Case "A10"
        KSKG1 = "(混合) 牡"
    Case "A11"
        KSKG1 = "(混合) 牡 (指定)"
    Case "A13"
        KSKG1 = "(混合) 牡 [指定]"
    Case "A14"
        KSKG1 = "(混合) 牡 (特指)"
    Case "A20"
        KSKG1 = "(混合) 牝"
    Case "A21"
        KSKG1 = "(混合) 牝 (指定)"
    Case "A23"
        KSKG1 = "(混合) 牝 [指定]"
    Case "A24"
        KSKG1 = "(混合) 牝 (特指)"
    Case "A30"
        KSKG1 = "(混合) 牡・ｾﾝ"
    Case "A31"
        KSKG1 = "(混合) 牡・ｾﾝ (指定)"
    Case "A33"
        KSKG1 = "(混合) 牡・ｾﾝ [指定]"
    Case "A34"
        KSKG1 = "(混合) 牡・ｾﾝ (特指)"
    Case "A40"
        KSKG1 = "(混合) 牡・牝"
    Case "A41"
        KSKG1 = "(混合) 牡・牝 (指定)"
    Case "B00"
        KSKG1 = "(父)"
    Case "B01"
        KSKG1 = "(父)(指定)"
    Case "B03"
        KSKG1 = "(父)[指定]"
    Case "B04"
        KSKG1 = "(父)(特指)"
    Case "C00"
        KSKG1 = "(市)"
    Case "C01"
        KSKG1 = "(市)(指定)"
    Case "C03"
        KSKG1 = "(市)[指定]"
    Case "C04"
        KSKG1 = "(市)(特指)"
    Case "D00"
        KSKG1 = "(抽)"
    Case "D01"
        KSKG1 = "(抽)(指定)"
    Case "D03"
        KSKG1 = "(抽)[指定]"
    Case "E00"
        KSKG1 = "[抽]"
    Case "E01"
        KSKG1 = "[抽](指定)"
    Case "E03"
        KSKG1 = "[抽][指定]"
    Case "F00"
        KSKG1 = "(市)(抽)"
    Case "F01"
        KSKG1 = "(市)(抽)(指定)"
    Case "F03"
        KSKG1 = "(市)(抽)[指定]"
    Case "F04"
        KSKG1 = "(市)(抽)(特指)"
    Case "G00"
        KSKG1 = "(抽) 関西配布馬"
    Case "G01"
        KSKG1 = "(抽) 関西配布馬 (指定)"
    Case "G03"
        KSKG1 = "(抽) 関西配布馬 [指定]"
    Case "H00"
        KSKG1 = "(抽) 関東配布馬"
    Case "H01"
        KSKG1 = "(抽) 関東配布馬 (指定)"
    Case "I00"
        KSKG1 = "[抽] 関西配布馬"
    Case "I01"
        KSKG1 = "[抽] 関西配布馬 (指定)"
    Case "I03"
        KSKG1 = "[抽] 関西配布馬 [指定]"
    Case "J00"
        KSKG1 = "[抽] 関東配布馬"
    Case "J01"
        KSKG1 = "[抽] 関東配布馬 (指定)"
    Case "K00"
        KSKG1 = "(市)(抽) 関西配布馬"
    Case "K01"
        KSKG1 = "(市)(抽) 関西配布馬 (指定)"
    Case "K03"
        KSKG1 = "(市)(抽) 関西配布馬 [指定]"
    Case "L00"
        KSKG1 = "(市)(抽) 関東配布馬"
    Case "L01"
        KSKG1 = "(市)(抽) 関東配布馬 (指定)"
    Case "L03"
        KSKG1 = "(市)(抽) 関東配布馬 [指定]"
    Case "M00"
        KSKG1 = "九州産馬"
    Case "M01"
        KSKG1 = "九州産馬 (指定)"
    Case "M03"
        KSKG1 = "九州産馬 [指定]"
    Case "M04"
        KSKG1 = "九州産馬 (特指)"
    Case "N00"
        KSKG1 = "(国際)"
    Case "N01"
        KSKG1 = "(国際)(指定)"
    Case "N03"
        KSKG1 = "(国際)[指定]"
        
    Case "N04"
        KSKG1 = "(国際)(特指)"
        
    Case "N20"
        KSKG1 = "(国際) 牝"
    Case "N21"
        KSKG1 = "(国際) 牝 (指定)"
        
    Case "N23"
        KSKG1 = "(国際) 牝 [指定]"
    Case "N24"
        KSKG1 = "(国際) 牝 (特指)"
    
    Case "N30"
        KSKG1 = "(国際) 牡・ｾﾝ"
    Case "N31"
        KSKG1 = "(国際) 牡・ｾﾝ (指定)"
    Case "N40"
        KSKG1 = "(国際) 牡・牝"
    Case "N41"
        KSKG1 = "(国際) 牡・牝 (指定)"
    End Select
End Function


'
'   機能: 競走記号コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function KSKG2(strCD As String) As String
    Select Case strCD
    Case "000"
        KSKG2 = ""
    Case "001"
        KSKG2 = "DSN"
    Case "002"
        KSKG2 = ""
    Case "003"
        KSKG2 = "DES"
    Case "004"
        KSKG2 = "SD"
    Case "020"
        KSKG2 = "F&M"
    Case "021"
        KSKG2 = "F&M DSN"
    Case "023"
        KSKG2 = "F&M DES"
    Case "024"
        KSKG2 = "F&M SD"
    Case "030"
        KSKG2 = "C･G"
    Case "031"
        KSKG2 = "C･G DSN"
    Case "033"
        KSKG2 = "C･G DES"
    Case "034"
        KSKG2 = "C･G SD"
    Case "040"
        KSKG2 = "C･F"
    Case "041"
        KSKG2 = "C･F DSN"
    Case "043"
        KSKG2 = "C･F DES"
    Case "044"
        KSKG2 = "C･F SD"
    Case "A00"
        KSKG2 = "MIX"
    Case "A01"
        KSKG2 = "MIX DSN"
    Case "A02"
        KSKG2 = "MIX"
    Case "A03"
        KSKG2 = "MIX DES"
    Case "A04"
        KSKG2 = "MIX SD"
    Case "A10"
        KSKG2 = "MIX C"
    Case "A11"
        KSKG2 = "MIX C"
    Case "A13"
        KSKG2 = "MIX C DES"
    Case "A14"
        KSKG2 = "MIX C SD"
    Case "A20"
        KSKG2 = "MIX F&M"
    Case "A21"
        KSKG2 = "MIX F&M DSN"
    Case "A23"
        KSKG2 = "MIX F&M DES"
    Case "A24"
        KSKG2 = "MIX F&M SD"
    Case "A30"
        KSKG2 = "MIX C･G"
    Case "A31"
        KSKG2 = "MIX C･G DSN"
    Case "A33"
        KSKG2 = "MIX C･G DES"
    Case "A34"
        KSKG2 = "MIX C･G SD"
    Case "A40"
        KSKG2 = "MIX C･F"
    Case "A41"
        KSKG2 = "MIX C･F DSN"
    Case "B00"
        KSKG2 = "(D)"
    Case "B01"
        KSKG2 = "(D) DSN"
    Case "B03"
        KSKG2 = "(D) DES"
    Case "B04"
        KSKG2 = "(D) SD"
    Case "C00"
        KSKG2 = "(A)"
    Case "C01"
        KSKG2 = "(A) DSN"
    Case "C03"
        KSKG2 = "(A) DES"
    Case "C04"
        KSKG2 = "(A) SD"
    Case "D00"
        KSKG2 = "(S)"
    Case "D01"
        KSKG2 = "(S) DSN"
    Case "D03"
        KSKG2 = "(S) DES"
    Case "E00"
        KSKG2 = ""
    Case "E01"
        KSKG2 = "DSN"
    Case "E03"
        KSKG2 = "DES"
    Case "F00"
        KSKG2 = "(A)(S)"
    Case "F01"
        KSKG2 = "(A)(S) DSN"
    Case "F03"
        KSKG2 = "(A)(S) DES"
    Case "F04"
        KSKG2 = "(A)(S) SD"
    Case "G00"
        KSKG2 = ""
    Case "G01"
        KSKG2 = "DSN"
    Case "G03"
        KSKG2 = "DES"
    Case "H00"
        KSKG2 = ""
    Case "H01"
        KSKG2 = "DSN"
    Case "I00"
        KSKG2 = ""
    Case "I01"
        KSKG2 = "DSN"
    Case "I03"
        KSKG2 = "DES"
    Case "J00"
        KSKG2 = ""
    Case "J01"
        KSKG2 = "DSN"
    Case "K00"
        KSKG2 = ""
    Case "K01"
        KSKG2 = "DSN"
    Case "K03"
        KSKG2 = "DES"
    Case "L00"
        KSKG2 = ""
    Case "L01"
        KSKG2 = "DSN"
    Case "L03"
        KSKG2 = "DES"
    Case "M00"
        KSKG2 = ""
    Case "M01"
        KSKG2 = "DSN"
    Case "M03"
        KSKG2 = "DES"
    Case "M04"
        KSKG2 = "SD"
    Case "N00"
        KSKG2 = "INT"
    Case "N01"
        KSKG2 = "INT DSN"
    Case "N03"
        KSKG2 = "INT DES"
        
    Case "N04"
        KSKG2 = "INT SD"
        
    Case "N20"
        KSKG2 = "INT F&M"
    Case "N21"
        KSKG2 = "INT F&M DSN"
    
    Case "N23"
        KSKG2 = "INT F&M DES"
    Case "N24"
        KSKG2 = "INT F&M SD"
        
    Case "N30"
        KSKG2 = "INT C･G"
    Case "N31"
        KSKG2 = "INT C･G DSN"
    Case "N40"
        KSKG2 = "INT C･F"
    Case "N41"
        KSKG2 = "INT C･F DSN"
    End Select
End Function


'
'   機能: 競走条件コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function KSJK1(strCD As String) As String
    Select Case val(strCD)
    Case 0
        KSJK1 = ""
    Case 1 To 99
        KSJK1 = StrConv(100 * val(strCD), vbWide) & "万円以下"
    Case 100
        KSJK1 = "１億円以下"
    Case "701"
        KSJK1 = "新馬"
    Case "702"
        KSJK1 = "未出走"
    Case "703"
        KSJK1 = "未勝利"
    Case "999"
        KSJK1 = "オープン"
    End Select
End Function


'
'   機能: 競走条件コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function KSJK2(strCD As String) As String
    Select Case val(strCD)
    Case 0
        KSJK2 = ""
    Case 1 To 99
        KSJK2 = val(strCD) & ",000,000 & LESS"
    Case 100
        KSJK2 = "100,000,000 & LESS"
    Case "701"
        KSJK2 = "NEWCOMER"
    Case "702"
        KSJK2 = "UNRACED"
    Case "703"
        KSJK2 = "MAIDEN"
    Case "999"
        KSJK2 = "OPEN"
    End Select
End Function


'
'   機能: 競走条件コード の 説明 を可読文字列にする
'
'   備考: なし
'
Public Function KSJK3(strCD As String) As String
    Select Case val(strCD)
    Case 0
        KSJK3 = "未設定・未整備時の初期値(主に地方競馬・海外国際レースに関するデータ)"
    Case 1 To 99
        KSJK3 = "収得賞金が" & val(strCD) & "00万円以下の馬が出走できる。"
    Case 100
        KSJK3 = "収得賞金が1億円以下の馬が出走できる。"
    Case "701"
        KSJK3 = "サラブレッド系の未出走馬が出走できる。"
    Case "702"
        KSJK3 = "当該競馬以外の競馬(地方競馬または外国の競馬を含む)において、出走したことがなく、当該競馬において、第1着または重賞競走の第2着になったことのない馬が出走できる。"
    Case "703"
        KSJK3 = "該当競馬以外の競馬(地方競馬または外国の競馬を含む)において出走し、収得賞金がない馬が出走できる。また未出走馬も出走できる。"
    Case "999"
        KSJK3 = "特に記載のない限り、すべての馬が出走できる競走"
    End Select
End Function


'
'   機能: 競走条件コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function KSJK4(strCD As String) As String
    Select Case val(strCD)
    Case 0
        KSJK4 = ""
    Case 1 To 99
        KSJK4 = 100 * val(strCD) & "万下"
    Case 100
        KSJK4 = "１億"
    Case "701"
        KSJK4 = "新馬"
    Case "702"
        KSJK4 = "未出走"
    Case "703"
        KSJK4 = "未勝利"
    Case "999"
        KSJK4 = "ｵｰﾌﾟﾝ"
    End Select
End Function


'
'   機能: 競走条件コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function KSJK5(strCD As String) As String
    Select Case val(strCD)
    Case 0
        KSJK5 = "      "
    Case 1 To 99
        KSJK5 = Format$(100 * val(strCD), "@@@@") & "下"
    Case 100
        KSJK5 = "１億  "
    Case "701"
        KSJK5 = "新馬  "
    Case "702"
        KSJK5 = "未出走"
    Case "703"
        KSJK5 = "未勝利"
    Case "999"
        KSJK5 = "ｵｰﾌﾟﾝ "
    End Select
End Function


'
'   機能: 重量種別コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function JRSB1(strCD As String) As String
    Select Case strCD
    Case "0"
        JRSB1 = ""
    Case "1"
        JRSB1 = "ハンデ"
    Case "2"
        JRSB1 = "別定"
    Case "3"
        JRSB1 = "馬齢"
    Case "4"
        JRSB1 = "定量"
    End Select
End Function


'
'   機能: 重量種別コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function JRSB2(strCD As String) As String
    Select Case strCD
    Case "0"
        JRSB2 = ""
    Case "1"
        JRSB2 = "HANDICAP"
    Case "2"
        JRSB2 = "SPECIAL WEIGHT"
    Case "3"
        JRSB2 = "WEIGHT FOR AGE"
    Case "4"
        JRSB2 = "SPECIAL WEIGHT"
    End Select
End Function


'
'   機能: 重量種別コード の 説明 を可読文字列にする
'
'   備考: なし
'
Public Function JRSB3(strCD As String) As String
    Select Case strCD
    Case "0"
        JRSB3 = "未設定・未整備時の初期値(主に地方競馬・海外国際レースに関するデータ)"
    Case "1"
        JRSB3 = "出走馬の実績等を考慮し、ハンデキャッパーが負担重量を決定するレース"
    Case "2"
        JRSB3 = "レースごとに負担重量を決定する基準が設けられているレース"
    Case "3"
        JRSB3 = "馬の年齢や性別によって負担重量を決定するレース"
    Case "4"
        JRSB3 = "別定レースであって、馬の年齢や性別によって負担重量を決定するレース"
    End Select
End Function


'
'   機能: トラックコード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function TRCK1(strCD As String) As String
    Select Case strCD
    Case "00"
        TRCK1 = ""
    Case "10"
        TRCK1 = "平地　芝 　　 直線"
    Case "11"
        TRCK1 = "平地　芝 　　 左回り"
    Case "12"
        TRCK1 = "平地　芝  　　左回り  外回り"
    Case "13"
        TRCK1 = "平地　芝  　　左回り  内－外回り"
    Case "14"
        TRCK1 = "平地　芝  　　左回り  外－内回り"
    Case "15"
        TRCK1 = "平地　芝  　　左回り  内２周"
    Case "16"
        TRCK1 = "平地　芝  　　左回り  外２周"
    Case "17"
        TRCK1 = "平地　芝  　　右回り"
    Case "18"
        TRCK1 = "平地　芝  　　右回り  外回り"
    Case "19"
        TRCK1 = "平地　芝  　　右回り  内－外回り"
    Case "20"
        TRCK1 = "平地　芝  　　右回り  外－内回り"
    Case "21"
        TRCK1 = "平地　芝  　　右回り  内２周"
    Case "22"
        TRCK1 = "平地　芝  　　右回り  外２周"
    Case "23"
        TRCK1 = "平地　ダート  左回り"
    Case "24"
        TRCK1 = "平地　ダート  右回り"
    Case "25"
        TRCK1 = "平地　ダート　左回り　内回り"
    Case "26"
        TRCK1 = "平地　ダート　右回り　外回り"
    Case "27"
        TRCK1 = "平地　サンド　左回り"
    Case "28"
        TRCK1 = "平地　サンド　右回り"
    Case "29"
        TRCK1 = "平地　ダート  直線"
    Case "51"
        TRCK1 = "障害　芝  襷"
    Case "52"
        TRCK1 = "障害　芝　ダート"
    Case "53"
        TRCK1 = "障害　芝・左"
    Case "54"
        TRCK1 = "障害　芝"
    Case "55"
        TRCK1 = "障害　芝  外回り"
    Case "56"
        TRCK1 = "障害　芝  外－内回り"
    Case "57"
        TRCK1 = "障害　芝  内－外回り"
    Case "58"
        TRCK1 = "障害　芝  内２周"
    Case "59"
        TRCK1 = "障害　芝  外２周"
    End Select
End Function


'
'   機能: トラックコード の 略名(6文字) を可読文字列にする
'
'   備考: なし
'
Public Function TRCK2(strCD As String) As String
    Select Case strCD
    Case "00"
        TRCK2 = ""
    Case "10"
        TRCK2 = "芝・直"
    Case "11"
        TRCK2 = "芝・左"
    Case "12"
        TRCK2 = "芝・左外"
    Case "13"
        TRCK2 = "芝・左内→外"
    Case "14"
        TRCK2 = "芝・左外→内"
    Case "15"
        TRCK2 = "芝・左内２周"
    Case "16"
        TRCK2 = "芝・左外２周"
    Case "17"
        TRCK2 = "芝・右"
    Case "18"
        TRCK2 = "芝・右外"
    Case "19"
        TRCK2 = "芝・右内→外"
    Case "20"
        TRCK2 = "芝・右外→内"
    Case "21"
        TRCK2 = "芝・右内２周"
    Case "22"
        TRCK2 = "芝・右外２周"
    Case "23"
        TRCK2 = "ダート・左"
    Case "24"
        TRCK2 = "ダート・右"
    Case "25"
        TRCK2 = "ダート・左内"
    Case "26"
        TRCK2 = "ダート・右外"
    Case "27"
        TRCK2 = "サンド・左"
    Case "28"
        TRCK2 = "サンド・右"
    Case "29"
        TRCK2 = "ダート・直"
    Case "51"
        TRCK2 = "芝・襷"
    Case "52"
        TRCK2 = "芝→ダート"
    Case "53"
        TRCK2 = "芝・左"
    Case "54"
        TRCK2 = "芝"
    Case "55"
        TRCK2 = "芝・外"
    Case "56"
        TRCK2 = "芝・外→内"
    Case "57"
        TRCK2 = "芝・内→外"
    Case "58"
        TRCK2 = "芝・内２周"
    Case "59"
        TRCK2 = "芝・外２周"
    End Select
End Function


'
'   機能: トラックコード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function TRCK3(strCD As String) As String
    Select Case strCD
    Case "00"
        TRCK3 = ""
    Case "10"
        TRCK3 = "Turf Str."
    Case "11"
        TRCK3 = "Turf"
    Case "12"
        TRCK3 = "Turf"
    Case "13"
        TRCK3 = "Turf"
    Case "14"
        TRCK3 = "Turf"
    Case "15"
        TRCK3 = "Turf"
    Case "16"
        TRCK3 = "Turf"
    Case "17"
        TRCK3 = "Turf"
    Case "18"
        TRCK3 = "Turf"
    Case "19"
        TRCK3 = "Turf"
    Case "20"
        TRCK3 = "Turf"
    Case "21"
        TRCK3 = "Turf"
    Case "22"
        TRCK3 = "Turf"
    Case "23"
        TRCK3 = "Dirt"
    Case "24"
        TRCK3 = "Dirt"
    Case "25"
        TRCK3 = "Dirt"
    Case "26"
        TRCK3 = "Dirt"
    Case "27"
        TRCK3 = "Sand"
    Case "28"
        TRCK3 = "Sand"
    Case "29"
        TRCK3 = "Dirt Str."
    Case "51"
        TRCK3 = "Turf"
    Case "52"
        TRCK3 = "Turf→Dirt"
    Case "53"
        TRCK3 = "Turf"
    Case "54"
        TRCK3 = "Turf"
    Case "55"
        TRCK3 = "Turf"
    Case "56"
        TRCK3 = "Turf"
    Case "57"
        TRCK3 = "Turf"
    Case "58"
        TRCK3 = "Turf"
    Case "59"
        TRCK3 = "Turf"
    End Select
End Function


'
'   機能: トラックコード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function TRCK4(strCD As String) As String
    Select Case strCD
    Case "00"
        TRCK4 = ""
    Case "10" To "22"
        TRCK4 = "芝"
    Case "23" To 26, "29"
        TRCK4 = "ダ"
    Case "27", "28"
        TRCK4 = "砂"
    Case "51" To "59"
        TRCK4 = "障"
    End Select
End Function


'
'   機能: トラックコード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function TRCK5(strCD As String) As String
    Select Case strCD
    Case "00"
        TRCK5 = ""
    Case "10"
        TRCK5 = "芝直"
    Case "11" To "16"
        TRCK5 = "芝左"
    Case "16" To "22"
        TRCK5 = "芝右"
    Case "23", "25"
        TRCK5 = "ダ左"
    Case "24" To "26"
        TRCK5 = "ダ右"
    Case "27"
        TRCK5 = "砂左"
    Case "28"
        TRCK5 = "砂右"
    Case "29"
        TRCK5 = "ダ直"
    Case "51" To "59"
        TRCK5 = "障害"
    End Select
End Function


'
'   機能: 馬場状態コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function BBJT1(strCD As String) As String
    Select Case strCD
    Case "0"
        BBJT1 = ""
    Case "1"
        BBJT1 = "良"
    Case "2"
        BBJT1 = "稍重"
    Case "3"
        BBJT1 = "重"
    Case "4"
        BBJT1 = "不良"
    End Select
End Function


'
'   機能: 馬場状態コード の 欧字名（芝） を可読文字列にする
'
'   備考: なし
'
Public Function BBJT2(strCD As String) As String
    Select Case strCD
    Case "0"
        BBJT2 = ""
    Case "1"
        BBJT2 = "Firm"
    Case "2"
        BBJT2 = "Good"
    Case "3"
        BBJT2 = "Yielding"
    Case "4"
        BBJT2 = "Soft"
    End Select
End Function


'
'   機能: 馬場状態コード の 欧字名（ダート） を可読文字列にする
'
'   備考: なし
'
Public Function BBJT3(strCD As String) As String
    Select Case strCD
    Case "0"
        BBJT3 = ""
    Case "1"
        BBJT3 = "Standard"
    Case "2"
        BBJT3 = "Good"
    Case "3"
        BBJT3 = "Muddy"
    Case "4"
        BBJT3 = "Sloppy"
    End Select
End Function


'
'   機能: 馬場状態コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function BBJT4(strCD As String) As String
    Select Case strCD
    Case "0"
        BBJT4 = ""
    Case "1"
        BBJT4 = "良"
    Case "2"
        BBJT4 = "稍"
    Case "3"
        BBJT4 = "重"
    Case "4"
        BBJT4 = "不"
    End Select
End Function


'
'   機能: 天候コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function TNKO1(strCD As String) As String
    Select Case strCD
    Case "0"
        TNKO1 = ""
    Case "1"
        TNKO1 = "晴"
    Case "2"
        TNKO1 = "曇"
    Case "3"
        TNKO1 = "雨"
    Case "4"
        TNKO1 = "小雨"
    Case "5"
        TNKO1 = "雪"
    Case "6"
        TNKO1 = "小雪"
    End Select
End Function


'
'   機能: 天候コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function TNKO2(strCD As String) As String
    Select Case strCD
    Case "0"
        TNKO2 = ""
    Case "1"
        TNKO2 = "Fine"
    Case "2"
        TNKO2 = "Cloudy"
    Case "3"
        TNKO2 = "Rainy"
    Case "4"
        TNKO2 = "Drizzle"
    Case "5"
        TNKO2 = "Snow"
    Case "6"
        TNKO2 = "Light Snow"
    End Select
End Function


'
'   機能: 異常区分コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function IJKB1(strCD As String) As String
    Select Case strCD
    Case "0"
        IJKB1 = ""
    Case "1"
        IJKB1 = "出走取消"
    Case "2"
        IJKB1 = "発走除外"
    Case "3"
        IJKB1 = "競走除外"
    Case "4"
        IJKB1 = "競走中止"
    Case "5"
        IJKB1 = "失格"
    Case "6"
        IJKB1 = "落馬再騎乗"
    Case "7"
        IJKB1 = "降着"
    End Select
End Function


'
'   機能: 異常区分コード の 略名(2文字) を可読文字列にする
'
'   備考: なし
'
Public Function IJKB2(strCD As String) As String
    Select Case strCD
    Case "0"
        IJKB2 = ""
    Case "1"
        IJKB2 = "取消"
    Case "2"
        IJKB2 = "発除"
    Case "3"
        IJKB2 = "競除"
    Case "4"
        IJKB2 = "中止"
    Case "5"
        IJKB2 = "失格"
    Case "6"
        IJKB2 = "再騎"
    Case "7"
        IJKB2 = "降着"
    End Select
End Function


'
'   機能: 異常区分コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function IJKB3(strCD As String) As String
    Select Case strCD
    Case "0"
        IJKB3 = ""
    Case "1"
        IJKB3 = "SCRATCHED"
    Case "2"
        IJKB3 = "EXCLUDED BY STARTERS"
    Case "3"
        IJKB3 = "EXCLUDED BY STEWARDS"
    Case "4"
        IJKB3 = "FALL TO FINISH"
    Case "5"
        IJKB3 = "DISQUALIFIED"
    Case "6"
        IJKB3 = "REMOUNT AFTER A CROPPER"
    Case "7"
        IJKB3 = "DISQUALIFIED AND PLACED"
    End Select
End Function


'
'   機能: 異常区分コード の 欧字略名 を可読文字列にする
'
'   備考: なし
'
Public Function IJKB4(strCD As String) As String
    Select Case strCD
    Case "0"
        IJKB4 = ""
    Case "1"
        IJKB4 = "S"
    Case "2"
        IJKB4 = "ES"
    Case "3"
        IJKB4 = "ER"
    Case "4"
        IJKB4 = "FF"
    Case "5"
        IJKB4 = "DQ"
    Case "6"
        IJKB4 = "RM"
    Case "7"
        IJKB4 = "DQ&P"
    End Select
End Function


'
'   機能: 異常区分コード の 略名(2文字) を可読文字列にする
'
'   備考: なし
'
Public Function IJKB5(strCD As String) As String
    Select Case strCD
    Case "0"
        IJKB5 = ""
    Case "1"
        IJKB5 = "消"
    Case "2"
        IJKB5 = "除"
    Case "3"
        IJKB5 = "除"
    Case "4"
        IJKB5 = "止"
    Case "5"
        IJKB5 = "失"
    Case "6"
        IJKB5 = "落"
    Case "7"
        IJKB5 = "降"
    End Select
End Function


'
'   機能: 着差コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function CHKS1(strCD As String) As String
On Error GoTo Errorhandler
    Dim strP1 As String
    Dim strP2 As String
    Dim strP3 As String
    Dim strOut As String
    
    strP1 = Mid(strCD, 1, 1)
    strP2 = Mid(strCD, 2, 1)
    strP3 = Mid(strCD, 3, 1)
    Select Case strP1
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", " "
        strOut = StrConv(Trim(strP1), vbWide)
        Select Case strP2
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            strOut = strOut & strP2
            Select Case strP3
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                strOut = strOut & "/" & strP3 & "馬身"
            Case Else
                strOut = "" '書式エラー
            End Select
        Case Else
            If strOut <> "" Then
                strOut = strOut & "馬身"
            Else
                strOut = "" '書式エラー
            End If
        End Select
    Case "A"
        strOut = "アタマ"
    Case "D"
        strOut = "同着"
    Case "H"
        strOut = "ハナ"
    Case "K"
        strOut = "クビ"
    Case "T"
        strOut = "大差"
    Case "Z"
        strOut = "１０馬身"
    Case Else
        strOut = "" '書式エラー
    End Select
    
    CHKS1 = strOut
    Exit Function

Errorhandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: 着差コード の 略名(3文字) を可読文字列にする
'
'   備考: なし
'
Public Function CHKS2(strCD As String) As String
On Error GoTo Errorhandler
    Dim strP1 As String
    Dim strP2 As String
    Dim strP3 As String
    Dim strOut As String
    
    strP1 = Mid(strCD, 1, 1)
    strP2 = Mid(strCD, 2, 1)
    strP3 = Mid(strCD, 3, 1)
    Select Case strP1
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", " "
        strOut = StrConv(Trim(strP1), vbWide)
        Select Case strP2
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            strOut = strOut & strP2
            Select Case strP3
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                strOut = strOut & "/" & strP3
            Case Else
                strOut = "" '書式エラー
            End Select
        Case Else
            If strOut <> "" Then
                strOut = strOut
            Else
                strOut = "" '書式エラー
            End If
        End Select
    Case "A"
        strOut = "アタマ"
    Case "D"
        strOut = "同着"
    Case "H"
        strOut = "ハナ"
    Case "K"
        strOut = "クビ"
    Case "T"
        strOut = "大差"
    Case "Z"
        strOut = "１０"
    Case Else
        strOut = "" '書式エラー
    End Select
    
    CHKS2 = strOut
    Exit Function
    
Errorhandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: 着差コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function CHKS3(strCD As String) As String
On Error GoTo Errorhandler
    Dim strP1 As String
    Dim strP2 As String
    Dim strP3 As String
    Dim strOut As String
    
    strP1 = Mid(strCD, 1, 1)
    strP2 = Mid(strCD, 2, 1)
    strP3 = Mid(strCD, 3, 1)
    Select Case strP1
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", " "
        strOut = StrConv(Trim(strP1), vbWide)
        Select Case strP2
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            strOut = strOut & strP2
            Select Case strP3
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                strOut = strOut & "/" & strP3
            Case Else
                strOut = "" '書式エラー
            End Select
        Case Else
            If strOut <> "" Then
                strOut = strOut
            Else
                strOut = "" '書式エラー
            End If
        End Select
    Case "A"
        strOut = "HD"
    Case "D"
        strOut = "DH"
    Case "H"
        strOut = "NS"
    Case "K"
        strOut = "NK"
    Case "T"
        strOut = "DS"
    Case "Z"
        strOut = "１０"
    Case Else
        strOut = "" '書式エラー
    End Select
    
    CHKS3 = strOut
    Exit Function
    
Errorhandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: 品種コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function HINS1(strCD As String) As String
    Select Case strCD
    Case "0"
        HINS1 = ""
    Case "1"
        HINS1 = "サラブレッド"
    Case "2"
        HINS1 = "サラブレッド系種"
    Case "3"
        HINS1 = "準サラブレッド"
    Case "4"
        HINS1 = "軽半血種"
    Case "5"
        HINS1 = "アングロアラブ"
    Case "6"
        HINS1 = "アラブ系種"
    Case "7"
        HINS1 = "アラブ"
    Case "8"
        HINS1 = "中半血種"
    End Select
End Function


'
'   機能: 品種コード の 略名(3文字) を可読文字列にする
'
'   備考: なし
'
Public Function HINS2(strCD As String) As String
    Select Case strCD
    Case "0"
        HINS2 = ""
    Case "1"
        HINS2 = "サラ"
    Case "2"
        HINS2 = "サラ系"
    Case "3"
        HINS2 = "準サラ"
    Case "4"
        HINS2 = "軽半"
    Case "5"
        HINS2 = "アア"
    Case "6"
        HINS2 = "アラ系"
    Case "7"
        HINS2 = "アラブ"
    Case "8"
        HINS2 = "中半"
    End Select
End Function


'
'   機能: 品種コード の 説明 を可読文字列にする
'
'   備考: なし
'
Public Function HINS3(strCD As String) As String
    Select Case strCD
    Case "0"
        HINS3 = "未設定・未整備時の初期値(主に地方競馬・海外国際レースに関するデータ)"
    Case "1"
        HINS3 = "サラブレッドとして登録したもの。"
    Case "2"
        HINS3 = "アラブの血量が２５％未満"
    Case "3"
        HINS3 = "昭和49年6月1日の登録規定改定により「準サラ」は「サラ系」となり、「準サラ」の品種は廃止された。"
    Case "4"
        HINS3 = "アングロノルマン、アングロノルマン系種または中半血種と連続2代以上にあたり、軽種を交配したもの。軽半血種相互の交配によって生まれたもの。"
    Case "5"
        HINS3 = "アラブの血量が２５％以上（父母の組み合わせにより「アラ系」と異なる。）"
    Case "6"
        HINS3 = "アラブの血量が２５％以上（父母の組み合わせにより「アア」と異なる。）"
    Case "7"
        HINS3 = "純血のアラブ"
    Case "8"
        HINS3 = "異種の中間相互、中半血種の中間相互、の交配で生まれたもの。中半血種と軽種、アングロノルマンとサラブレッドを除く軽種、重半血種と軽種を交配して生まれたもの。"
    End Select
End Function


'
'   機能: 性別コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function SEIB1(strCD As String) As String
    Select Case strCD
    Case "0"
        SEIB1 = ""
    Case "1"
        SEIB1 = "牡馬"
    Case "2"
        SEIB1 = "牝馬"
    Case "3"
        SEIB1 = "セン馬"
    End Select
End Function


'
'   機能: 性別コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function SEIB2(strCD As String) As String
    Select Case strCD
    Case "0"
        SEIB2 = ""
    Case "1"
        SEIB2 = "満年齢0～4歳　C 満年齢5歳～ 　H"
    Case "2"
        SEIB2 = "満年齢0～4歳　F 満年齢5歳～ 　M"
    Case "3"
        SEIB2 = "満年齢0～4歳　G 満年齢5歳～ 　G"
    End Select
End Function


'
'   機能: 性別コード の 説明 を可読文字列にする
'
'   備考: なし
'
Public Function SEIB3(strCD As String) As String
    Select Case strCD
    Case "0"
        SEIB3 = ""
    Case "1"
        SEIB3 = "オス馬"
    Case "2"
        SEIB3 = "メス馬"
    Case "3"
        SEIB3 = "去勢したオス馬"
    End Select
End Function


'
'   機能: 性別コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function SEIB4(strCD As String) As String
    Select Case strCD
    Case "0"
        SEIB4 = ""
    Case "1"
        SEIB4 = "牡"
    Case "2"
        SEIB4 = "牝"
    Case "3"
        SEIB4 = "騙"
    End Select
End Function


'
'   機能: 毛色コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function KEIR1(strCD As String) As String
    Select Case strCD
    Case "00"
        KEIR1 = ""
    Case "01"
        KEIR1 = "栗毛"
    Case "02"
        KEIR1 = "栃栗毛"
    Case "03"
        KEIR1 = "鹿毛"
    Case "04"
        KEIR1 = "黒鹿毛"
    Case "05"
        KEIR1 = "青鹿毛"
    Case "06"
        KEIR1 = "青毛"
    Case "07"
        KEIR1 = "芦毛"
    Case "08"
        KEIR1 = "栗粕毛"
    Case "09"
        KEIR1 = "鹿粕毛"
    Case "10"
        KEIR1 = "青粕毛"
    Case "11"
        KEIR1 = "白毛"
    End Select
End Function


'
'   機能: 毛色コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function KEIR2(strCD As String) As String
    Select Case strCD
    Case "00"
        KEIR2 = ""
    Case "01"
        KEIR2 = "chestnut"
    Case "02"
        KEIR2 = "dark chestnut"
    Case "03"
        KEIR2 = "bay"
    Case "04"
        KEIR2 = "dark bay"
    Case "05"
        KEIR2 = "brown"
    Case "06"
        KEIR2 = "black"
    Case "07"
        KEIR2 = "grey"
    Case "08"
        KEIR2 = ""
    Case "09"
        KEIR2 = ""
    Case "10"
        KEIR2 = ""
    Case "11"
        KEIR2 = "white"
    End Select
End Function


'
'   機能: 毛色コード の 欧字略名(5文字) を可読文字列にする
'
'   備考: なし
'
Public Function KEIR3(strCD As String) As String
    Select Case strCD
    Case "00"
        KEIR3 = ""
    Case "01"
        KEIR3 = "ch."
    Case "02"
        KEIR3 = "d.ch."
    Case "03"
        KEIR3 = "b."
    Case "04"
        KEIR3 = "d.b."
    Case "05"
        KEIR3 = "br."
    Case "06"
        KEIR3 = "bl."
    Case "07"
        KEIR3 = "g."
    Case "08"
        KEIR3 = ""
    Case "09"
        KEIR3 = ""
    Case "10"
        KEIR3 = ""
    Case "11"
        KEIR3 = "w."
    End Select
End Function


'
'   機能: 馬記号コード の 名称 を可読文字列にする
'
'   備考: なし
'
Public Function UMKG1(strCD As String) As String
    Select Case strCD
    Case "00"
        UMKG1 = ""
    Case "01"
        UMKG1 = "(抽)"
    Case "02"
        UMKG1 = "[抽]"
    Case "03"
        UMKG1 = "(父)"
    Case "04"
        UMKG1 = "(市)"
    Case "05"
        UMKG1 = "(地)"
    Case "06"
        UMKG1 = "(外)"
    Case "07"
        UMKG1 = "(父)(抽)"
    Case "08"
        UMKG1 = "(父)(市)"
    Case "09"
        UMKG1 = "(父)(地)"
    Case "10"
        UMKG1 = "(市)(地)"
    Case "11"
        UMKG1 = "(外)(地)"
    Case "12"
        UMKG1 = "(父)(市)(地)"
    Case "15"
        UMKG1 = "(招)"
    Case "16"
        UMKG1 = "(招)(外)"
    Case "17"
        UMKG1 = "(招)(父)"
    Case "18"
        UMKG1 = "(招)(市)"
    Case "19"
        UMKG1 = "(招)(父)(市)"
    Case "20"
        UMKG1 = "(父)(外)"
    Case "21"
        UMKG1 = "[地]"
    Case "22"
        UMKG1 = "(外)[地]"
    Case "23"
        UMKG1 = "(父)[地]"
    Case "24"
        UMKG1 = "(市)[地]"
    Case "25"
        UMKG1 = "(父)(市)[地]"
    Case "26"
        UMKG1 = "[外]"
    Case "27"
        UMKG1 = "(父)[外]"
    Case "40"
        UMKG1 = "(父)(外)(地)"
    Case "41"
        UMKG1 = "(父)(外)[地]"
    End Select
End Function


'
'   機能: 馬記号コード の 欧字名 を可読文字列にする
'
'   備考: なし
'
Public Function UMKG2(strCD As String) As String
    Select Case strCD
    Case "00"
        UMKG2 = ""
    Case "01"
        UMKG2 = "(S)"
    Case "02"
        UMKG2 = ""
    Case "03"
        UMKG2 = "(D)"
    Case "04"
        UMKG2 = "(A)"
    Case "05"
        UMKG2 = "(R)"
    Case "06"
        UMKG2 = "(F)"
    Case "07"
        UMKG2 = "(D)(S)"
    Case "08"
        UMKG2 = "(D)(A)"
    Case "09"
        UMKG2 = "(D)(R)"
    Case "10"
        UMKG2 = "(A)(R)"
    Case "11"
        UMKG2 = "(F)(R)"
    Case "12"
        UMKG2 = "(D)(A)(R)"
    Case "15"
        UMKG2 = "(I)"
    Case "16"
        UMKG2 = "(I)(F)"
    Case "17"
        UMKG2 = "(I)(D)"
    Case "18"
        UMKG2 = "(I)(A)"
    Case "19"
        UMKG2 = "(I)(D)(A)"
    Case "20"
        UMKG2 = ""
    Case "21"
        UMKG2 = "[R]"
    Case "22"
        UMKG2 = "(F)[R]"
    Case "23"
        UMKG2 = "(D)[R]"
    Case "24"
        UMKG2 = "(A)[R]"
    Case "25"
        UMKG2 = "(D)(A)[R]"
    Case "26"
        UMKG2 = "[F]"
    Case "27"
        UMKG2 = "(D)[F]"
    Case "40"
        UMKG2 = "(D)(F)(R)"
    Case "41"
        UMKG2 = "(D)(F)[R]"
    End Select
End Function


'
'   機能: 馬記号コード の 説明 を可読文字列にする
'
'   備考: なし
'
Public Function UMKG3(strCD As String) As String
    Select Case strCD
    Case Else
        UMKG3 = ""
    End Select
End Function


'
'   機能: 東西所属コード の 名称１ を可読文字列にする
'
'   備考: なし
'
Public Function TZSZ1(strCD As String) As String
    Select Case strCD
    Case "0"
        TZSZ1 = ""
    Case "1"
        TZSZ1 = "関東"
    Case "2"
        TZSZ1 = "関西"
    Case "3"
        TZSZ1 = "地方招待"
    Case "4"
        TZSZ1 = "外国招待"
    End Select
End Function


'
'   機能: 東西所属コード の 名称２ を可読文字列にする
'
'   備考: なし
'
Public Function TZSZ2(strCD As String) As String
    Select Case strCD
    Case "0"
        TZSZ2 = ""
    Case "1"
        TZSZ2 = "美浦"
    Case "2"
        TZSZ2 = "栗東"
    Case "3"
        TZSZ2 = "招待"
    Case "4"
        TZSZ2 = "招待"
    End Select
End Function


'
'   機能: 東西所属コード の 説明 を可読文字列にする
'
'   備考: なし
'
Public Function TZSZ3(strCD As String) As String
    Select Case strCD
    Case "0"
        TZSZ3 = ""
    Case "1"
        TZSZ3 = "美浦トレーニングセンターに所属する。"
    Case "2"
        TZSZ3 = "栗東トレーニングセンターに所属する。"
    Case "3"
        TZSZ3 = "地方からの招待"
    Case "4"
        TZSZ3 = "外国からの招待"
    End Select
End Function


'
'   機能: 東西所属コード の 名称１ を可読文字列にする
'
'   備考: なし
'
Public Function TZSZ4(strCD As String) As String
    Select Case strCD
    Case "0"
        TZSZ4 = ""
    Case "1"
        TZSZ4 = "(東)"
    Case "2"
        TZSZ4 = "(西)"
    Case "3"
        TZSZ4 = "(招)"
    Case "4"
        TZSZ4 = "(招)"
    End Select
End Function


'
'   機能: 東西所属コード の 名称１ を可読文字列にする
'
'   備考: なし
'
Public Function TZSZ5(strCD As String) As String
    Select Case strCD
    Case "0"
        TZSZ5 = ""
    Case "1"
        TZSZ5 = "東"
    Case "2"
        TZSZ5 = "西"
    Case "3"
        TZSZ5 = "地"
    Case "4"
        TZSZ5 = "外"
    End Select
End Function


'
'   機能: 騎乗資格コード の 内容 を可読文字列にする
'
'   備考: なし
'
Public Function KSSK1(strCD As String) As String
    Select Case strCD
    Case "0"
        KSSK1 = ""
    Case "1"
        KSSK1 = "平・障"
    Case "2"
        KSSK1 = "平地"
    Case "3"
        KSSK1 = "障害"
    End Select
End Function


'
'   機能: 騎手見習コード の 略名 を可読文字列にする
'
'   備考: なし
'
Public Function KSMN1(strCD As String) As String
    Select Case strCD
    Case "0"
        KSMN1 = ""
    Case "1"
        KSMN1 = "☆"
    Case "2"
        KSMN1 = "△"
    Case "3"
        KSMN1 = "▲"
    End Select
End Function


'
'   機能: 騎手見習コード の 減量値 を可読文字列にする
'
'   備考: なし
'
Public Function KSMN2(strCD As String) As String
    Select Case strCD
    Case "0"
        KSMN2 = ""
    Case "1"
        KSMN2 = "" ' "1Kg減"
    Case "2"
        KSMN2 = "" ' "2Kg減"
    Case "3"
        KSMN2 = "" ' "3Kg減"
    End Select
End Function


'
'   機能: 騎手見習コード の 説明 を可読文字列にする
'
'   備考: なし
'
Public Function KSMN3(strCD As String) As String
    Select Case strCD
    Case "0"
        KSMN3 = ""
    Case "1"
        KSMN3 = "" ' "通算免許期間が3年未満の騎手で、勝利数が31回以上100回以下。"
    Case "2"
        KSMN3 = "" ' "通算免許期間が3年未満の騎手で、勝利数が21回以上30回以下。"
    Case "3"
        KSMN3 = "" ' "通算免許期間が3年未満の騎手で、勝利数が20回以下。"
    End Select
End Function




