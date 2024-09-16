Option Explicit On 

Module basConst
    '========================================================================
    '  JRA-VAN Data Lab.プログラミングパーツ「Public変数定義ファイル」
    '
    '
    '   作成: JRA-VAN ソフトウェア工房  2004年12月28日
    '
    '========================================================================
    '   (C) Copyright Turf Media System Co.,Ltd. 2003 All rights reserved
    '========================================================================

    ' -----データ区分-----
    ' レース詳細
    Public Const ID_RACE As String = "RA"

    ' 馬毎レース情報
    Public Const ID_RACE_UMA As String = "SE"

    ' 競走馬マスタ
    Public Const ID_UMA As String = "UM"

    ' 騎手マスタ
    Public Const ID_KISYU As String = "KS"

    ' 特別登録馬
    Public Const ID_TOKU As String = "TK"

    ' 払戻
    Public Const ID_HARAI As String = "HR"

    ' 票数1
    Public Const ID_HYOSU As String = "H1"

    ' 票数6(3連単)
    Public Const ID_HYOSU_SANRENTAN As String = "H6"

    ' オッズ1(単複枠)
    Public Const ID_ODDS_TANPUKU As String = "O1"

    ' オッズ2(馬連)
    Public Const ID_ODDS_UMAREN As String = "O2"

    ' オッズ3(ワイド)
    Public Const ID_ODDS_WIDE As String = "O3"

    ' オッズ4(馬単)
    Public Const ID_ODDS_UMATAN As String = "O4"

    ' オッズ5(3連複)
    Public Const ID_ODDS_SANREN As String = "O5"

    ' オッズ6(3連単)
    Public Const ID_ODDS_SANRENTAN As String = "O6"

    ' 調教師マスタ
    Public Const ID_CHOKYO As String = "CH"

    ' 生産者マスタ
    Public Const ID_SEISAN As String = "BR"

    ' 馬主マスタ
    Public Const ID_BANUSI As String = "BN"

    ' 繁殖馬マスタ
    Public Const ID_HANSYOKU As String = "HN"

    ' 産駒マスタ
    Public Const ID_SANKU As String = "SK"

    ' レコードマスタ
    Public Const ID_RECORD As String = "RC"

    ' 坂路調教
    Public Const ID_HANRO As String = "HC"

    ' 馬体重
    Public Const ID_BATAIJYU As String = "WH"

    ' 天候馬場状態
    Public Const ID_TENKO_BABA As String = "WE"

    ' 出走取消・競走除外
    Public Const ID_TORIKESI_JYOGAI As String = "AV"

    ' 騎手変更
    Public Const ID_KISYU_CHANGE As String = "JC"

    ' データマイニング予想
    Public Const ID_MINING As String = "DM"

    ' 開催スケジュール
    Public Const ID_SCHEDULE As String = "YS"

    ' 発走時刻変更
    Public Const ID_HASSOU_JIKOKU_CHANGE As String = "TC"

    ' コース変更
    Public Const ID_COURSE_CHANGE As String = "CC"


    ' -----JV-Linkステータス-----

    ' JVRead - 正常読み込み
    Public Const ST_READ_SUCCESS As Integer = 0

    ' JVRead - エラー 
    Public Const ST_READ_ERR As Integer = -2

    ' JVRead - ファイルリスト読み込み終了 
    Public Const ST_READ_EOL As Integer = 0

    ' JVRead - ファイルの区切れ 
    Public Const ST_READ_EOF As Integer = -1

    ' JVRead - ダウンロード中 
    Public Const ST_READ_DOWNLOAD_NOW As Integer = -3


    ' -----コード変換-----

    ' 競馬場コード
    Public Const CV_JO_CD As String = "2001"

    ' 曜日コード
    Public Const CV_WD_CD As String = "2002"

    ' グレードコード
    Public Const CV_GR_CD As String = "2003"

    ' 競走種別コード
    Public Const CV_RS_CD As String = "2005"

    ' 競走記号コード
    Public Const CV_RK_CD As String = "2006"

    ' 競走条件コード
    Public Const CV_RJ_CD As String = "2007"

    ' 重量種別コード
    Public Const CV_WH_CD As String = "2008"

    ' トラックコード
    Public Const CV_TR_CD As String = "2009"

    ' 馬場状態コード
    Public Const CV_BC_CD As String = "2010"

    ' 天候コード
    Public Const CV_WE_CD As String = "2011"

    ' 異常区分コード
    Public Const CV_IR_CD As String = "2101"

    ' 着差コード
    Public Const CV_TS_CD As String = "2102"

    ' 品種コード
    Public Const CV_HS_CD As String = "2201"

    ' 性別コード
    Public Const CV_SX_CD As String = "2202"

    ' 毛色コード
    Public Const CV_FC_CD As String = "2203"

    ' 馬記号コード
    Public Const CV_UK_CD As String = "2204"

    ' 東西所属コード
    Public Const CV_TZ_CD As String = "2301"

    ' 騎乗資格コード
    Public Const CV_KQ_CD As String = "2302"

    ' 騎手見習コード
    Public Const CV_KM_CD As String = "2303"


    ' -----データ区分-----

    ' 出走馬名表(木曜)
    Public Const KB_THU As String = "1"

    ' 出馬表(金・土曜)
    Public Const KB_FRI As String = "2"

    ' 成績速報(3着まで確定)
    Public Const KB_S3 As String = "3"

    ' 成績速報(5着まで確定)
    Public Const KB_S5 As String = "4"

    ' 成績速報(全馬着順確定)
    Public Const KB_SALL As String = "5"

    ' 成績速報(全馬着順+ｺｰﾅｰ通過順)
    Public Const KB_SCOR As String = "6"

    ' 成績(月曜)
    Public Const KB_MON As String = "7"

    ' 地方競馬
    Public Const KB_LKL As String = "A"

    ' 海外国際レース
    Public Const KB_FOR As String = "B"

    ' レース中止
    Public Const KB_CAN As String = "9"

    ' 該当データ削除
    Public Const KB_DEL As String = "0"


End Module