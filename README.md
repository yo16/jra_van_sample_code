# jra_van_sample_code
JRA-VANのプログラミングパーツ・開発支援ツール提供コーナーのコード

- [競馬ソフト開発 プログラミングパーツ提供コーナー｜競馬ソフト使い放題の会員サービス DataLab.（データラボ）｜競馬情報ならJRA-VAN](https://jra-van.jp/dlb/sdv/pgm.html)
- 以下のファイルは、上記ホームページよりダウンロードし、説明は、上記のホームページより引用
- ※ 2024/9/16時点

# ファイル
- 馬吉ソース公開版
    - 過去、弊社が作成・公開していた競馬ソフト「馬吉 for Datalab.（以下、馬吉）」のソースを公開いたします。
    - 競馬ソフトを作成する際の参考プログラムとしてご利用下さい。

    - ソース
        - Microsoft VisualBasic 6（2009/4/13）
            - `./src/umakichi`

    - 馬吉は Microsoft VisualBasic 6.0(SP6) Professional Editionを使用して開発されています。
    - 一部のコントロールは、Learning Editionに含まれない開発ライセンスのコントロールを使用しております。
    - Visual Basic .NETでは動作いたしません。コンバージョン作業が必要になります。（Visual Basicアップグレードウィザードのみでは動作いたしません。)
    - Visual Basic .NETでは動作いたしません。プログラムコンバージョンが必要になります。
    - [馬吉ソース公開版のライセンスについてはこちら](https://jra-van.jp/dlb/sdv/umakichi_license.html)
    - 馬吉 for Data Lab.のソース公開にあたって[NPO法人オープンソースソフトウェア協会](http://www.ossaj.org/)にご協力いただきました。
    - 【不具合修正2009年4月13日】frmDBUpdate.frmファイルの700行目に未知のレコード種別の場合DBをOpenしない処理を追加。これまでは（else文のように）単純に全てのレコード種別に対してOpenDBを行っていた。

- データベース作成クラス
    - データベース作成クラスは、JV-Dataを利用したソフトウェアのデータ参照元となるデータベースを管理するためのプログラミングパーツです。データベースを新規作成、登録済みデータの削除、最適化の操作を行う際に使用します。
    - 作成されるデータベースの仕様は同梱のデータベース仕様書を参照下さい。

    - ソース
        - Microsoft Visual Basic 2019 (2023/8/8)
            - `./src/VB2019-Builder`
        - Microsoft Visual C++ 2019 (2023/8/8)
            - `./src/VC2019-Builder`

- JV-Data登録クラス
    - JV-Data登録クラスは、JV-Linkのメソッド“JVRead”で読み込んだJV-Dataレコードをデータベースに登録するためのプログラミングパーツです。このクラスは、同梱されているデータベースファイル（Data.mdb)をデータベースとして使用します。また、パーツとして提供しているデータベース作成クラスを使ってデータベースを作成することも可能です。

    - ソース
        - Microsoft Visual Basic 2019 （2024/8/7）
            - `./src/VB2019-parts-Import`
        - Microsoft Visual C++ 2019 （2024/8/7）
            - `./src/VC2019-parts-Import`

- コード変換クラス
    - 「コード変換クラス」は、JV-Dataで扱われているコード値（競走場、競走記号コード等）を対応するコード名称に変換するためのプログラミングパーツです。コード表の参照には、CSV形式のテキストファイルを使用しています。
    - コードファイルのフォーマット仕様は、プログラミングパーツのデータベース仕様書を参照して下さい。
    - コード仕様は、JV-Data仕様書を参照して下さい。

    - ソース
        - Microsoft Visual Basic 2019 （2021/5/26）
            - `./src/VB2019-parts-CodeConv`
        - Microsoft Visual C++ 2019 （2021/5/26）
            - `./src/VC2019-parts-CodeConv`

- データベースファイル
    - JV-Dataを使用する競馬ソフトで自由に利用できるJV-Data対応標準データベースファイル。
    - データベース作成クラスを使用することで、同様のデータベースを作成することが可能です。
    - 作成されるデータベースの仕様は[JRA-VAN SDK](https://jra-van.jp/dlb/sdv/sdk.html)同梱のJV-Data仕様書を参照下さい。

    - データ（テーブル定義のみ）
        - Microsoft Access (mdb形式) （2021/5/26）
            - `./data/Data-mdb`
        - Microsoft Access (accdb形式) （2023/8/8）
            - `./data/Data-accdb`

- サンプルプログラム
    - ここではプログラミングパーツとして提供されている各パーツの利用方法や、パーツを応用して作成したサンプルプログラムを提供しています。

    - プログラム
        - 出馬表サンプル
            - Microsoft Visual Basic 2019 （2023/8/8）
                - `./src/VB2019-sample-denma`
        - 出走別着度数処理用サンプル
            - Microsoft Visual Basic 2019 （2023/8/8）
                - `./src/VB2019-parts-CKImport`

- VBAサンプルプログラム
    - VBA(VisualBasic for Application)とは、Microsoft社の Word,Excel,Access 等のソフトに搭載されているマクロ言語です。ここでは、サンプルプログラムとして、JV-Dataのデータ種別「レース詳細」データより一部のデータを取込み、表示出力するサンプルプログラムを提供しています
    
    - プログラム
        - Microsoft Access 2019 （2023/8/21）
            - `./src/Access2019-parts-sample`
        - Microsoft Excel 2019 （2023/8/21）
            - `./src/Excel2019-parts-sample`

- 開発支援ツール提供
    - ソース
        - DataLab.検証ツール(Ver.2.6.0) (2023/8/8）
            - `./src/JVDataCheckToolVer2.6.0`

