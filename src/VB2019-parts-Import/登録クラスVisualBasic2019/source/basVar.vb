Option Strict On
Option Explicit On
Module basVar
    '========================================================================
    '  JRA-VAN Data Lab.プログラミングパーツ「Public変数定義ファイル」
    '
    '
    '   作成: JRA-VAN ソフトウェア工房  2003年6月3日
    '
    '========================================================================
    '   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
    '========================================================================


    Public gCon As ADODB.Connection ''コネクション変数
    Public SS As String 'mdb用SQLカラム名サポート("[")
    Public SE As String 'mdb用SQLカラム名サポート("]")
End Module