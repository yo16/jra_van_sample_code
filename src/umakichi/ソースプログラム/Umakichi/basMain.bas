Attribute VB_Name = "basMain"
'
'   �N�����W���[��
'
'   �������̃��[�e�B���e�B�[Function���܂�
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' Assert���[�h �G���[���O�����������_�Œ�~����ꍇ 1
Public ASSERTMODE As Long

Public gApp As clsApp               '' �A�v���P�[�V�����I�u�W�F�N�g
Public gCC As clsCodeConverter      '' �R�[�h�ϊ��I�u�W�F�N�g
Public gSC As clsStringConverter    '' ������ϊ��I�u�W�F�N�g

Public gJVLinkSID As String ' Main�֐���Exe�w�b�_���琶�����܂��B

Public gDebugCounter_clsGridData As Long
Public gDebugCounter_clsGridItem As Long

Public gColDarkBG As Long
Public gColBG     As Long

Public gstrMDBName(0 To 49) As String '' ���ڐڑ��p
Public gCallStack As New Collection

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���萔
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Const cAppName As String = "�n�g�I�[�v���\�[�X�� for DataLab."
Public Const cHelpFileName As String = "Umakichi.chm"

' �F�萔
Public Const ColorMother As Long = &HEEEEFF     '��n��i�����j
Public Const ColorFather As Long = &HFFFFFF     '���n��i�����j
Public Const ColorODBack0 As Long = &HFFFFFF    '���i�I�b�Y�j
Public Const ColorODBack1 As Long = &H10101     '���i�I�b�Y�j
Public Const ColorODFore0 As Long = &H1FFFF     '���i�I�b�Y�j
Public Const ColorODForeH As Long = &H101FF     '�ԁi�I�b�Y�j
Public Const ColorODForeM As Long = &HFF0101    '�i�I�b�Y�j
Public Const ColorODForeL As Long = &H10101     '���i�I�b�Y�j
Public Const ColorLinkExist As Long = &HFF0101  '�i�S�����N�j
Public Const ColorLinked As Long = &HFF00FF     '�s���N�i�S�����N�j

' ���̑��萔
Public Const cRegistrySubKey As String = "Umakichi5"
Public Const cFromtimeFN As String = "Fromtime.dat"                 ' Fromtime�ۑ��t�@�C����
Public Const cFromtimeThisWeekFN As String = "FromtimeThisWeek.dat" ' FromtimeThisWeek�ۑ��t�@�C����

Public Const ASCII_ZERO  As Byte = 48
Public Const ASCII_TWO   As Byte = 50
Public Const ASCII_SEVEN As Byte = 55


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �N��
'
'   ���l: ��VB�v���W�F�N�g�̊J�n�v���V�[�W��
'
Public Sub Main()
    Dim splashWindow As frmSplash
    
    ASSERTMODE = 0
    
    ' SID�𐶐�
    With App
    gJVLinkSID = "Umakichi/OpenSource"
    End With
    
    If App.PrevInstance Then
        End
    End If
    
    ' �X�v���b�V���E�C���h�E�ŃR�s�[���C�g�̕\��
    Set splashWindow = New frmSplash
    splashWindow.Show
    splashWindow.Refresh
    
    Call init
    
    ' �A�v���P�[�V�����I�u�W�F�N�g�̐���
    Set gApp = New clsApp
    Set gCC = New clsCodeConverter
    Set gSC = New clsStringConverter
    ' �N��
    gApp.start
    
    ' �X�v���b�V���E�C���h�E��j������
    splashWindow.kill
End Sub


'
'   �@�\: �������̂����A�傫���ق��̒l��Ԃ�
'
'   ���l: �Ȃ�
'
Public Function Bigger(a As Long, b As Long) As Long
    Bigger = IIf(a > b, a, b)
End Function


'
'   �@�\: �������̂����A�������ق��̒l��Ԃ�
'
'   ���l: �Ȃ�
'
Public Function Smaller(a As Long, b As Long) As Long
    Smaller = IIf(a < b, a, b)
End Function


'
'   �@�\: �����񒆂̘A�������󔒂�" "�ɒu��������
'
'   ���l: �Ȃ�
'
Public Function ContractSpace(str As String) As String
    Dim i    As Long
    Dim p    As Long
    Dim out  As String
    Dim flag As Boolean
    Dim c    As String
    
    p = 1
    flag = False
    For i = 1 To Len(str)
        c = Mid$(str, i, 1)
        If c = " " Or c = "�@" Then
            If Not flag Then
                out = out & IIf(p = 1, "", " ") & Mid$(str, p, i - p)
                flag = True
            End If
        Else
            If flag Then
                p = i
                flag = False
            End If
        End If
    Next i
    
    If Not flag Then
        out = out & " " & Mid$(str, p, i - p)
    Else
        out = out & " "
    End If
    
    ContractSpace = out
End Function


'
'   �@�\: ���K�\���ŃX�y�[�X���폜
'
'   ���l: �Ȃ�
'
Public Function DelSpace(strString As String) As String
On Error GoTo Errorhandler
    Dim rx As New RegExp
    With rx
        .Global = True
        .Pattern = "\s|�@"
        DelSpace = .Replace(strString, "")
    End With
    Exit Function
Errorhandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   �@�\: �O���C�X�P�[�����Q�l������
'
'   ���l: �Ȃ�
'
Public Function Contrast(color As Long) As Long
    Dim r As Long
    Dim G As Long
    Dim b As Long
    Dim Gray As Long
    
    r = color Mod 256
    G = (color \ 256) Mod 256
    b = (color \ 65536) Mod 256
    
    Gray = 0.2126 * r ^ 2.2 + 0.7152 * G ^ 2.2 + 0.0724 * b ^ 2.2
    Gray = Gray ^ (1 / 2.2)
    
    If Gray < 128 Then
        Contrast = &HFFFFFF
    Else
        Contrast = &H0
    End If
End Function


'
'   �@�\: ���R�[�h�Z�b�g���J��
'
'   ���l: �Ȃ�
'
Public Function OpenTableDirect(rs As ADODB.Recordset, cn As ADODB.Connection, TableName As String) As Boolean
On Error GoTo Errorhandler
    rs.CursorLocation = adUseServer
    rs.Index = "PrimaryKey"
    rs.Open TableName, cn, adOpenKeyset, adLockReadOnly, adCmdTableDirect
    OpenTableDirect = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    OpenTableDirect = False
End Function


'
'   �@�\: �R�l�N�V�������J������
'
'   ���l: �Ȃ�
'
Public Sub freecn(cn As ADODB.Connection)

    If Not cn Is Nothing Then
        Do While cn.State And adStateExecuting
            Call cn.Cancel
            gApp.Log "freecn Cancel"
        Loop
        Do While cn.State And adStateOpen
            cn.Close
            gApp.Log "freers Close"
        Loop
        Set cn = Nothing
    Else
        gApp.Log "freecn Nothing"
    End If
End Sub


'
'   �@�\: ���R�[�h�Z�b�g�����L�����Z�����[�e�B���e�B�[
'
'   ���l: �Ȃ�
'
Public Sub freers(rs As ADODB.Recordset)
    If Not rs Is Nothing Then
        Do While rs.State And adStateExecuting
            Call rs.Cancel
            
            gApp.Log "freers Cancel"
        Loop
        Do While rs.State And adStateOpen
            rs.Close
            gApp.Log "freers Close"
        Loop
        Set rs = Nothing
    Else
        gApp.Log "freers Nothing"
    End If
End Sub


'
'   �@�\: ���S��Seek����
'
'   ���l: ���ʂ�Seek����ƁASeek�o���Ă��Ȃ��ꍇ�������
'
Public Sub SafeSeek(ByRef rs As ADODB.Recordset, ByRef Fields As Variant, ByRef Values As Variant)
On Error GoTo Errorhandler
    Dim i As Long
    Dim c As Long
    Dim NG As Boolean

    If rs.EOF And rs.BOF Then
        Exit Sub
    End If

    rs.MoveFirst
    
    Do
        rs.Seek Values
        If rs.EOF Or rs.BOF Then
            Exit Do
        End If
        
        NG = False
        For i = 0 To UBound(Fields)
            NG = NG Or (rs(Fields(i)) <> Values(i))
        Next i
        
        If Not NG Then
            Exit Do
        End If
        
        gApp.Log "SafeSeek"
        For i = 0 To UBound(Fields)
            gApp.Log c & ":: " & Fields(i) & " : " & rs(Fields(i)) & " <-> " & Values(i)
        Next i
        c = c + 1
        If c > 10 Then
            gApp.Log "SafaSeek failed"
            Exit Do
        End If
    Loop
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    gApp.Log "SafeSeek Error"
    Resume Next
End Sub


'
'   �@�\: "&"��"&&"�ɒu������
'
'   ���l: ���x����"&"���������"_"�ɂȂ��
'
Public Function ReplaceAmpersand(str As String) As String
    ReplaceAmpersand = Replace(str, "&", "&&")
End Function


'
'   �@�\: �\�[�g�̂��߂ɋ󔒂�"��"�ɒu��������
'
'   ���l: �Ȃ�
'
Public Function FormatForSort(str As String) As String
    str = ContractSpace(str)
    If str = Space(1) Then
        str = "�K"          ' "�K" is Unicode's last ���� according to value
    Else
        str = Trim$(str)
    End If
    FormatForSort = str
End Function


'
'   �@�\: ���R�[�h�Z�b�g����l���擾����
'
'   ���l: �Ȃ�
'
Public Function IfExist(rs As ADODB.Recordset, FieldName As String) As String
    If Not rs.EOF Then
        If Not IsNull(rs(FieldName).value) Then
            IfExist = rs(FieldName)
        End If
    End If
End Function


'
'   �@�\: ��f�[�^(�����l)�̔��f
'
'   ���l: " ","0"�����̃f�[�^��"":�f�[�^�Ȃ�
'
Public Function IfBe(str As String) As String
    If Space(Len(str)) = str Then
        IfBe = ""
    ElseIf String$(Len(str), "0") = str Then
        IfBe = ""
    Else
        IfBe = str
    End If
End Function


'
'   �@�\: �o�C�g�z��ɒl��}������
'
'   ���l: �Ȃ�
'
Public Sub ByteInsert(ByRef b() As Byte, pos As Long, width As Long, val() As Byte)
    Dim i As Long
    For i = 0 To width - 1
        If i <= UBound(val) Then
            b(pos + i) = val(i)
        End If
    Next i
End Sub


'
'   �@�\: JVOpen�̃G���[���b�Z�[�W��ϊ�����
'
'   ���l: �Ȃ�
'
Public Function ErrMsgJVOpen(lngRet As Long) As String
    Select Case lngRet
    Case 0
        ErrMsgJVOpen = "����" & vbCrLf & ""
    Case -1
        ErrMsgJVOpen = "�Y���f�[�^����" & vbCrLf & "�w�肳�ꂽ�p�����[�^�ɍ��v����V�����f�[�^���T�[�o�[�ɑ��݂��Ȃ�����ͤ�ŐV�o�[�W���������J���꤃��[�U�[���ŐV�o�[�W�����̃_�E�����[�h��I�����܂����JVClose���Ăяo���Ď�荞�ݏ������I�����Ă��������"
    Case -2
        ErrMsgJVOpen = "�Z�b�g�A�b�v�_�C�A���O�ŃL�����Z���������ꂽ" & vbCrLf & "�Z�b�g�A�b�v�p�f�[�^�̎�荞�ݎ��Ƀ��[�U�[���_�C�A���O�ŃL�����Z���������܂����JVClose���Ăяo���Ď�荞�ݏ������I�����Ă�������� "
    Case -111
        ErrMsgJVOpen = "dataspec�p�����[�^���s��" & vbCrLf & "�p�����[�^�̓n�������p�����[�^�̓��e�ɖ�肪����Ǝv���܂���T���v���v���O���������Q�Ƃ���������p�����[�^��JV -Link�ɓn���Ă��邩�m�F���Ă�������� "
    Case -112
        ErrMsgJVOpen = "fromdate�p�����[�^���s��" & vbCrLf & "�p�����[�^�̓n�������p�����[�^�̓��e�ɖ�肪����Ǝv���܂���T���v���v���O���������Q�Ƃ���������p�����[�^��JV -Link�ɓn���Ă��邩�m�F���Ă�������� "
    Case -114
        ErrMsgJVOpen = "key�p�����[�^���s��" & vbCrLf & "�p�����[�^�̓n�������p�����[�^�̓��e�ɖ�肪����Ǝv���܂���T���v���v���O���������Q�Ƃ���������p�����[�^��JV -Link�ɓn���Ă��邩�m�F���Ă�������� "
    Case -115
        ErrMsgJVOpen = "option�p�����[�^���s��" & vbCrLf & "�p�����[�^�̓n�������p�����[�^�̓��e�ɖ�肪����Ǝv���܂���T���v���v���O���������Q�Ƃ���������p�����[�^��JV -Link�ɓn���Ă��邩�m�F���Ă�������� "
    Case -116
        ErrMsgJVOpen = "dataspec��option�̑g�ݍ��킹���s��" & vbCrLf & "�p�����[�^�̓n�������p�����[�^�̓��e�ɖ�肪����Ǝv���܂���T���v���v���O���������Q�Ƃ���������p�����[�^��JV -Link�ɓn���Ă��邩�m�F���Ă�������� "
    Case -201
        ErrMsgJVOpen = "�i�u�h���������s�Ȃ��Ă��Ȃ�" & vbCrLf & "JVOpen/JVRTOpen�ɐ旧����JVInit���Ă΂�Ă��Ȃ��Ǝv���܂���K��JVInit���ɌĂяo���Ă�������� "
    Case -202
        ErrMsgJVOpen = "�O���JVOpen/JVRTOpen�ɑ΂���JVClose���Ă΂�Ă��Ȃ��i�I�[�v�����j" & vbCrLf & "�O��Ăяo����JVOpen/JVRTOpen��JVClose�ɂ���ăN���[�Y����Ă��Ȃ��Ǝv���܂��JVOpen/JVRTOpen���Ăяo������͎��ɌĂяo���܂ł̊Ԃ�JVClose��K���Ăяo���Ă�������� "
    Case -211
        ErrMsgJVOpen = "���W�X�g�����e���s���i���W�X�g�����e���s���ɕύX���ꂽ�j" & vbCrLf & "JV-Link�̓��W�X�g���ɒl���Z�b�g����ۂɒl�̃`�F�b�N���s���܂��i�Ⴆ�΃T�[�r�X�L�[�̌����Ȃǁj���A���W�X�g������l��ǂݏo���Ďg�p����ۂɖ�肪��������Ƃ��̃G���[���������܂�����W�X�g�������ڏ���������ꂽ�Ȃǂ̏󋵂��l�����Ȃ��ꍇ�ɂ�JRA-VAN�ւ��A�����������B"
    Case -301
        ErrMsgJVOpen = "�F�؃G���[" & vbCrLf & "�T�[�r�X�L�[���������Ȃ��B���邢�͕����̃}�V���œ���T�[�r�X�L�[���g�p�����ꍇ�ɔ������܂��B�����̃}�V���œ����T�[�r�X�L�[�����悤�����ꍇ�ɂ́A���̃G���[�����������}�V����JV-Link���A���C���X�g�[�����A�ăC���X�g�[����A���p�L�[�̍Ĕ��s���K�v�ƂȂ�܂��B"
    Case -302
        ErrMsgJVOpen = "�T�[�r�X�L�[�̗L�������؂�" & vbCrLf & "Data Lab.�T�[�r�X�̗L���������؂�Ă��܂��B�T�[�r�X���̎�����������~���Ă���Ǝv���܂����������ɂ̓T�[�r�X���̍čw�����K�v�ł�� "
    Case -303
        ErrMsgJVOpen = "�T�[�r�X�L�[���ݒ肳��Ă��Ȃ��i�T�[�r�X�L�[����l�j" & vbCrLf & "�T�[�r�X�L�[��ݒ肵�Ă��Ȃ��Ǝv���܂��BJVLink�C���X�g�[������̓T�[�r�X�L�[����Ȃ̂ŕK���ݒ肷��K�v������܂�� "
    Case -401
        ErrMsgJVOpen = "JV-Link�����G���[" & vbCrLf & "JV-Link�����ŃG���[�����������Ǝv���܂��BJRAVAN�ւ��A����������� "
    Case -411
        ErrMsgJVOpen = "�T�[�o�[�G���[�i HTTP �X�e�[�^�X404NotFount�j" & vbCrLf & "���W�X�g�������ڕύX���ꂽ���AData Lab.�p�T�[�o�[�ɖ�肪���������Ǝv���܂��JRA -VAN�̃����e�i���X���łȂ��ꍇ�Ť���̃G���[�������ꍇ��JRA-VAN�ւ��A�����������B"
    Case -412
        ErrMsgJVOpen = "�T�[�o�[�G���[�i HTTP �X�e�[�^�X403Forbidden�j" & vbCrLf & "Data Lab.�p�T�[�o�[�ɖ�肪���������Ǝv���܂�����̃G���[�������ꍇ��JRA -VAN�ւ��A����������� "
    Case -413
        ErrMsgJVOpen = "�T�[�o�[�G���[�iHTTP�X�e�[�^�X200,403,404�ȊO�j" & vbCrLf & "Data Lab.�p�T�[�o�[�ɖ�肪���������Ǝv���܂�����̃G���[�������ꍇ��JRA -VAN�ւ��A����������� "
    Case -421
        ErrMsgJVOpen = "�T�[�o�[�G���[�i�T�[�o�[�̉������s���j" & vbCrLf & "Data Lab.�p�T�[�o�[�ɖ�肪���������Ǝv���܂�����̃G���[�������ꍇ��JRA -VAN�ւ��A����������� "
    Case -431
        ErrMsgJVOpen = "�T�[�o�[�G���[�i�T�[�o�[�A�v���P�[�V���������G���[�j" & vbCrLf & "Data Lab.�p�T�[�o�[�ɖ�肪���������Ǝv���܂�����̃G���[�������ꍇ��JRA -VAN�ւ��A����������� "
    Case -501
        ErrMsgJVOpen = "�Z�b�g�A�b�v�����ɂ����Ăb�c�|�q�n�l������" & vbCrLf & "JRA-VAN���񋟂���������CD-ROM���Z�b�g���Ă��Ȃ��Ǝv���܂��������CD -ROM���Z�b�g���Ă�������� "
    Case -504 '�ǉ�
        ErrMsgJVOpen = "�T�[�o�[�����e�i���X��" & vbCrLf & "�T�[�o�[�������e�i���X���ł��B"
    Case Else
        ErrMsgJVOpen = "�z��O�̃G���[���������܂����B" & vbCrLf & ""
    End Select
End Function



'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �f�[�^�x�[�X��������
'
'   ���l: �Ȃ�
'
Private Sub init()
    gstrMDBName(0) = "subBANUSI.mdb"
    gstrMDBName(1) = "subBATAIJYU.mdb"
    gstrMDBName(2) = "subCHOKYO.mdb"
    gstrMDBName(3) = "subCHOKYO_SEISEKI.mdb"
    gstrMDBName(4) = "subHANRO.mdb"
    gstrMDBName(5) = "subHANSYOKU.mdb"
    gstrMDBName(6) = "subHARAI.mdb"
    gstrMDBName(7) = "subKISHU.mdb"
    gstrMDBName(8) = "subKISHU_CHANGE.mdb"
    gstrMDBName(9) = "subKISHU_SEISEKI.mdb"
    gstrMDBName(10) = "subMINING.mdb"
    gstrMDBName(11) = "subODDS_SANREN0.mdb"
    gstrMDBName(12) = "subODDS_SANREN1.mdb"
    gstrMDBName(13) = "subODDS_SANREN2.mdb"
    gstrMDBName(14) = "subODDS_SANREN3.mdb"
    gstrMDBName(15) = "subODDS_SANREN4.mdb"
    gstrMDBName(16) = "subODDS_SANREN5.mdb"
    gstrMDBName(17) = "subODDS_SANREN6.mdb"
    gstrMDBName(18) = "subODDS_SANREN7.mdb"
    gstrMDBName(19) = "subODDS_SANREN8.mdb"
    gstrMDBName(20) = "subODDS_SANREN9.mdb"
    gstrMDBName(21) = "subODDS_TANPUKUWAKU.mdb"
    gstrMDBName(22) = "subODDS_UMAREN.mdb"
    gstrMDBName(23) = "subODDS_UMATAN0.mdb"
    gstrMDBName(24) = "subODDS_UMATAN1.mdb"
    gstrMDBName(25) = "subODDS_UMATAN2.mdb"
    gstrMDBName(26) = "subODDS_UMATAN3.mdb"
    gstrMDBName(27) = "subODDS_UMATAN4.mdb"
    gstrMDBName(28) = "subODDS_UMATAN5.mdb"
    gstrMDBName(29) = "subODDS_UMATAN6.mdb"
    gstrMDBName(30) = "subODDS_UMATAN7.mdb"
    gstrMDBName(31) = "subODDS_UMATAN8.mdb"
    gstrMDBName(32) = "subODDS_UMATAN9.mdb"
    gstrMDBName(33) = "subODDS_WIDE.mdb"
    gstrMDBName(34) = "subRACE.mdb"
    gstrMDBName(35) = "subRECORD.mdb"
    gstrMDBName(36) = "subSANKU.mdb"
    gstrMDBName(37) = "subSCHEDULE.mdb"
    gstrMDBName(38) = "subSEISAN.mdb"
    gstrMDBName(39) = "subTENKO_BABA.mdb"
    gstrMDBName(40) = "subTOKU.mdb"
    gstrMDBName(41) = "subTOKU_RACE.mdb"
    gstrMDBName(42) = "subTORIKESI_JYOGAI.mdb"
    gstrMDBName(43) = "subUMA.mdb"
    gstrMDBName(44) = "subUMA_RACE_A.mdb"
    gstrMDBName(45) = "subUMA_RACE_B.mdb"
    gstrMDBName(46) = "LinkTables.mdb"
    gstrMDBName(47) = "subRAKaiSel.mdb"
    gstrMDBName(48) = "subHASSOU_CHANGE.mdb"
    gstrMDBName(49) = "subCOURSE_CHANGE.mdb"
End Sub


