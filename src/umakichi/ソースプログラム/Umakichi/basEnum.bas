Attribute VB_Name = "basEnum"
'
'   Enum�錾���W���[��
'

Option Explicit


'
'   �@�\: JV-Link �擾���[�h
'
'   ���l: �Ȃ�
'
Public Enum ukJVLMode
    ukjUsual
    ukjThisWeek
End Enum


'
'   �@�\: ����擾���[�h
'
'   ���l: �Ȃ�
'
Public Enum ukPromptMode
    ukpRA
    ukpOD
    ukpPALLET
End Enum


'
'   �@�\: �f�[�^�G�N�X�|�[�g���[�h
'
'   ���l: �Ȃ�
'
Public Enum ukExportMode
    ukeJVDATA
    ukeCSV
End Enum


'
'   �@�\: �f�[�^�G�N�X�|�[�g�t�B���^���[�h
'
'   ���l: �Ȃ�
'
Public Enum ukExportFilter
    ukfNone
    ukfDate
    ukfJyoCD
    ukfDateJyoCD
End Enum


'
'   �@�\: ctlPane ���[�h
'
'   ���l: �Ȃ�
'
Public Enum ukCtlPaneMode
    ukcpNowFetching
    ukcpNoData
    ukcpShowControls
    ukcpHideControls
End Enum

