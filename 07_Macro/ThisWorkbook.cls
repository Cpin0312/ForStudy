VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
'==========================================================
'�y�v���V�[�W�����zWorkbook_Open()���[�N�u�b�NOpen�C�x���g
'�y�T�@�v�z���j���[�o�[��ǉ�
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

    Dim cbrCmd As CommandBar
    Dim cbcMenu As CommandBarControl

    'Tb_Posget

    Set cbrCmd = Application.CommandBars("Worksheet Menu Bar")

    '--- Mod Start S.Iwanaga 2010/04/13
    Set cbcMenu = cbrCmd.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbcMenu.Caption = "SPC SQL�쐬"


    Set cbcMenu2 = cbcMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbcMenu2.Caption = "�����y�[�W�ݒ�(&P)"

    With cbcMenu2.Controls.Add(Type:=msoControlButton)
        .Caption = "�����y�[�W�쐬(&I)"
        .OnAction = "CreateInitialPage"
    End With

    With cbcMenu2.Controls.Add(Type:=msoControlButton)
        .Caption = "�e�[�u���y�[�W�ǉ�(&N)"
        .OnAction = "NewInitialPage"
    End With
    
    Set cbcMenu3 = cbcMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbcMenu3.Caption = "�����f�[�^(&I)"

    With cbcMenu3.Controls.Add(Type:=msoControlButton)
        .Caption = "SQL �o��(&O)"
        .OnAction = "CreateInitialData"
        .FaceId = 317
    End With

    Set cbcMenu99 = cbcMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    cbcMenu99.Caption = "�o�[�W����(&V)"
    cbcMenu99.OnAction = "ShowVersion"
    cbcMenu99.BeginGroup = True

    '--- Mod End
    Set cbrCmd = Nothing
    Set cbcMenu = Nothing

    Exit Sub

End Sub

