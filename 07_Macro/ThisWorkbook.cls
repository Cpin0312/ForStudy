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
'【プロシージャ名】Workbook_Open()ワークブックOpenイベント
'【概　要】メニューバーを追加
'【引　数】なし
'【戻り値】なし
'==========================================================

    Dim cbrCmd As CommandBar
    Dim cbcMenu As CommandBarControl

    'Tb_Posget

    Set cbrCmd = Application.CommandBars("Worksheet Menu Bar")

    '--- Mod Start S.Iwanaga 2010/04/13
    Set cbcMenu = cbrCmd.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbcMenu.Caption = "SPC SQL作成"


    Set cbcMenu2 = cbcMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbcMenu2.Caption = "初期ページ設定(&P)"

    With cbcMenu2.Controls.Add(Type:=msoControlButton)
        .Caption = "初期ページ作成(&I)"
        .OnAction = "CreateInitialPage"
    End With

    With cbcMenu2.Controls.Add(Type:=msoControlButton)
        .Caption = "テーブルページ追加(&N)"
        .OnAction = "NewInitialPage"
    End With
    
    Set cbcMenu3 = cbcMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbcMenu3.Caption = "初期データ(&I)"

    With cbcMenu3.Controls.Add(Type:=msoControlButton)
        .Caption = "SQL 出力(&O)"
        .OnAction = "CreateInitialData"
        .FaceId = 317
    End With

    Set cbcMenu99 = cbcMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
    cbcMenu99.Caption = "バージョン(&V)"
    cbcMenu99.OnAction = "ShowVersion"
    cbcMenu99.BeginGroup = True

    '--- Mod End
    Set cbrCmd = Nothing
    Set cbcMenu = Nothing

    Exit Sub

End Sub


