Attribute VB_Name = "MainCode"
Option Explicit


Sub Main()
    frmSplash.Show
End Sub




Sub MoveDesignSetting(MenuName As Object)
    With frmMain
    .mnuWaveSetting.Checked = False
    .mnuMIDISetting.Checked = False
    .mnuCDSetting.Checked = False
    .mnuAVISetting.Checked = False
    End With
    MenuName.Checked = True
End Sub


Sub ToolBarDesignSetting(Index As Integer)
    Dim Counter As Integer
    With frmMain.Toolbar1.Buttons
        For Counter = 1 To 4
            .Item(Counter).Value = tbrUnpressed
        Next Counter
        .Item(Index).Value = tbrPressed
    End With
    
End Sub


