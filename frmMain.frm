VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Orange"
   ClientHeight    =   750
   ClientLeft      =   -1785
   ClientTop       =   2820
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   327680
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件"
      Begin VB.Menu mnuOpen 
         Caption         =   "打开..."
      End
      Begin VB.Menu mnuClose 
         Caption         =   "关闭"
      End
      Begin VB.Menu mnuFenGe2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "设置"
      Begin VB.Menu mnuWaveSetting 
         Caption         =   "Wave"
      End
      Begin VB.Menu mnuMIDISetting 
         Caption         =   "MIDI"
      End
      Begin VB.Menu mnuAVISetting 
         Caption         =   "AVI"
      End
      Begin VB.Menu mnuCDSetting 
         Caption         =   "CD"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助"
      Begin VB.Menu mnuFenGe1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
' Set properties needed by MCI to open.
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    mnuWaveSetting.Checked = True
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuAVISetting_Click()
    MMControl1.Command = "Close"
    MMControl1.DeviceType = "AVIVideo"
    MMControl1.filename = ""
    movedesignsetting mnuAVISetting

End Sub

Private Sub mnuCDSetting_Click()
    MMControl1.Command = "Close"
    MMControl1.DeviceType = "CDAudio"
    MMControl1.filename = ""
    movedesignsetting mnuCDSetting
End Sub

Private Sub mnuClose_Click()
    MMControl1.Command = "Close"
End Sub

Private Sub mnuExit_Click()
    MMControl1.Command = "Close"
    End
End Sub

Private Sub mnuMIDISetting_Click()
    MMControl1.Command = "Close"
    MMControl1.DeviceType = "Sequencer"
    movedesignsetting mnuMIDISetting
End Sub

Private Sub mnuOpen_Click()
        ' Set CancelError is True
    MMControl1.Command = "Close"
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    Select Case MMControl1.DeviceType
        Case "WaveAudio"
            CommonDialog1.Filter = "Wave Files" & _
                "(*.wav)|*.wav|All Files (*.*)|*.*"
        Case "Sequencer"
            CommonDialog1.Filter = "MIDI Files" & _
                "(*.mid,*.rmi)|*.mid;*.rmi|All Files(*.*)|*.*"
        Case "CDAudio"
            MMControl1.Command = "Open"
            Exit Sub
        Case "AVIVideo"
            CommonDialog1.Filter = "AVI Files" & _
                "(*.avi)|*.avi|All Files(*.*)|*.*"
    End Select
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    MMControl1.filename = CommonDialog1.filename
    ' Open the MCI WaveAudio device.
    MMControl1.Command = "Open"

    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub



Private Sub mnuWaveSetting_Click()
    MMControl1.Command = "Close"
    MMControl1.DeviceType = "WaveAudio"
    'mnuWaveSetting.Checked = True
    'mnuMIDISetting.Checked = False
    'mnuCDSetting.Checked = False
    'mnuAVISetting.Checked = False
    movedesignsetting mnuWaveSetting
End Sub


