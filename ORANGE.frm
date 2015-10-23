VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 MyApp"
   ClientHeight    =   3555
   ClientLeft      =   3480
   ClientTop       =   2265
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "ORANGE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "系统信息(&S)..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   0
      Top             =   3075
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "张毅 金达祺"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblCompany 
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      Picture         =   "ORANGE.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      Caption         =   "ALL RUN"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本 1.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告：本软件属于ORANGE公司所有，不得盗用       否则属于侵权行为，我公司有权上述。       "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   3630
   End
   Begin VB.Label lblDescription 
      Caption         =   "本软件可运行多种程序，有着良好的实用性。本软件包括看图、抓图、浏览文本，可观看AVI、MOV等文件，同时可浏览WAV、MID等文件。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 注册键安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册键 ROOT 类型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode 以 Null 结尾的字符串
Const REG_DWORD = 4                      ' 32-位数字

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblCompany.Caption = App.CompanyName
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 试图从注册表得到系统信息程序路径\名称...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' 试图从注册表得到系统信息程序路径...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 验证已知 32 位文件版本的存在
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' 错误 - 文件未找到...
        Else
            GoTo SysInfoErr
        End If
    ' 错误 - 注册项未找到...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "此时系统信息无效", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' 循环指针
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 打开的注册键的句柄
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 注册键的数据类型
    Dim tmpVal As String                                    ' 注册键的临时存储区
    Dim KeyValSize As Long                                  ' 注册键变量的大小
    '------------------------------------------------------------
    ' 在根键 {HKEY_LOCAL_MACHINE...} 下打开注册键
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册键
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 句柄错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量大小
    
    '------------------------------------------------------------
    ' 检索注册键值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 获得/创建键值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 句柄错误
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 添加以 Null 结尾的字符串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null 找到，从字符串提取
    Else                                                    ' WinNT 不需要以 Null 结束字符串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null 未找到， 仅提取字符串
    End If
    '------------------------------------------------------------
    ' 为了转换而决定键值类型..
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜索数据类型...
    Case REG_SZ                                             ' 字符串型注册键数据类型
        KeyVal = tmpVal                                     ' 复制字符串值
    Case REG_DWORD                                          ' 双字型注册键数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 转换每一位
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地建立值
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 转换双字型为字符串型
    End Select
    
    GetKeyValue = True                                      ' 返回成功
    rc = RegCloseKey(hKey)                                  ' 关闭注册键
    Exit Function                                           ' 退出
    
GetKeyError:      ' 发生错误后清除...
    KeyVal = ""                                             ' 设置返回值为空字符串
    GetKeyValue = False                                     ' 返回失败
    rc = RegCloseKey(hKey)                                  ' 关闭注册键
End Function

