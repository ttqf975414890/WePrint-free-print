VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00FAF7F7&
   BorderStyle     =   0  'None
   Caption         =   "WePrint-free-print"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Exit 
      BackColor       =   &H00FAF7F7&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1560
   End
   Begin VB.CommandButton Start 
      BackColor       =   &H00FAF7F7&
      Caption         =   "����"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   3360
   End
   Begin VB.Label ˵�� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ѧϰ��;�����ڱ���� 10 ������ɾ��������������ҵ��;"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   5160
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���빹�����ڣ�2018-12-08��V0.9"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   3480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����΢����Ѵ�ӡ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1800
      TabIndex        =   0
      Top             =   420
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   960
      Picture         =   "MainForm.frx":10DC2
      Stretch         =   -1  'True
      Top             =   300
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00302020&
      FillStyle       =   0  'Solid
      Height          =   1800
      Left            =   0
      Top             =   0
      Width           =   5400
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: ttqf

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const CS_DROPSHADOW = &H20000
Private Const GCL_STYLE = (-26)

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const Auhor = "ttqf"

Private Sub Exit_Click()
    ˵��.Caption = "���ڻָ���ؽ���": DoEvents
    ShellBlocked "cmd.exe /c net start PrinterSpoolMonitor", vbHide
    End
End Sub

Private Sub Form_Load()
    SetClassLong Me.hwnd, GCL_STYLE, GetClassLong(Me.hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Start_Click()
    Dim curTime As Long, regString As String, regTime As Long, timeDiff As Long
    regString = GetSetting("WePrint-free-print", "0", "StartTrialTime", "Null")
    curTime = DateDiff("s", "2018-12-08 00:00:00", Now)
    If regString = "Null" Then
        SaveSetting "WePrint-free-print", "0", "StartTrialTime", curTime
        regString = curTime
    End If
    regTime = Val(regString)
    If curTime - regTime <= 600 Then
        ˵��.Caption = "���ڽ�����ؽ��� (1/3)": DoEvents
        ShellBlocked "taskkill /f /im PrinterSpoolMonitor.exe", vbHide
        ˵��.Caption = "���ڽ�����ӡ���� (2/3)": DoEvents
        ShellBlocked "cmd.exe /c net stop Spooler", vbHide
        ˵��.Caption = "���ڻָ���ӡ���� (3/3)": DoEvents
        ShellBlocked "cmd.exe /c net start Spooler", vbHide
        ˵��.Caption = "����ѧϰ��;�����ڱ���� 10 ������ɾ��������������ҵ��;": DoEvents
    Else
        MsgBox "���ѳ�ʱʹ�á��벻Ҫ���ó���������ҵ��;��лл��"
    End If
End Sub

Private Sub ShellBlocked(Path As String, Mode As VbAppWinStyle)
    Dim pID As Long
    Dim pHwnd As Long
    
    pID = Shell(Path, Mode)
    pHwnd = OpenProcess(SYNCHRONIZE, 0, pID)
    If pHwnd <> 0 Then
        WaitForSingleObject pHwnd, INFINITE
        CloseHandle pHwnd
    Else
        Debug.Print "error"
    End If
End Sub

Private Sub ˵��_DblClick()
    On Error Resume Next
    DeleteSetting "WePrint-free-print", "0", "StartTrialTime"
    MsgBox "ע�����Ϣ�������"
End Sub
