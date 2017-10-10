VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7215
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check5 
      Caption         =   "置顶"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "挂起父进程"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "挂起进程"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CheckBox Check4 
      Caption         =   "自动"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "自动"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Text            =   "Form1.frx":000D
      Top             =   2400
      Width           =   6975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "自动凝固"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "启动抓捕行动"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6720
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "挂起需要UAC允许/管理员身份"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    If Check1.Value Then
        Hook Me.hWnd
    Else
        UnHook Me.hWnd
        Text1.BackColor = &H80000005
        Form1.Command1.Enabled = False
        Form1.Command2.Enabled = False
    End If
End Sub



Private Sub Check2_Click()
    If Check2.Value = 0 Then Text1.BackColor = &H80000005
End Sub

Private Sub Check5_Click()
    Dim lrtn As Long
    Timer2.Enabled = (Check5.Value = 1)
    lrtn = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
End Sub

Private Sub Command1_Click()
    SuspendProcess (gblProcessId)
End Sub

Private Sub Command2_Click()
    SuspendProcess (gblInheritedPID)
End Sub

Private Sub Form_Terminate()
    If Check1.Value Then
        UnHook Me.hWnd
    End If
End Sub

Private Sub Timer1_Timer()
    Text2.Text = GetForegroundWindowInfo
End Sub

Private Sub Timer2_Timer()
    Dim lrtn As Long
    lrtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
End Sub
