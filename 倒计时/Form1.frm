VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "倒计时"
   ClientHeight    =   360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleWidth      =   3975
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Text            =   "0"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "时"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Private Sub Command1_Click()
Timer1.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
If Val(Text3.Text) = 0 Then
If Val(Text2.Text) = 0 Then
If Val(Text1.Text) = 0 Then
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
MsgBox "时间到！！！", , "Tips"
Else
Text1.Text = Val(Text1.Text) - 1
Text2.Text = 59
Text3.Text = 59
End If
Else
Text2.Text = Val(Text2.Text) - 1
Text3.Text = 59
End If
Else
Text3.Text = Val(Text3.Text) - 1
End If
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
End Sub

Private Sub Text3_Change()
If Val(Text3.Text) > 60 Then
Text3.Text = 60
End If
End Sub
Private Sub Text2_Change()
If Val(Text2.Text) > 60 Then
Text2.Text = 60
End If
End Sub
Private Sub Text1_Change()
If Val(Text1.Text) > 65535 Then
Text1.Text = 65535
End If
End Sub

Private Sub Timer1_Timer()
If Val(Text3.Text) = 0 Then
If Val(Text2.Text) = 0 Then
If Val(Text1.Text) = 0 Then
MsgBox "时间到！！！", , "Tips"
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Else
Text1.Text = Val(Text1.Text) - 1
Text2.Text = 59
Text3.Text = 59
End If
Else
Text2.Text = Val(Text2.Text) - 1
Text3.Text = 59
End If
Else
Text3.Text = Val(Text3.Text) - 1
End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' 允许输入数字（0-9）、Backspace和退格键
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' 阻止非数字字符的输入
    End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
    ' 允许输入数字（0-9）、Backspace和退格键
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' 阻止非数字字符的输入
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    ' 允许输入数字（0-9）、Backspace和退格键
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' 阻止非数字字符的输入
    End If
End Sub
