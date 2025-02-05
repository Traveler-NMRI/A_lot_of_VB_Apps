VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3765
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
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
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   735
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
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Text            =   "5"
      Top             =   480
      Width           =   735
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
      Height          =   405
      Left            =   1080
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   0
      Width           =   2655
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
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "照片速度"
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
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "图片目录"
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
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Dim a As String
Dim b As String
Private Sub Command2_Click()
WriteINI "Settings", "FilePath", Text1.Text, ".\Settings.ini"
WriteINI "Settings", "Times", Text2.Text, ".\Settings.ini"
Form1.ReStart
Me.Hide
End Sub

Private Sub Command3_Click()
Text1.Text = a
Text2.Text = b
Form1.Timer1.Enabled = True
Me.Hide
End Sub

Public Function Shows()
Me.Show
a = ReadINI("Settings", "FilePath", ".\Settings.ini", "C:\")
b = ReadINI("Settings", "Times", ".\Settings.ini", "5")
Text1.Text = a
Text2.Text = b
End Function
' 读取INI文件
Public Function ReadINI(ByVal section As String, ByVal key As String, ByVal filePath As String, Optional ByVal defaultValue As String = "") As String
    Dim returnValue As String
    Dim bufSize As Long
    bufSize = 2048 ' 根据需要设定缓冲区大小
    returnValue = String(bufSize, 0)
    GetPrivateProfileString section, key, defaultValue, returnValue, bufSize, filePath
    ReadINI = Left$(returnValue, InStr(returnValue, Chr$(0)) - 1)
End Function
 
' 写入INI文件
Public Function WriteINI(ByVal section As String, ByVal key As String, ByVal value As String, ByVal filePath As String) As Boolean
    WritePrivateProfileString section, key, value, filePath
    WriteINI = (LenB(value) > 0)
End Function

Private Sub Text2_Change()
If Val(Text2.Text) > 60 Then
Text2.Text = 60
End If
If Val(Text2.Text) <= 0 Then
Text2.Text = 1
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    ' 允许输入数字（0-9）、Backspace和退格键
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' 阻止非数字字符的输入
    End If
End Sub

