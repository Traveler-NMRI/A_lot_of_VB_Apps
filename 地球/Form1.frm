VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "地球"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1125
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   1125
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   600
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":381A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "地球"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Form1.frx":3C6C
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Dim aaa As String
Dim timesss  As Integer
Dim hour As Integer

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
aaa = Now
hour = DatePart("h", aaa)
If (hour >= 17) And (hour <= 19) And timesss <> 6 Then
MsgBox "傍晚好！", , "Tips"
timesss = 6
ElseIf (hour >= 20) And (hour <= 23) And timesss <> 7 Then
MsgBox "晚上好！", , "Tips"
timesss = 7
ElseIf (hour >= 0) And (hour <= 4) And timesss <> 0 Then
MsgBox "午夜好！", , "Tips"
timesss = 0
ElseIf (hour = 5) And timesss <> 1 Then
MsgBox "凌晨好！", , "Tips"
timesss = 1
ElseIf (hour >= 6) And (hour <= 7) And timesss <> 2 Then
MsgBox "早上好！", , "Tips"
timesss = 2
ElseIf (hour >= 8) And (hour <= 11) And timesss <> 3 Then
MsgBox "上午好！", , "Tips"
timesss = 3
ElseIf (hour >= 12) And (hour <= 13) And timesss <> 4 Then
MsgBox "中午好！", , "Tips"
timesss = 4
ElseIf (hour >= 14) And (hour <= 16) And timesss <> 5 Then
MsgBox "下午好！", , "Tips"
timesss = 5
End If
End Sub

Private Sub Timer1_Timer()
aaa = Now
hour = DatePart("h", aaa)
If (hour >= 17) And (hour <= 19) And timesss <> 6 Then
MsgBox "傍晚好！", , "Tips"
timesss = 6
ElseIf (hour >= 20) And (hour <= 23) And timesss <> 7 Then
MsgBox "晚上好！", , "Tips"
timesss = 7
ElseIf (hour >= 0) And (hour <= 4) And timesss <> 0 Then
MsgBox "午夜好！", , "Tips"
timesss = 0
ElseIf (hour = 5) And timesss <> 1 Then
MsgBox "凌晨好！", , "Tips"
timesss = 1
ElseIf (hour >= 6) And (hour <= 7) And timesss <> 2 Then
MsgBox "早上好！", , "Tips"
timesss = 2
ElseIf (hour >= 8) And (hour <= 11) And timesss <> 3 Then
MsgBox "上午好！", , "Tips"
timesss = 3
ElseIf (hour >= 12) And (hour <= 13) And timesss <> 4 Then
MsgBox "中午好！", , "Tips"
timesss = 4
ElseIf (hour >= 14) And (hour <= 16) And timesss <> 5 Then
MsgBox "下午好！", , "Tips"
timesss = 5
End If
End Sub
