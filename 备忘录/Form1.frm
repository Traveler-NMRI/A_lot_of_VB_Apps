VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "备忘录"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2950
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "请在这里输入内容..."
      Top             =   0
      Width           =   4580
      _ExtentX        =   8070
      _ExtentY        =   5212
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Dim iscontrol As Boolean

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
End Sub

Private Sub Form_Resize()
RichTextBox1.Width = Me.Width - 220
RichTextBox1.Height = Me.Height - 560
End Sub
