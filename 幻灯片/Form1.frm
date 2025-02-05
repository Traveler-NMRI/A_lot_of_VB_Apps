VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "幻灯片"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1560
      Top             =   720
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Dim c As String
Dim i As Integer
Dim d As Integer
Dim imagelist() As String
Dim len_list As Integer

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
ReStart
End Sub

Private Sub Form_Resize()
Image1.Width = Me.Width
Image1.Height = Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()
Form2.Shows
Timer1.Enabled = False
End Sub

Public Function Search(Path As String)
Dim fso As Object
Dim folder As Object
Dim file As Object
Dim folderPath As String
len_list = 0
Erase imagelist
' 创建FileSystemObject实例
Set fso = CreateObject("Scripting.FileSystemObject")
 
' 设置需要遍历的目录路径
folderPath = Path
 
' 获取目录对象
Set folder = fso.GetFolder(folderPath)
 
' 遍历目录中的文件
For Each file In folder.Files
    ' 检查文件扩展名是否为.txt
    If LCase(fso.GetExtensionName(file.Path)) = "jpg" Then
        ' 输出文件路径
        len_list = len_list + 1
        ReDim Preserve imagelist(len_list) As String
        imagelist(len_list - 1) = file.Path
    End If
Next
For Each file In folder.Files
    ' 检查文件扩展名是否为.txt
    If LCase(fso.GetExtensionName(file.Path)) = "bmp" Then
        ' 输出文件路径
        len_list = len_list + 1
        ReDim Preserve imagelist(len_list) As String
        imagelist(len_list - 1) = file.Path
    End If
Next
' 清理
Set file = Nothing
Set folder = Nothing
Set fso = Nothing
End Function

Private Sub Timer1_Timer()
If len_list = 0 Then
Timer1.Enabled = False
Exit Sub
End If
If i = len_list Then
i = 0
End If
Image1.Picture = LoadPicture(imagelist(i))
i = i + 1
End Sub

Public Function ReStart()
c = Form2.ReadINI("Settings", "FilePath", ".\Settings.ini", "C:\")
d = Val(Form2.ReadINI("Settings", "Times", ".\Settings.ini", "5"))
Timer1.Interval = d * 1000
i = 0
If Dir(c, vbDirectory) <> "" Then
Search (c)
End If
Timer1.Enabled = True
If len_list = 0 Then
Timer1.Enabled = False
Exit Function
End If
If i = len_list Then
i = 0
End If
Image1.Picture = LoadPicture(imagelist(i))
i = i + 1
End Function
