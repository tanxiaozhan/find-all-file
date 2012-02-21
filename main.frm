VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15945
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   15945
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   9420
      Left            =   240
      TabIndex        =   2
      Top             =   1230
      Width           =   14730
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   720
      Left            =   14520
      TabIndex        =   1
      Top             =   255
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "main.frx":0000
      Top             =   270
      Width           =   13665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'常量
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const BIF_RETURNONLYFSDIRS = 1
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
    '定义类 (用于查找文件)
Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
'定义类 (用于浏览文件夹窗口)
Private Type BrowseInfo
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type

    '查找第一个文件的API
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    '查找下一个文件的API
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    '获取文件属性的API
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
    '关闭查找文件的API
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    '以下为调用浏览文件夹窗口的API
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

  '  自定义函数
Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function
    '自定义函数
Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, _
DirCount As Integer)
Dim FileName As String     '文件名
Dim DirName As String      '子目录名
Dim dirNames() As String   '目录数组
Dim nDir As Integer        '当前路径的目录数
Dim i As Integer           '循环计数器变量
Dim hSearch As Long        '搜索句柄变量
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
If Right(path, 1) <> "\" Then path = path & "\"
'搜索子目录
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*/", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
If (DirName <> "./") And (DirName <> "../") Then
If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD)     '获取下一个子目录
Loop
Cont = FindClose(hSearch)
End If
     '遍历目录并累计文件总数
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> "./") And (FileName <> "../") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
List1.AddItem path & FileName
End If
Cont = FindNextFile(hSearch, WFD)   '　 获取下一个文件
Wend
Cont = FindClose(hSearch)
End If
'如果子目录存在则遍历之
If nDir > 0 Then
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(i) & "\", _
SearchStr, FileCount, DirCount)
Next i
End If
End Function
    '查找按钮代码
Sub Command1_Click()
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo
With udtBI
    '设置浏览窗口
.hWndOwner = Me.hWnd
'返回选中的目录
.ulFlags = BIF_RETURNONLYFSDIRS
End With
    '调出浏览窗口
lpIDList = SHBrowseForFolder(udtBI)
If lpIDList Then
sPath = String$(MAX_PATH, 0)
   ' 获取路径
SHGetPathFromIDList lpIDList, sPath
    '释放内存
CoTaskMemFree lpIDList
iNull = InStr(sPath, vbNullChar)
If iNull Then
sPath = Left$(sPath, iNull - 1)
End If
End If
Screen.MousePointer = vbHourglass
List1.Clear
SearchPath = sPath     '选中的目录为搜索的起始路径
FindStr = "*.*"     '搜索所有类型的文件(此处可另作定义)
FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
Text1.Text = "查找到的文件数：" & NumFiles & vbCrLf & "查找的目录数：" & _
NumDirs + 1 & vbCrLf & "文件大小总共为：" & vbCrLf & _
Format(FileSize, "#,###,###,##0") & "字节"
Screen.MousePointer = vbDefault
End Sub


