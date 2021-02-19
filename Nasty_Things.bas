Attribute VB_Name = "Nasty_Things"

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Declare Function GetLogicalDriveStrings Lib "Kernel32" Alias _
  "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal _
  lpBuffer As String) As Long
Declare Function GetDriveType Lib "Kernel32" Alias _
  "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetSystemDirectoryA Lib "Kernel32" _
   (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer


Global Const SW_HIDE = 0
Global Const SW_NORMAL = 1
Global Const SW_MAXIMIZE = 3
Global Const SW_MINIMIZE = 6
Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4
Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200

Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Long
     fAnyOperationsAborted As Long
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Enum eFileAttribute
    ATTR_READONLY = &H1
    ATTR_HIDDEN = &H2
    ATTR_SYSTEM = &H4
    ATTR_DIRECTORY = &H10
    ATTR_ARCHIVE = &H20
    ATTR_NORMAL = &H80
    ATTR_TEMPORARY = &H100
End Enum


Private Declare Function GetFileAttributes Lib "Kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long






Public Function AddBackslash(s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function
'
'  Returns the system directory.
'
Public Function GetSystemDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetSystemDirectoryA("", 0)
   s = Space(i)
   Call GetSystemDirectoryA(s, i)
   GetSystemDirectory = AddBackslash(Left$(s, i - 1))
End Function
Function ShowDriveType(drvpath) As String
    Dim fs, d, s, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    Select Case d.DriveType
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    s = t
    ShowDriveType = s
End Function

Sub DOShell(sShellString As String, iWinType As Integer)
Dim iInstanceHandle As Integer, X As Integer
On Error Resume Next
iInstanceHandle = Shell(sShellString, iWinType)
On Error Resume Next
End Sub
Public Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum
    FileExists = IIf(Err, False, True)
    Close intFileNum

    Err = 0
End Function
Sub Get_User_Name()

                ' Dimension variables
                Dim lpBuff As String * 25
                Dim ret As Long, UserName As String

                ' Get the user name minus any trailing spaces found in the name.
                ret = GetUserName(lpBuff, 25)
                UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

                ' Display the User Name
                'FrmFreg.ur = UserName
End Sub






            
Function Shella(Program As String, Optional ShowCmd As Long = vbNormalNoFocus, Optional ByVal WorkDir As Variant) As Long

    Dim FirstSpace As Integer, Slash As Integer


    If Left(Program, 1) = """" Then
        FirstSpace = InStr(2, Program, """")


        If FirstSpace <> 0 Then
            Program = Mid(Program, 2, FirstSpace - 2) & Mid(Program, FirstSpace + 1)
            FirstSpace = FirstSpace - 1
        End If

    Else
        FirstSpace = InStr(Program, " ")
    End If

    If FirstSpace = 0 Then FirstSpace = Len(Program) + 1


    If IsMissing(WorkDir) Then


        For Slash = FirstSpace - 1 To 1 Step -1
            If Mid(Program, Slash, 1) = "\" Then Exit For
        Next



        If Slash = 0 Then
            WorkDir = CurDir
        ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
            WorkDir = Left(Program, Slash)
        Else
            WorkDir = Left(Program, Slash - 1)
        End If

    End If

    Shella = ShellExecute(0, vbNullString, _
    Left(Program, FirstSpace - 1), LTrim(Mid(Program, FirstSpace)), _
    WorkDir, ShowCmd)
    If Shella < 32 Then VBA.Shell Program, ShowCmd 'To raise Error
End Function















