VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zharudar 1.9 Components"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5400
      Top             =   120
   End
   Begin ZharuComp.XandersXPProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   13405804
      Scrolling       =   3
      Value           =   62
   End
   Begin VB.FileListBox filelist 
      Height          =   4185
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin MSComctlLib.ListView files 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Location"
         Object.Width           =   4621
      EndProperty
   End
   Begin VB.Label txt 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   4695
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Dim IngSuccess As Long
Dim v As Integer
Dim sys As String

Private Sub get_OCX()
filelist.Pattern = "*.ocx"  '.............................Set File filters to *.OCX
filelist.Path = App.Path    '.............................Set File List Path
If filelist.ListCount <> 0 Then pb.Max = filelist.ListCount
pb.Value = 0
Me.Caption = "Collecting *.ocx files..."

    For v = 0 To filelist.ListCount
        If filelist.List(v) <> "" Then
            With files.ListItems.Add
                .Text = filelist.List(v)
                .SubItems(1) = "OCX"
                .SubItems(2) = App.Path & "\" & filelist.List(v)
            End With
            pb.Value = pb.Value + 1
            Sleep 250
        End If
        txt = files.ListItems.Count & " Itms added."
    Next v

filelist.Pattern = "*.dll"  '.............................Set File filters to *.OCX
filelist.Path = App.Path    '.............................Set File List Path
If filelist.ListCount <> 0 Then pb.Max = filelist.ListCount
pb.Value = 0
Me.Caption = "Collecting *.dll files..."

    For v = 0 To filelist.ListCount
        If filelist.List(v) <> "" Then
            With files.ListItems.Add
                .Text = filelist.List(v)
                .SubItems(1) = "DLL"
                .SubItems(2) = App.Path & "\" & filelist.List(v)
            End With
            pb.Value = pb.Value + 1
            Sleep 250
        End If
        txt = files.ListItems.Count & " Items added."
    Next v
    
Me.Caption = "Collected " & files.ListItems.Count & " Files"
files.Refresh
If files.ListItems.Count <> 0 Then
    txt = "Processing Files..."
    txt.Refresh
    reg_them
    pb.Value = 0
Else
    MsgBox "No item founded to process!", vbExclamation, "Error"
    End
End If
End Sub

Private Sub reg_them()
Dim cmd As String
On Error Resume Next
sys = GetSystemDirectory

Sleep 1000
Me.Caption = "Registering Collected *.ocx & *.dll ..."

pb.Max = files.ListItems.Count
pb.Value = 0

    For v = 0 To files.ListItems.Count
    If files.ListItems.Item(v) <> 0 Then
            Me.Caption = "Handeling file " & files.ListItems.Item(v) & " ..."
            pb.Value = pb.Value + 1
            files.ListItems.Item(v).Selected = True
            Sleep 100
            cmd = ""
            files.Refresh
            cmd = sys & "regsvr32.exe -s" & " " & Chr(34) & App.Path & "\" & files.ListItems.Item(v) & Chr(34)
            If cmd <> "" Then
                Sleep 250
                DOShell cmd, vbHide '........................................calling regsvr32.exe
            End If
    End If
    Next v
txt = "Finished"
txt.Refresh
MsgBox "Processing completed!", vbInformation, "Done"
End
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
txt = "Collecting files..."
pb.Value = 0
tmr.Enabled = True
End Sub

Private Sub tmr_Timer()
get_OCX
tmr.Enabled = False
End Sub

