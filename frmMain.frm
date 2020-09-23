VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registry and INI Viewer"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   ".REG > .INI"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Load INI"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2040
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   2460
      TabIndex        =   6
      Top             =   3480
      Width           =   2115
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Write Setting"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Setting"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1485
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Selected Key"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Selected Section"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Keys"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sections"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Text1.Text = INIReadWrite.Read_Ini(Text2.Text, Text4.Text, "<NONE>")
End Sub

Private Sub Command2_Click()
    INIReadWrite.Write_Ini Text2.Text, Text4.Text, Text1.Text
End Sub

Private Sub Command5_Click()
Dim tmp As String, EndFile As String
    Dialog.DialogTitle = "Find .REG"
    Dialog.Filter = "Registry Files|*.REG"
    Dialog.ShowOpen
    If Dialog.FileName = "" Then Exit Sub
    Open Dialog.FileName For Input As #1
        Line Input #1, tmp
        Line Input #1, tmp
        Do Until EOF(1)
            Line Input #1, tmp
            tmp = Replace(tmp, "\""", "-=*chr(34)")
            tmp = Replace(tmp, """", "")
            tmp = Replace(tmp, "-=*chr(34)", """")
            If InStr(1, tmp, "=") > 42 Or InStr(1, tmp, "=") = 0 And Left(tmp, 1) <> "[" And Right(tmp, 1) <> "]" And Right(EndFile, 1) <> "]" Then
                EndFile = EndFile & "%%&&Chr(13)&&%%" & tmp
            ElseIf Left(tmp, 1) = "[" And Right(tmp, 1) = "]" And Right(EndFile, 15) = "%%&&Chr(13)&&%%" Then
                EndFile = Left(EndFile, Len(EndFile) - 13) & vbCrLf & vbCrLf & tmp
            Else
                EndFile = EndFile & vbCrLf & tmp
            End If
            Me.Caption = "Registry Messer Upper - Processed " & cnt & " entries"
            cnt = cnt + 1
            DoEvents
        Loop
    Close #1
    Dialog.DialogTitle = "Save As"
    Dialog.Filter = "INI Files|*.ini"
    Dialog.ShowSave
    If Dialog.FileName = "" Then Exit Sub
    If Dir(Dialog.FileName) <> "" Then Kill Dialog.FileName
    Open Dialog.FileName For Output As #1
        Print #1, EndFile
    Close #1
End Sub

Private Sub Command6_Click()
    Load_INI
End Sub

Private Sub Form_Load()
    Load_INI
End Sub

Private Sub Load_INI()
    Dialog.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
    Dialog.DialogTitle = "Load INI / Registry File"
    Dialog.Filter = "INI Files|*.ini|Registry Files|*.reg"
    Dialog.ShowOpen
    If Dialog.FileName = "" Then Exit Sub
    INIReadWrite.INISetup Dialog.FileName, 3500
    Dim tmp() As String
    Dim tmps As String
    tmps = INIReadWrite.Read_Sections
    tmp = Split(tmps, Chr(0))
    List1.Clear
    For x = 0 To UBound(tmp)
        List1.AddItem tmp(x)
    Next x
    If List1.ListCount > 0 Then List1.ListIndex = 0
End Sub

Private Sub List1_Click()
    Text2.Text = List1.List(List1.ListIndex)
    Dim tmp() As String
    Dim tmps As String
    tmps = INIReadWrite.Read_Keys(Text2.Text)
    tmp = Split(tmps, Chr(0))
    List2.Clear
    For x = 0 To UBound(tmp)
        List2.AddItem tmp(x)
    Next x
End Sub

Private Sub List2_Click()
    Text4.Text = List2.List(List2.ListIndex)
    Call Command1_Click
End Sub
