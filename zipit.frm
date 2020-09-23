VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZIPit"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "ZipIT"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   5715
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ZIP - it"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Frame Frame2 
         Caption         =   "Pick Files "
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5775
         Begin VB.FileListBox File1 
            Height          =   4185
            Left            =   2760
            MultiSelect     =   2  'Extended
            System          =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   2895
         End
         Begin VB.DirListBox Dir1 
            Height          =   3690
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   2535
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2535
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Save File "
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   4920
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fileselected
Private Function CompressFiles(theFiles, outputfile)
 Set javaObject = GetObject("java:ZipFunctions")
 strResult = javaObject.ZipFile(theFiles, outputfile)
 Set javaObject = Nothing
 CompressFiles = strResult
End Function

Private Sub Command2_Click()
For cnt = 0 To File1.ListCount - 1
    If File1.Selected(cnt) Then
        If Trim(fileselected) = "" Then
            fileselected = Dir1.Path & "\" & File1.List(cnt)
        Else
            fileselected = fileselected & "*" & Dir1.Path & "\" & File1.List(cnt)
            'Debug.Print fileselected
        End If
    End If
Next
If Trim(fileselected) <> "" And Trim(Text1.Text) <> "" Then
    Command2.Enabled = False
    Form1.MousePointer = 13
    status.SimpleText = "Compressing Files"
    Debug.Print "Compress Start: " & Time
    Debug.Print CompressFiles(fileselected, Text1.Text)
    Debug.Print "Compress End: " & Time
    Form1.MousePointer = 1
    Command2.Enabled = True
    status.SimpleText = "Done"
    Text1.Text = ""
  
Else
    MsgBox "Please select files to compress as well as destination file !!"
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
fileselected = ""
Text1.Text = Dir1.Path & "\Zip-it-File" & Replace(Replace(Time, ":", ""), " ", "") & ".zip"
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
fileselected = ""
End Sub

 
Private Sub Form_Load()
Text1.Text = Dir1.Path & "\Zip-it-File" & Replace(Replace(Time, ":", ""), " ", "") & ".zip"
End Sub

