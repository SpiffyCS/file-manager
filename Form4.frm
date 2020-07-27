VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   3615
   ClientLeft      =   2475
   ClientTop       =   6105
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4680
   Begin VB.CommandButton Search 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton DeleteFiles 
      Caption         =   "Delete files"
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton CopyFiles 
      Caption         =   "Copy files"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "*.txt"
      Top             =   600
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "File format (e.g. *.txt)"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Choose drive"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim files As FileSystemObject
Dim filetype As String
Dim drive As String

Private Sub copyfiles_Click()
Form5.Show
End Sub

Private Sub DeleteFiles_Click()
'file1contents = Drive1 + Text1
Kill file1contents
End Sub

Private Sub Drive1_Change()
drive = Drive1
driveextfrom = Drive1
End Sub

Private Sub File1_Click()
File1 = Drive1 + Text1
'file1contents = File1
End Sub

Private Sub Text1_Change()
filetype = Text1
End Sub
