VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Folder"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   3795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
   End
   Begin ComctlLib.ProgressBar Bar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   18
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label percent 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Select a folder to find files in:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim strFolder As String
    Dim intImages As Integer
    Dim strFile As String
    Dim strImageFile() As String
    strFolder = Dir1.Path
    strFolder = strFolder & IIf(Right$(strFolder, 1) = "\", "", "\")
    strFile = Dir$(strFolder & "*.jpg")
    If strFile <> "" Then
        intImages = 1
    End If
    Do While strFile <> ""
        strFile = Dir$()
        If strFile <> "" Then
            intImages = intImages + 1
            ReDim Preserve strImageFile(intImages)
            strImageFile(intImages) = strFile
        End If
    Loop
    If intImages >= 18 Then
        percent.Visible = True
        Command1.Visible = False
        Bar1.Visible = True
        Form1.Show
    Else
        MsgBox ("Not Enough Images " & "Only " & intImages & " files.")
    End If
End Sub


Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
