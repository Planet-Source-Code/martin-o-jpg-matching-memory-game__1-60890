VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matching Game"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   35
      Left            =   7560
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   34
      Left            =   6120
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   33
      Left            =   4680
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   32
      Left            =   3240
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   31
      Left            =   1800
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   30
      Left            =   360
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   29
      Left            =   7560
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   28
      Left            =   6120
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   27
      Left            =   4680
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   26
      Left            =   3240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   25
      Left            =   1800
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   24
      Left            =   360
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   23
      Left            =   7560
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   22
      Left            =   6120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   21
      Left            =   4680
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   20
      Left            =   3240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   19
      Left            =   1800
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   18
      Left            =   360
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   17
      Left            =   7560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   16
      Left            =   6120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   15
      Left            =   4680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   14
      Left            =   3240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   13
      Left            =   1800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   12
      Left            =   360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   11
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   10
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   9
      Left            =   4680
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   8
      Left            =   3240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   7
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   6
      Left            =   360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   5
      Left            =   7560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   4
      Left            =   6120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   3
      Left            =   4680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton button 
      Height          =   1455
      Index           =   0
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   36
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   35
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   34
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   33
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   32
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   31
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   30
      Left            =   360
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   29
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   28
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   27
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   26
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   25
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   24
      Left            =   360
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   23
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   22
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   21
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   20
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   19
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   18
      Left            =   360
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   17
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   16
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   15
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   14
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   13
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   12
      Left            =   360
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   11
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   10
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   9
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   8
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   7
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   6
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   5
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   4
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   3
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   2
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   1
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   360
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clicked As Integer
Private partner(35) As Integer
Private time As Integer
Option Explicit

Private Sub button_Click(Index As Integer)
If Timer1.Enabled = False Then
    button(Index).Visible = False
    If clicked = 1000 Then
        clicked = Index
    Else
        Label1.Caption = Label1.Caption + 1
        If clicked = partner(Index) Then
            clicked = 1000
        Else
            Timer1.Enabled = True
            Do
                DoEvents
            Loop Until Timer1.Enabled = False
            button(Index).Visible = True
            button(clicked).Visible = True
            clicked = 1000
        End If
    End If
 End If
End Sub

Private Sub Form_Load()
    Randomize Timer
    clicked = 1000
    Dim strImageFile() As String
    Dim intImages As Integer
    Dim booUsedImage() As Boolean
    Dim strFolder As String
    Dim strFile As String
    Dim number1 As Integer
    Dim number2 As Integer
    Dim goOn As Boolean
    Dim assigned(35) As Boolean
    Dim i As Integer
    strFolder = Form2.Dir1.Path
    strFile = " "
    strFolder = strFolder & IIf(Right$(strFolder, 1) = "\", "", "\")
    strFile = Dir$(strFolder & "*.jpg")
    intImages = 1
    ReDim Preserve strImageFile(intImages)
    strImageFile(1) = strFile
    Do While strFile <> ""
        strFile = Dir$()
        If strFile <> "" Then
            intImages = intImages + 1
            ReDim Preserve strImageFile(intImages)
            strImageFile(intImages) = strFile
        End If
    Loop
    
    If intImages < 18 Then
        MsgBox "Not Enough Images"
        Form2.Show
        Me.Hide
    Else
        ReDim booUsedImage(intImages)
        For i = 1 To 18
            goOn = False
            Do
                number1 = Rnd * 35
                If assigned(number1) = False Then
                    assigned(number1) = True
                    goOn = True
                End If
            Loop Until goOn = True
        
            goOn = False
            Do
            number2 = Rnd * 35
            If assigned(number2) = False Then
                    assigned(number2) = True
                    goOn = True
                End If
            Loop Until goOn = True
            partner(number1) = number2
            partner(number2) = number1
            goOn = False
            Dim intPic
            Do
                intPic = Int(Rnd * intImages) + 1
                If booUsedImage(intPic) = False Then goOn = True
            Loop Until goOn = True
            booUsedImage(intPic) = True
            Image1(number1).Picture = LoadPicture(strFolder & strImageFile(intPic))
            Image1(number2).Picture = Image1(number1).Picture
            Form2.Bar1.Value = i
            Form2.percent.Caption = Int(i / 18 * 100) & "%"
            DoEvents
        Next i
    End If
    Unload Form2
End Sub


Private Sub Timer1_Timer()
    time = time + 1
    If time = 2 Then
        time = 0
        Timer1.Enabled = False
    End If
End Sub
