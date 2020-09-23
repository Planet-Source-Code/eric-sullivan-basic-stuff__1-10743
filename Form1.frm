VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool stuff"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   5640
      X2              =   5640
      Y1              =   1800
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   4080
      X2              =   5640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4080
      X2              =   4080
      Y1              =   1320
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   4080
      X2              =   5640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   " Cool Stuff: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   40
      Width           =   1125
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1905
      Left            =   135
      Shape           =   4  'Rounded Rectangle
      Top             =   135
      Width           =   5745
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1935
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "  The author does not seem to put effort in to his programs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   855
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   5
      Left            =   240
      Picture         =   "Form1.frx":0000
      Top             =   840
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   240
      Picture         =   "Form1.frx":0397
      Top             =   840
      Width           =   285
   End
   Begin VB.Label Label2 
      Caption         =   "  This Program is neat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   615
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   4
      Left            =   240
      Picture         =   "Form1.frx":0701
      Top             =   600
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":0A98
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "  This Program is useful"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   375
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   3
      Left            =   240
      Picture         =   "Form1.frx":0E02
      Top             =   360
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":1199
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Image1(3).Visible = False
    Image1(4).Visible = False
    Image1(5).Visible = False
    
    Call SetButtonValue(False, False, False, False)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetButtonValue(False, False, False, False)
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Image1(Index).Index
        Case 0: Call ConvertImages(Image1(3), Image1(0))
        Case 1: Call ConvertImages(Image1(4), Image1(1))
        Case 2: Call ConvertImages(Image1(5), Image1(2))
        Case 3: Call ConvertImages(Image1(0), Image1(3))
        Case 4: Call ConvertImages(Image1(1), Image1(4))
        Case 5: Call ConvertImages(Image1(2), Image1(5))
    End Select
End Sub

Private Sub Label5_Click()
    End
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetButton(&H808080, &HFFFFFF)
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetButtonValue(True, True, True, True)
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetButton(&HFFFFFF, &H808080)
End Sub
