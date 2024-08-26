VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Lista2 (Zurba_Ferrante).frx":0000
   ScaleHeight     =   14115
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   4
      Left            =   7440
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   3
      Left            =   7440
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   2
      Left            =   7440
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Lista2 (Zurba_Ferrante).frx":0342
      Left            =   9240
      List            =   "Lista2 (Zurba_Ferrante).frx":0344
      TabIndex        =   2
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   0
      Left            =   7440
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   7440
      TabIndex        =   0
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Matematica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   4800
      TabIndex        =   11
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Lengua"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "C. Ciudadana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4800
      TabIndex        =   9
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4800
      TabIndex        =   8
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Geografia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer
Private Sub Command1_Click()

If List1.ListCount < 5 Then

    For A = 0 To 4
       
        If Val(Text1(A).Text) < 0 Or Val(Text1(A).Text) > 10 Then
            
            List1.AddItem "Escribiste mal la nota"
            
        Else
            
            List1.AddItem Label1(A).Caption & ": " & Text1(A).Text
            
        End If
        
    Next A

End If

End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index >= 0 And Index <= 4 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
    Else
        
    End If
    
End Sub
