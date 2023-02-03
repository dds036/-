VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test"
   ClientHeight    =   6465
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   11745
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   9
      Left            =   2160
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Comm2 
      Caption         =   "Настройки"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   8
      Left            =   2160
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   7
      Left            =   2160
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   6
      Left            =   2160
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   5
      Left            =   2160
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   4
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Начать!"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      MaskColor       =   &H8000000D&
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   600
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   600
      TabIndex        =   25
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   600
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   600
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Label5"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Внимание!"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   27.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   $"Form1.frx":0342
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   5520
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "  Fortitude Comp."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   10680
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Научись считать"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   48.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1200
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   7815
   End
   Begin VB.Menu File 
      Caption         =   "Файл"
      Begin VB.Menu Upd 
         Caption         =   "Обновить"
      End
      Begin VB.Menu Exit 
         Caption         =   "Выйти"
      End
   End
   Begin VB.Menu spravka 
      Caption         =   "Справка"
      Begin VB.Menu about 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()

End Sub

Private Sub Form_Load()

End Sub
