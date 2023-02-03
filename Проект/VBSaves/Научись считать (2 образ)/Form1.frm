VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Научись считать"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Начать"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   68
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Обновить"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   67
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox status_dif 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox status_tip 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text_10_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_10_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text_10_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   55
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Rz_10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   5760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Rz_9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text_9_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   51
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text_9_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_9 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_9_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Rz_8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   4560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text_8_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   45
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text_8_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_8_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text_7_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_7_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text_7_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   37
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Rz_7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Rz_6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text_6_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   33
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text_6_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_6_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text_5_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_5_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text_5_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Rz_5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Rz_4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text_4_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text_4_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_4_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text_3_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_3_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text_3_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Rz_3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Rz_2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text_2_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text_2_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_2_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Rz_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text_1_3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text_1_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Zn_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text_1_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label welcome 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "И добро пожаловать в"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   4800
      TabIndex        =   75
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label halo 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Здраствуйте,"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   4800
      TabIndex        =   74
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label halo_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   6960
      TabIndex        =   73
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label status_score 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   48
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1455
      Left            =   11040
      TabIndex        =   72
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label score 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Оценка:"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   27.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   71
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label status_err 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   70
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label errs 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Количество ошибок:"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   69
      Top             =   1560
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label chosen_dif 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Выбранная сложность:"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   65
      Top             =   960
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label chosen_tip 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Выбранный тип:"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   64
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Внимание!"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5880
      TabIndex        =   62
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3960
      TabIndex        =   61
      Top             =   5520
      Width           =   6495
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Научись считать"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   68.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   2520
      TabIndex        =   60
      Top             =   1080
      Width           =   9375
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   59
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   53
      Top             =   5160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   47
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   41
      Top             =   3960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu file 
      Caption         =   "Файл"
      Begin VB.Menu upd 
         Caption         =   "Обновить"
      End
      Begin VB.Menu exit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu option 
      Caption         =   "Настройки"
      Begin VB.Menu mode 
         Caption         =   "Сложность"
         Begin VB.Menu easy 
            Caption         =   "Легко"
            Checked         =   -1  'True
         End
         Begin VB.Menu norm 
            Caption         =   "Нормально"
         End
         Begin VB.Menu hard 
            Caption         =   "Сложно"
         End
      End
      Begin VB.Menu tip 
         Caption         =   "Тип"
         Begin VB.Menu plus 
            Caption         =   "Сложение"
            Checked         =   -1  'True
         End
         Begin VB.Menu minus 
            Caption         =   "Вычитание"
         End
         Begin VB.Menu umnogenie 
            Caption         =   "Умножение"
         End
         Begin VB.Menu delit 
            Caption         =   "Деление"
         End
         Begin VB.Menu All 
            Caption         =   "Все"
         End
      End
   End
   Begin VB.Menu help 
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
Dim MyEr As Integer
Dim now_score As Integer

Private Sub All_Click()
Me![plus].Checked = False
Me![minus].Checked = False
Me![All].Checked = True
End Sub

Private Sub Command1_Click()
Dim dif As Integer
Dim i As Integer
 For i = 1 To 10
  If Len(Me.Controls("Rz_" & i)) > 0 Then
     Me.Controls("Rz_" & i).Text = ""
  End If
 Next i

If Me![norm].Checked = True Then
    dif = 100
End If

If Me![easy].Checked = True Then
    dif = 20
End If

If Me![hard].Checked = True Then
    dif = 1000
End If

Randomize
If Me![minus].Checked = True Then
    Rz_1.Visible = False
    Text_1_3 = ""
    Text_1_1 = Int((dif * Rnd) + 1)
    Zn_1 = "-"
    Text_1_2 = Int((Text_1_1 * Rnd) + 1)
    
    Rz_2.Visible = False
    Text_2_3 = ""
    Text_2_1 = Int((dif * Rnd) + 1)
    Zn_2 = "-"
    Text_2_2 = Int((Text_2_1 * Rnd) + 1)

    Rz_3.Visible = False
    Text_3_3 = ""
    Text_3_1 = Int((dif * Rnd) + 1)
    Zn_3 = "-"
    Text_3_2 = Int((Text_3_1 * Rnd) + 1)

    Rz_4.Visible = False
    Text_4_3 = ""
    Text_4_1 = Int((dif * Rnd) + 1)
    Zn_4 = "-"
    Text_4_2 = Int((Text_4_1 * Rnd) + 1)
    
    Rz_5.Visible = False
    Text_5_3 = ""
    Text_5_1 = Int((dif * Rnd) + 1)
    Zn_5 = "-"
    Text_5_2 = Int((Text_5_1 * Rnd) + 1)
    
    Rz_6.Visible = False
    Text_6_3 = ""
    Text_6_1 = Int((dif * Rnd) + 1)
    Zn_6 = "-"
    Text_6_2 = Int((Text_6_1 * Rnd) + 1)
    
    Rz_7.Visible = False
    Text_7_3 = ""
    Text_7_1 = Int((dif * Rnd) + 1)
    Zn_7 = "-"
    Text_7_2 = Int((Text_7_1 * Rnd) + 1)
    
    Rz_8.Visible = False
    Text_8_3 = ""
    Text_8_1 = Int((dif * Rnd) + 1)
    Zn_8 = "-"
    Text_8_2 = Int((Text_8_1 * Rnd) + 1)
    
    Rz_9.Visible = False
    Text_9_3 = ""
    Text_9_1 = Int((dif * Rnd) + 1)
    Zn_9 = "-"
    Text_9_2 = Int((Text_9_1 * Rnd) + 1)
    
    Rz_10.Visible = False
    Text_10_3 = ""
    Text_10_1 = Int((dif * Rnd) + 1)
    Zn_10 = "-"
    Text_10_2 = Int((Text_10_1 * Rnd) + 1)
End If

If Me![plus].Checked = True Then
    Rz_1.Visible = False
    Text_1_3 = ""
    Text_1_1 = Int((dif * Rnd) + 1)
    Zn_1 = "+"
    Text_1_2 = Int(((dif - Text_1_1) * Rnd) + 1)
    
    Rz_2.Visible = False
    Text_2_3 = ""
    Text_2_1 = Int((dif * Rnd) + 1)
    Zn_2 = "+"
    Text_2_2 = Int(((dif - Text_2_1) * Rnd) + 1)

    Rz_3.Visible = False
    Text_3_3 = ""
    Text_3_1 = Int((dif * Rnd) + 1)
    Zn_3 = "+"
    Text_3_2 = Int(((dif - Text_3_1) * Rnd) + 1)

    Rz_4.Visible = False
    Text_4_3 = ""
    Text_4_1 = Int((dif * Rnd) + 1)
    Zn_4 = "+"
    Text_4_2 = Int(((dif - Text_4_1) * Rnd) + 1)
    
    Rz_5.Visible = False
    Text_5_3 = ""
    Text_5_1 = Int((dif * Rnd) + 1)
    Zn_5 = "+"
    Text_5_2 = Int(((dif - Text_5_1) * Rnd) + 1)
    
    Rz_6.Visible = False
    Text_6_3 = ""
    Text_6_1 = Int((dif * Rnd) + 1)
    Zn_6 = "+"
    Text_6_2 = Int(((dif - Text_6_1) * Rnd) + 1)
    
    Rz_7.Visible = False
    Text_7_3 = ""
    Text_7_1 = Int((dif * Rnd) + 1)
    Zn_7 = "+"
    Text_7_2 = Int(((dif - Text_7_1) * Rnd) + 1)
    
    Rz_8.Visible = False
    Text_8_3 = ""
    Text_8_1 = Int((dif * Rnd) + 1)
    Zn_8 = "+"
    Text_8_2 = Int(((dif - Text_8_1) * Rnd) + 1)
    
    Rz_9.Visible = False
    Text_9_3 = ""
    Text_9_1 = Int((dif * Rnd) + 1)
    Zn_9 = "+"
    Text_9_2 = Int(((dif - Text_9_1) * Rnd) + 1)
    
    Rz_10.Visible = False
    Text_10_3 = ""
    Text_10_1 = Int((dif * Rnd) + 1)
    Zn_10 = "+"
    Text_10_2 = Int(((dif - Text_10_1) * Rnd) + 1)
    
End If

If Me![All].Checked = True Then
    Rz_1.Visible = False
    Text_1_3 = ""
    Text_1_1 = Int((dif * Rnd) + 1)
    Zn_1 = "+" Or "-"
    Text_1_2 = Int(((dif - Text_1_1) * Rnd) + 1)
    
    Rz_2.Visible = False
    Text_2_3 = ""
    Text_2_1 = Int((dif * Rnd) + 1)
    Zn_2 = "+" Or "-"
    Text_2_2 = Int(((dif - Text_2_1) * Rnd) + 1)

    Rz_3.Visible = False
    Text_3_3 = ""
    Text_3_1 = Int((dif * Rnd) + 1)
    Zn_3 = "+" Or "-"
    Text_3_2 = Int(((dif - Text_3_1) * Rnd) + 1)

    Rz_4.Visible = False
    Text_4_3 = ""
    Text_4_1 = Int((dif * Rnd) + 1)
    Zn_4 = "+" Or "-"
    Text_4_2 = Int(((dif - Text_4_1) * Rnd) + 1)
    
    Rz_5.Visible = False
    Text_5_3 = ""
    Text_5_1 = Int((dif * Rnd) + 1)
    Zn_5 = "+" Or "-"
    Text_5_2 = Int(((dif - Text_5_1) * Rnd) + 1)
    
    Rz_6.Visible = False
    Text_6_3 = ""
    Text_6_1 = Int((dif * Rnd) + 1)
    Zn_6 = "+" Or "-"
    Text_6_2 = Int(((dif - Text_6_1) * Rnd) + 1)
    
    Rz_7.Visible = False
    Text_7_3 = ""
    Text_7_1 = Int((dif * Rnd) + 1)
    Zn_7 = "+" Or "-"
    Text_7_2 = Int(((dif - Text_7_1) * Rnd) + 1)
    
    Rz_8.Visible = False
    Text_8_3 = ""
    Text_8_1 = Int((dif * Rnd) + 1)
    Zn_8 = "+" Or "-"
    Text_8_2 = Int(((dif - Text_8_1) * Rnd) + 1)
    
    Rz_9.Visible = False
    Text_9_3 = ""
    Text_9_1 = Int((dif * Rnd) + 1)
    Zn_9 = "+" Or "-"
    Text_9_2 = Int(((dif - Text_9_1) * Rnd) + 1)
    
    Rz_10.Visible = False
    Text_10_3 = ""
    Text_10_1 = Int((dif * Rnd) + 1)
    Zn_10 = "+" Or "-"
    Text_10_2 = Int(((dif - Text_10_1) * Rnd) + 1)
    
End If


If Me![umnogenie].Checked = True Then
    Rz_1.Visible = False
    Text_1_3 = ""
    Text_1_1 = Int((dif * Rnd) + 1)
    Zn_1 = "*"
    Text_1_2 = Int(((dif - Text_1_1) * Rnd) + 1)
    
    Rz_2.Visible = False
    Text_2_3 = ""
    Text_2_1 = Int((dif * Rnd) + 1)
    Zn_2 = "*"
    Text_2_2 = Int(((dif - Text_2_1) * Rnd) + 1)

    Rz_3.Visible = False
    Text_3_3 = ""
    Text_3_1 = Int((dif * Rnd) + 1)
    Zn_3 = "*"
    Text_3_2 = Int(((dif - Text_3_1) * Rnd) + 1)

    Rz_4.Visible = False
    Text_4_3 = ""
    Text_4_1 = Int((dif * Rnd) + 1)
    Zn_4 = "*"
    Text_4_2 = Int(((dif - Text_4_1) * Rnd) + 1)
    
    Rz_5.Visible = False
    Text_5_3 = ""
    Text_5_1 = Int((dif * Rnd) + 1)
    Zn_5 = "*"
    Text_5_2 = Int(((dif - Text_5_1) * Rnd) + 1)
    
    Rz_6.Visible = False
    Text_6_3 = ""
    Text_6_1 = Int((dif * Rnd) + 1)
    Zn_6 = "*"
    Text_6_2 = Int(((dif - Text_6_1) * Rnd) + 1)
    
    Rz_7.Visible = False
    Text_7_3 = ""
    Text_7_1 = Int((dif * Rnd) + 1)
    Zn_7 = "*"
    Text_7_2 = Int(((dif - Text_7_1) * Rnd) + 1)
    
    Rz_8.Visible = False
    Text_8_3 = ""
    Text_8_1 = Int((dif * Rnd) + 1)
    Zn_8 = "*"
    Text_8_2 = Int(((dif - Text_8_1) * Rnd) + 1)
    
    Rz_9.Visible = False
    Text_9_3 = ""
    Text_9_1 = Int((dif * Rnd) + 1)
    Zn_9 = "*"
    Text_9_2 = Int(((dif - Text_9_1) * Rnd) + 1)
    
    Rz_10.Visible = False
    Text_10_3 = ""
    Text_10_1 = Int((dif * Rnd) + 1)
    Zn_10 = "*"
    Text_10_2 = Int(((dif - Text_10_1) * Rnd) + 1)
End If


'If Me![delit].Checked = True Then
'Rz_1.Visible = False
 '   Text_1_3 = ""
 '   Text_1_1 = Int((dif * Rnd) + 1)
  '  Zn_1 = "/"
 '   Text_1_2 = Int(((dif - Text_1_1) * Rnd) + 1)
'If (Text_1_1 > Text_1_2) and (Text_1_1 / 2 = mod


'End If



MyEr = 0
status_err.Caption = MyEr
status_score.Visible = False
status_score.Refresh

Call ocenka


End Sub

Private Sub Command2_Click()

Label11.Visible = False
Command2.Visible = False
Label12.Visible = False
Label13.Visible = False
halo.Visible = False
halo_name.Visible = False
welcome.Visible = False

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
errs.Visible = True
status_err.Visible = True
Zn_1.Visible = True
Zn_2.Visible = True
Zn_3.Visible = True
Zn_4.Visible = True
Zn_5.Visible = True
Zn_6.Visible = True
Zn_7.Visible = True
Zn_8.Visible = True
Zn_9.Visible = True
Zn_10.Visible = True

Rz_1.Visible = True
Rz_2.Visible = True
Rz_3.Visible = True
Rz_4.Visible = True
Rz_5.Visible = True
Rz_6.Visible = True
Rz_7.Visible = True
Rz_8.Visible = True
Rz_9.Visible = True
Rz_10.Visible = True


score.Visible = True

Command1.Visible = True

Text_1_1.Visible = True
Text_2_1.Visible = True
Text_3_1.Visible = True
Text_4_1.Visible = True
Text_5_1.Visible = True
Text_6_1.Visible = True
Text_7_1.Visible = True
Text_8_1.Visible = True
Text_9_1.Visible = True
Text_10_1.Visible = True
Text_1_2.Visible = True
Text_2_2.Visible = True
Text_3_2.Visible = True
Text_4_2.Visible = True
Text_5_2.Visible = True
Text_6_2.Visible = True
Text_7_2.Visible = True
Text_8_2.Visible = True
Text_9_2.Visible = True
Text_10_2.Visible = True
Text_1_3.Visible = True
Text_2_3.Visible = True
Text_3_3.Visible = True
Text_4_3.Visible = True
Text_5_3.Visible = True
Text_6_3.Visible = True
Text_7_3.Visible = True
Text_8_3.Visible = True
Text_9_3.Visible = True
Text_10_3.Visible = True

status_tip.Visible = True
status_dif.Visible = True

chosen_tip.Visible = True
chosen_dif.Visible = True

Call Command1_Click
Call easy_Click
Call plus_Click

End Sub

Private Sub Command3_Click()

End Sub

Private Sub easy_Click()

Me![norm].Checked = False
Me![easy].Checked = True
Me![hard].Checked = False

Me![status_dif] = "Легко"

Call Command1_Click

End Sub

Private Sub exit_click()

 End
 
End Sub

Private Sub Form_Load()
MyEr = 0
halo_name.Caption = CreateObject("wscript.network").UserName
End Sub

Private Sub halo_name_Click()

End Sub

Private Sub hard_Click()

Me![norm].Checked = False
Me![easy].Checked = False
Me![hard].Checked = True

Me![status_dif] = "Сложно"
Call Command1_Click
End Sub

Private Sub norm_Click()

Me![norm].Checked = True
Me![easy].Checked = False
Me![hard].Checked = False

Me![status_dif] = "Нормально"
Call Command1_Click
End Sub

Private Sub plus_Click()

Me![plus].Checked = True
Me![minus].Checked = False
Me![All].Checked = False
Me![umnogenie].Checked = False
Me![delit].Checked = False

Me![status_tip] = "Сложение"
Call Command1_Click
End Sub
Private Sub minus_Click()

Me![minus].Checked = True
Me![plus].Checked = False
Me![All].Checked = False
Me![umnogenie].Checked = False
Me![delit].Checked = False

Me![status_tip] = "Вычитание"
Call Command1_Click
End Sub

Private Sub status_score_Click()
If MyEr >= 1 Then
    now_score = 4
    status_score.Caption = now_score
    status_score.ForeColor = &HFFFF&
    status_score.Visible = True
End If

If MyEr >= 3 Then
    now_score = 3
    status_score.Caption = now_score
    status_score.ForeColor = &H80FF&
    status_score.Visible = True
End If

If MyEr >= 5 Then
    now_score = 2
    status_score.Caption = now_score
    status_score.ForeColor = &HFF&
    status_score.Visible = True
End If
End Sub

Private Sub umnogenie_Click()

Me![minus].Checked = False
Me![plus].Checked = False
Me![All].Checked = False
Me![umnogenie].Checked = True
Me![delit].Checked = False

Me![status_tip] = "Умножение"
Call Command1_Click
End Sub
Private Sub delit_Click()

Me![minus].Checked = False
Me![plus].Checked = False
Me![All].Checked = False
Me![umnogenie].Checked = False
Me![delit].Checked = True

Me![status_tip] = "Деление"
Call Command1_Click
End Sub


Private Sub Text_1_3_KeyPress(KeyAscii As Integer)

On Error GoTo err_my

If KeyAscii = 13 Then
        
        
        If Zn_1 = "-" Then
            If CInt(Text_1_3) = (CInt(Text_1_1) - CInt(Text_1_2)) Then
                Rz_1 = "Правильно!!"
            Else
                Rz_1 = "Неправильно!!"
                MyEr = MyEr + 1
                status_err.Caption = MyEr
            End If
        End If

        If Zn_1 = "+" Then
            If Text_1_3 = (CInt(Text_1_1) + CInt(Text_1_2)) Then
                Rz_1 = "Правильно!!"
            Else
                Rz_1 = "Неправильно!!"
                MyEr = MyEr + 1
                status_err.Caption = MyEr
            End If
        End If

        If Zn_1 = "*" Then
            If Text_1_3 = (CInt(Text_1_1) * CInt(Text_1_2)) Then
                Rz_1 = "Правильно!!"
            Else
                Rz_1 = "Неправильно!!"
                MyEr = MyEr + 1
                status_err.Caption = MyEr
            End If
        End If
        
        If Zn_1 = "/" Then
            If Text_1_3 = (CInt(Text_1_1) / CInt(Text_1_2)) Then
                Rz_1 = "Правильно!!"
            Else
                Rz_1 = "Неправильно!!"
            End If
        End If

        Rz_1.Refresh
        Rz_1.Visible = True
        Call ocenka
End If
Exit Sub
err_my:
  If Err.Number = 13 Then
     MsgBox ("Введено недопустимое значение")
     Text_1_3 = ""
  Else
     MsgBox (CStr(Err.Number) & " " & Err.Description)
  End If
  
End Sub

Private Sub Text_10_3_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

 Rz_10.Visible = True
 
If Zn_10 = "-" Then
  If CInt(Text_10_3) = (CInt(Text_10_1) - CInt(Text_10_2)) Then
    Rz_10 = "Правильно!!"
  Else
    Rz_10 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If

If Zn_10 = "+" Then
  If Text_10_3 = (CInt(Text_10_1) + CInt(Text_10_2)) Then
    Rz_10 = "Правильно!!"
  Else
    Rz_10 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  
  End If
End If

If Zn_10 = "*" Then
  If Text_10_3 = (CInt(Text_10_1) * CInt(Text_10_2)) Then
    Rz_10 = "Правильно!!"
  Else
    Rz_10 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr




  End If
End If

Rz_10.Refresh
Call ocenka
End If

End Sub

Private Sub Text_2_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_2.Visible = True
If Zn_2 = "-" Then
  If CInt(Text_2_3) = (CInt(Text_2_1) - CInt(Text_2_2)) Then
    Rz_2 = "Правильно!!"
  Else
    Rz_2 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_2 = "+" Then
  If Text_2_3 = (CInt(Text_2_1) + CInt(Text_2_2)) Then
    Rz_2 = "Правильно!!"
  Else
    Rz_2 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_2 = "*" Then
  If Text_2_3 = (CInt(Text_2_1) * CInt(Text_2_2)) Then
    Rz_2 = "Правильно!!"
  Else
    Rz_2 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If

End If
Rz_2.Refresh
Call ocenka
End If

End Sub

Private Sub Text_3_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_3.Visible = True
If Zn_3 = "-" Then
  If CInt(Text_3_3) = (CInt(Text_3_1) - CInt(Text_3_2)) Then
    Rz_3 = "Правильно!!"
  Else
    Rz_3 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_3 = "+" Then
  If Text_3_3 = (CInt(Text_3_1) + CInt(Text_3_2)) Then
    Rz_3 = "Правильно!!"
  Else
    Rz_3 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_3 = "*" Then
  If Text_3_3 = (CInt(Text_3_1) * CInt(Text_3_2)) Then
    Rz_3 = "Правильно!!"
  Else
    Rz_3 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
  
End If
Rz_3.Refresh
Call ocenka
End If
End Sub

Private Sub Text_4_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_4.Visible = True
If Zn_4 = "-" Then
  If CInt(Text_4_3) = (CInt(Text_4_1) - CInt(Text_4_2)) Then
    Rz_4 = "Правильно!!"
  Else
    Rz_4 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_4 = "+" Then
  If Text_4_3 = (CInt(Text_4_1) + CInt(Text_4_2)) Then
    Rz_4 = "Правильно!!"
  Else
    Rz_4 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_4 = "*" Then
  If Text_4_3 = (CInt(Text_4_1) * CInt(Text_4_2)) Then
    Rz_4 = "Правильно!!"
  Else
    Rz_4 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If

End If
Rz_4.Refresh
Call ocenka
End If

End Sub

Private Sub Text_5_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_5.Visible = True
If Zn_5 = "-" Then
  If CInt(Text_5_3) = (CInt(Text_5_1) - CInt(Text_5_2)) Then
    Rz_5 = "Правильно!!"
  Else
    Rz_5 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_5 = "+" Then
  If Text_5_3 = (CInt(Text_5_1) + CInt(Text_5_2)) Then
    Rz_5 = "Правильно!!"
  Else
    Rz_5 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_5 = "*" Then
  If Text_5_3 = (CInt(Text_5_1) * CInt(Text_5_2)) Then
    Rz_5 = "Правильно!!"
  Else
    Rz_5 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
Rz_5.Refresh
Call ocenka
End If
End Sub

Private Sub Text_6_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_6.Visible = True
If Zn_6 = "-" Then
  If CInt(Text_6_3) = (CInt(Text_6_1) - CInt(Text_6_2)) Then
    Rz_6 = "Правильно!!"
  Else
    Rz_6 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_6 = "+" Then
  If Text_6_3 = (CInt(Text_6_1) + CInt(Text_6_2)) Then
    Rz_6 = "Правильно!!"
  Else
    Rz_6 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_6 = "*" Then
  If Text_6_3 = (CInt(Text_6_1) * CInt(Text_6_2)) Then
    Rz_6 = "Правильно!!"
  Else
    Rz_6 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If

End If
Rz_6.Refresh
Call ocenka
End If
End Sub

Private Sub Text_7_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_7.Visible = True
If Zn_7 = "-" Then
  If CInt(Text_7_3) = (CInt(Text_7_1) - CInt(Text_7_2)) Then
    Rz_7 = "Правильно!!"
  Else
    Rz_7 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_7 = "+" Then
  If Text_7_3 = (CInt(Text_7_1) + CInt(Text_7_2)) Then
    Rz_7 = "Правильно!!"
  Else
    Rz_7 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_7 = "*" Then
  If Text_7_3 = (CInt(Text_7_1) * CInt(Text_7_2)) Then
    Rz_7 = "Правильно!!"
  Else
    Rz_7 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If

End If
Rz_7.Refresh
Call ocenka
End If
End Sub

Private Sub Text_8_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_8.Visible = True
If Zn_8 = "-" Then
  If CInt(Text_8_3) = (CInt(Text_8_1) - CInt(Text_8_2)) Then
    Rz_8 = "Правильно!!"
  Else
    Rz_8 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_8 = "+" Then
  If Text_8_3 = (CInt(Text_8_1) + CInt(Text_8_2)) Then
    Rz_8 = "Правильно!!"
  Else
    Rz_8 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_8 = "*" Then
  If Text_8_3 = (CInt(Text_8_1) * CInt(Text_8_2)) Then
    Rz_8 = "Правильно!!"
  Else
    Rz_8 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If

End If
Rz_8.Refresh
Call ocenka
End If
End Sub

Private Sub Text_9_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Rz_9.Visible = True
If Zn_9 = "-" Then
  If CInt(Text_9_3) = (CInt(Text_9_1) - CInt(Text_9_2)) Then
    Rz_9 = "Правильно!!"
  Else
    Rz_9 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_9 = "+" Then
  If Text_9_3 = (CInt(Text_9_1) + CInt(Text_9_2)) Then
    Rz_9 = "Правильно!!"
  Else
    Rz_9 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If
End If
If Zn_9 = "*" Then
  If Text_9_3 = (CInt(Text_9_1) * CInt(Text_9_2)) Then
    Rz_9 = "Правильно!!"
  Else
    Rz_9 = "Неправильно!!"
    MyEr = MyEr + 1
    status_err.Caption = MyEr
  End If

End If
Rz_9.Refresh
Call ocenka
End If
End Sub

Private Sub upd_Click()
Call Command1_Click
End Sub
Sub ocenka()
Dim i As Integer
Dim f As Boolean
f = True
 For i = 1 To 10
  If Len(Me.Controls("Rz_" & i)) = 0 Then
  'If Me.Controls("Rz_" & i).Visible = False Then
     f = False
     Exit For
  End If
 Next i
 If f = True Then
   Call status_score_Click
 End If
End Sub
