VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Calendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   4755
   ClientLeft      =   225
   ClientTop       =   465
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6450
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   3120
      Picture         =   "Calendario.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   1800
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Calendario.frx":3483
      Top             =   5280
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   20054017
      CurrentDate     =   39039
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20054017
      CurrentDate     =   39039
   End
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'theese are the clock stuffs
Private Sub Form_Load()
Skin1.ApplySkin hWnd

    Timer1.Interval = 1
    
DTPicker1 = Format(Date, "Short Date")
MonthView1 = Format(Date, "Short Date")

With Calendario
    .Height = 5190
    .Width = 6540
End With

End Sub


Private Sub Timer1_Timer()
    Text1.Text = Time
End Sub
