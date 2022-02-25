VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Agradecimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agradecimiento"
   ClientHeight    =   4875
   ClientLeft      =   225
   ClientTop       =   465
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   11385
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   3360
      OleObjectBlob   =   "Agradecimiento.frx":0000
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   240
      Picture         =   "Agradecimiento.frx":007E
      ScaleHeight     =   3555
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   3360
      OleObjectBlob   =   "Agradecimiento.frx":3218
      TabIndex        =   2
      Top             =   960
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   735
      Left            =   3360
      OleObjectBlob   =   "Agradecimiento.frx":3346
      TabIndex        =   3
      Top             =   1680
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   3360
      OleObjectBlob   =   "Agradecimiento.frx":3490
      TabIndex        =   4
      Top             =   2640
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   495
      Left            =   3360
      OleObjectBlob   =   "Agradecimiento.frx":359E
      TabIndex        =   5
      Top             =   3360
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   495
      Left            =   3360
      OleObjectBlob   =   "Agradecimiento.frx":368E
      TabIndex        =   6
      Top             =   4080
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Agradecimiento.frx":377A
      Top             =   120
   End
End
Attribute VB_Name = "Agradecimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
End Sub
