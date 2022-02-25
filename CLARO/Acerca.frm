VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Acerca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca del sistema de control para tarjetas telefónicas ""CLARO"""
   ClientHeight    =   4680
   ClientLeft      =   210
   ClientTop       =   465
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   13800
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   -360
      Picture         =   "Acerca.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Acerca.frx":3EA6
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "Acerca.frx":40BC7
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "Acerca.frx":40C85
      TabIndex        =   2
      Top             =   480
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "Acerca.frx":40D31
      TabIndex        =   3
      Top             =   1200
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "Acerca.frx":40DBF
      TabIndex        =   4
      Top             =   1560
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   1815
      Left            =   5760
      OleObjectBlob   =   "Acerca.frx":40E2B
      TabIndex        =   5
      Top             =   2280
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "Acerca.frx":41181
      TabIndex        =   6
      Top             =   4200
      Width           =   7695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "Acerca.frx":4120B
      TabIndex        =   7
      Top             =   840
      Width           =   7695
   End
End
Attribute VB_Name = "Acerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    
    With Acerca
        .Height = 5145
        .Width = 13920
    End With
    
End Sub
