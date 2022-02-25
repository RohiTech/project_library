VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Contacto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacto"
   ClientHeight    =   5355
   ClientLeft      =   105
   ClientTop       =   420
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6465
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   240
      Picture         =   "Contacto.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Contacto.frx":46B7
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "Contacto.frx":413D8
      TabIndex        =   1
      Top             =   4800
      Width           =   4455
   End
End
Attribute VB_Name = "Contacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
End Sub
