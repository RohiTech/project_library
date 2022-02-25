VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Autor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autor"
   ClientHeight    =   8235
   ClientLeft      =   105
   ClientTop       =   420
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10170
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   240
      Picture         =   "Autor.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   240
      Width           =   9615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Autor.frx":905F
      TabIndex        =   1
      Top             =   7680
      Width           =   9615
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Autor.frx":90F3
      Top             =   0
   End
End
Attribute VB_Name = "Autor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
End Sub
