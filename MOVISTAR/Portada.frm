VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Portada 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5370
   ClientLeft      =   375
   ClientTop       =   1470
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Portada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   -6840
      Picture         =   "Portada.frx":000C
      ScaleHeight     =   1575
      ScaleWidth      =   13935
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   30
         Left            =   4080
         OleObjectBlob   =   "Portada.frx":1C1B9
         TabIndex        =   1
         Top             =   1560
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   1200
      Top             =   6720
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "Portada.frx":1C233
      Top             =   6720
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   5400
      OleObjectBlob   =   "Portada.frx":66738
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   4200
      OleObjectBlob   =   "Portada.frx":667B0
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   1455
      Left            =   120
      OleObjectBlob   =   "Portada.frx":66840
      TabIndex        =   4
      Top             =   2400
      Width           =   6615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   4200
      OleObjectBlob   =   "Portada.frx":6691E
      TabIndex        =   5
      Top             =   4680
      Width           =   4215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   3360
      OleObjectBlob   =   "Portada.frx":669A8
      TabIndex        =   6
      Top             =   4320
      Width           =   5535
   End
End
Attribute VB_Name = "Portada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
End Sub

Private Sub Timer1_Timer()
    frmconectar.Show
    Unload Me
End Sub
