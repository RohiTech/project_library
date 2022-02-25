VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Seguridad_Admon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguridad del Administrador"
   ClientHeight    =   1440
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   270
      Width           =   3735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Seguridad_Admon.frx":0000
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7080
      OleObjectBlob   =   "Seguridad_Admon.frx":006C
      Top             =   1320
   End
End
Attribute VB_Name = "Seguridad_Admon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
End Sub

Private Sub OKButton_Click()
    Principal.Text3.Text = UCase(Text1.Text)
End Sub
