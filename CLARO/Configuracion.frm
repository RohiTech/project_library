VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Configuracion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración"
   ClientHeight    =   4275
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Tarjetas"
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   3015
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Configuracion.frx":0000
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comisiones"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3015
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Configuracion.frx":007A
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "Configuracion.frx":00EA
         TabIndex        =   6
         Top             =   480
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Configuracion.frx":014A
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "Configuracion.frx":01BC
         TabIndex        =   8
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Configuracion.frx":021C
      Top             =   4680
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   3600
      TabIndex        =   9
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   975
      Left            =   4560
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
   End
End
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
End Sub
