VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Restaurar_BD 
   Caption         =   "Restaurar BD"
   ClientHeight    =   1860
   ClientLeft      =   150
   ClientTop       =   405
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   7290
   Begin VB.TextBox Txtdestino 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   840
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Restaurar_BD.frx":0000
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Restaurar_BD.frx":0078
      Top             =   0
   End
End
Attribute VB_Name = "Restaurar_BD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NombreBaseDatos As String
Private Sub Command1_Click()
On Error Resume Next
    Dialog.FileName = ""

    Dialog.Filter = "Archivos File (*.bak) |*.bak|"
    Dialog.ShowOpen
    
    If Dialog.FileName <> "" Then
        Txtdestino = Dialog.FileName
        Text1 = Dialog.FileTitle
      
    End If

End Sub

Private Sub Command2_Click()
On Error GoTo e
cn.Execute "restaurar '" & strDB & "','" & strDB & "Resp1" & "'"
MsgBox "Se ha restaurado la base de datos satisfactoriamente", vbInformation, "Confirmacion"
Unload Me
e:
If Err.Number <> 0 Then
  MsgBox "Error al restaurar Base de datos", vbInformation, "AVISO"
End If
End Sub

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
 
 With Restaurar_BD
    .Height = 2325
    .Width = 7410
 End With
 
End Sub

