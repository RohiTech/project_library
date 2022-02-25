VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Generar_Backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESPALDO"
   ClientHeight    =   2130
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   7185
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Txtdestino 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BACKUP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   240
      Picture         =   "Generar_Backup.frx":0000
      ScaleHeight     =   1020
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   840
      Width           =   795
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "Generar_Backup.frx":2AC2
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "Generar_Backup.frx":2B2E
      Top             =   2880
   End
End
Attribute VB_Name = "Generar_Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NombreBaseDatos As String
Dim Driver As String

Private Sub Command1_Click()
On Error Resume Next
    Randomize Time
    Dialog.FileName = ""

    Dialog.Filter = "Archivos File (*.bak) |*.bak|"
    Dialog.ShowSave
    If Dialog.FileName <> "" Then
        Txtdestino = Dialog.FileName
        Driver = "disk"
        NombreBaseDatos = (Rnd * 100)
        cn.Execute " sp_addumpdevice '" & Driver & "','" & NombreBaseDatos & "','" & Txtdestino & "'"
    Else
        Txtdestino = ""
      
    End If

End Sub

Private Sub Command2_Click()
On Error GoTo e
cn.Execute "respaldo '" & strDB & "','" & NombreBaseDatos & "'"
MsgBox "Se ha realizado el respaldo satisfactoriamente", vbInformation, "Confirmacion"
Unload Me
e:
If Err.Number <> 0 Then
  MsgBox "Error al realizar el respaldo", vbInformation, "AVISO"
End If
End Sub
Private Sub Form_Load()
 Skin1.ApplySkin hWnd

    With Generar_Backup
        .Height = 2580
        .Width = 7290
    End With
 
End Sub

Private Sub Picture1_Click()
Picture1.Appearance = 0
End Sub

