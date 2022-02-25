VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Cambiar_Contraseña 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar contraseña"
   ClientHeight    =   2355
   ClientLeft      =   210
   ClientTop       =   420
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4755
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3120
      OleObjectBlob   =   "Cambiar_Contraseña.frx":0000
      Top             =   2880
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3720
      Picture         =   "Cambiar_Contraseña.frx":3CD21
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "Cambiar_Contraseña.frx":3D259
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Cambiar_Contraseña.frx":3D2DF
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Cambiar_Contraseña.frx":3D35D
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Cambiar_Contraseña.frx":3D3DD
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Cambiar_Contraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim contraseña As String

Private Sub Command1_Click()
    If Text3.Text <> Text4.Text Then
        MsgBox "Error al digitar contraseña, verificar contraseña", vbExclamation, "Verificar contraseña"
        Text3.Text = ""
        Text4.Text = ""
        Text3.SetFocus
    Else
        On Error GoTo error
            cn.Execute "Modificar_Contraseña '" & Text1.Text & "','" & Text4.Text & "'"
            Adodc1.Refresh
            MsgBox "Su contraseña se ha modificado", vbInformation, "CLARO"
error:
        If Err.Number <> 0 Then
            MsgBox "Error al modificar contraseña", vbCritical, "Error"
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Skin1.ApplySkin hWnd
    
    Adodc1.ConnectionString = cn
    Adodc1.RecordSource = "Select * from Usuario"
    Adodc1.Refresh
    
    With Cambiar_Contraseña
        .Height = 2820
        .Width = 4875
    End With
    
End Sub
    
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.SetFocus
    End If
End Sub

