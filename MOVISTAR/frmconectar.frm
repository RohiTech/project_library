VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmconectar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexión"
   ClientHeight    =   3450
   ClientLeft      =   11985
   ClientTop       =   8280
   ClientWidth     =   6225
   Icon            =   "frmconectar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "frmconectar.frx":0442
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdconectar 
      Caption         =   "&Conectar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "frmconectar.frx":04A8
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1320
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.TextBox txtbasedatos 
      BackColor       =   &H00E0E0E0&
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtservidor 
      BackColor       =   &H00E0E0E0&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtcontraseña 
      BackColor       =   &H00E0E0E0&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtusuario 
      BackColor       =   &H00E0E0E0&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   -120
      ScaleHeight     =   3435
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   0
      Width           =   1695
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   1215
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   2143
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
      End
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   0
         Top             =   480
      End
      Begin VB.Line Line1 
         X1              =   960
         X2              =   960
         Y1              =   960
         Y2              =   2160
      End
      Begin VB.Image Image2 
         Height          =   1215
         Left            =   240
         Picture         =   "frmconectar.frx":3D1C9
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   600
         Picture         =   "frmconectar.frx":3DCF1
         Top             =   120
         Width           =   645
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "frmconectar.frx":3E23E
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "frmconectar.frx":3E2AA
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "frmconectar.frx":3E312
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmconectar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valor1 As Integer

Private Sub cmdconectar_Click()
Timer1.Enabled = True
Line1.Visible = False
ProgressBar1.Visible = True
ProgressBar1.Value = 100
valor1 = 100
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()

Timer1.Enabled = False

Skin1.ApplySkin hWnd

txtusuario.Text = ""
txtcontraseña.Text = ""
txtservidor.Text = UCase(Winsock1.LocalHostName) & "\SISTEMA"
txtservidor.Enabled = False
txtbasedatos.Text = "MOVISTAR"

txtusuario.Enabled = False
txtcontraseña.Enabled = False
txtservidor.Enabled = False
txtbasedatos.Enabled = False

'cmdconectar_Click

End Sub

Private Sub conexion()
Dim tipo As String
Dim contraseña As String
On Error GoTo er
conectando

If banderas = 1 Then
If (txtusuario <> "") Then
Me.Adodc1.ConnectionString = cadena
Adodc1.RecordSource = "select * from Usuario where Nombre = '" & txtusuario & "'"
Adodc1.Refresh
tipo = Adodc1.Recordset("Tipo")

With Principal
    .Text1.Text = CStr(tipo)
    Select Case tipo
        Case "Administrador(a)"
            ' Puede hacer todo
        Case "Arqueador(a)"
            '.mnuVendedor.Item(1).Enabled = False
            '.mnuUsuario.Item(13).Enabled = False
            .mnuLibroUsuarios.Item(14).Enabled = False
            .mnuCrearUsuario.Item(15).Enabled = False
            .mnuRestaurarBD.Item(22).Enabled = False
        Case "Vendedor(a)"
    End Select
    .Show
End With

End If
End If
Timer1.Enabled = False
Unload Me

er:
If Err.Number <> 0 Then
Timer1.Enabled = False
  MsgBox "No se puede realizar la conexion", vbInformation, "AVISO"
End If
End Sub


Private Sub Timer1_Timer()
valor1 = valor1 + 1
If valor1 > ProgressBar1.Max Then
 conexion

End If
 If valor1 <= 100 Then
  ProgressBar1.Value = valor1
 End If
End Sub

Private Sub Timer2_Timer()
    
    If frmconectar.Caption = "Conexión" Then
        frmconectar.Caption = "Connection"
    Else
        frmconectar.Caption = "Conexión"
    End If
    
End Sub

Private Sub txtcontraseña_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdconectar.SetFocus
    End If
End Sub

Private Sub txtservidor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdconectar.SetFocus
    End If
End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtcontraseña.SetFocus
    End If
End Sub
