VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Crear_Usuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear usuario"
   ClientHeight    =   3705
   ClientLeft      =   870
   ClientTop       =   630
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7320
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   240
      Picture         =   "Crear_Usuario.frx":0000
      ScaleHeight     =   1935
      ScaleWidth      =   2055
      TabIndex        =   13
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton CmdAgregar 
      Height          =   615
      Left            =   6000
      Picture         =   "Crear_Usuario.frx":1139
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Agregar"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton CmdEliminar 
      Height          =   615
      Left            =   6000
      Picture         =   "Crear_Usuario.frx":157B
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Eliminar"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3960
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "MOVISTAR"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "sysadmin"
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      Picture         =   "Crear_Usuario.frx":19BD
      ScaleHeight     =   795
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   2520
      OleObjectBlob   =   "Crear_Usuario.frx":245D
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   2520
      OleObjectBlob   =   "Crear_Usuario.frx":24D1
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   2520
      OleObjectBlob   =   "Crear_Usuario.frx":253D
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Crear_Usuario.frx":25A5
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   2520
      OleObjectBlob   =   "Crear_Usuario.frx":3F2C6
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Crear_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdEliminar_Click()
    On Error GoTo error
        If (Text1 <> "" And Text2 <> "") Then
            cn.Execute "sp_droplogin '" & Text1 & "'"
            cn.Execute "Eliminar_Usuario " & CInt(Text5.Text)
            Adodc1.Refresh
            MsgBox "El usuario ha sido eliminado", vbInformation, "Confirmacion"
            Unload Me
        Else
            MsgBox "Faltan campos por llenar", vbInformation, "Confirmacion"
        End If
error:
    If Err.Number <> 0 Then
        MsgBox "Error al eliminar usuario", vbCritical, "Error"
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error GoTo error
            CmdAgregar.SetFocus
error:
        If Err.Number <> 0 Then
            CmdEliminar.SetFocus
        End If
    End If
End Sub

Private Sub CmdAgregar_Click()
    On Error GoTo error
        If (Text1 <> "" And Text2 <> "") Then
            cn.Execute "sp_addlogin '" & Text1 & "','" & Text2 & "','" & Text3 & "'"
            cn.Execute "sp_addsrvrolemember '" & Text1 & "'," & Text4.Text
            cn.Execute "Crear_Usuario '" & Text1 & "','" & Text2 & "','" & Combo1 & "'"
            Adodc1.Refresh
            MsgBox "El usuario ha sido creado", vbInformation, "Confirmacion"
            Unload Me
        Else
            MsgBox "Faltan campos por llenar", vbInformation, "Confirmacion"
        End If
error:
    If Err.Number <> 0 Then
        MsgBox "El nombre del usuario ya existe", vbCritical, "Error"
    End If
End Sub
       
Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    
    Crear_Usuario.Height = 4185
    Crear_Usuario.Width = 7410
     
    Me.Adodc1.ConnectionString = cn
    Me.Adodc1.RecordSource = "select * from Usuario "
    Me.Adodc1.Refresh
    
    With Combo1
        .AddItem "Administrador"
        .AddItem "Arqueador(a)"
        .ListIndex = 0
    End With
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub
