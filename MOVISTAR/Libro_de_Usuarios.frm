VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Libro_de_Usuarios 
   Caption         =   "Libro de usuarios"
   ClientHeight    =   5925
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   10290
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   -3840
      Picture         =   "Libro_de_Usuarios.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   9495
      TabIndex        =   7
      Top             =   0
      Width           =   9495
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Left            =   9480
         ScaleHeight     =   1095
         ScaleWidth      =   15
         TabIndex        =   8
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Libro de Usuarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3960
         TabIndex        =   9
         Top             =   600
         Width           =   8775
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   3960
      Picture         =   "Libro_de_Usuarios.frx":1C1AD
      ScaleHeight     =   1095
      ScaleWidth      =   6375
      TabIndex        =   6
      Top             =   0
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   9000
      Picture         =   "Libro_de_Usuarios.frx":1D0B4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar"
      Top             =   4440
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   4440
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   5400
      OleObjectBlob   =   "Libro_de_Usuarios.frx":1D4F6
      TabIndex        =   2
      Top             =   4560
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   1440
      OleObjectBlob   =   "Libro_de_Usuarios.frx":1D55A
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1200
      Top             =   6360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "Libro_de_Usuarios.frx":1D5C4
      Top             =   6360
   End
End
Attribute VB_Name = "Libro_de_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer

Private Sub Command1_Click()

    Dim seleccion As Integer
    
    seleccion = Combo1.ListIndex
    
    Select Case seleccion
        Case 0
            BuscarPor_Id_Usuario
        Case 1
            BuscarPor_Nombre
        Case 2
            BuscarPor_Contraseña
        Case 3
            BuscarPor_Tipo
    End Select
    
End Sub

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    
    Adodc1.ConnectionString = cn
    Adodc1.RecordSource = "Select * from Usuario"
    Adodc1.Refresh
    
    Dim Gonzalez As Integer
    
    For Gonzalez = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.ColWidth(Gonzalez) = 2000
    Next Gonzalez
    
    With Libro_de_Usuarios
        .Height = 6390
        .Width = 10410
    End With
        
    With Combo1
        .AddItem "Id_Usuario"
        .AddItem "Nombre"
        .AddItem "Contraseña"
        .AddItem "Tipo"
        .ListIndex = 0
    End With
    
    Poner_Campo
    
    frank
        
End Sub

Public Sub Poner_Campo()
    With MSFlexGrid1
        .TextMatrix(0, 0) = "Id_Usuario"
        .TextMatrix(0, 1) = "Nombre"
        .TextMatrix(0, 2) = "Contraseña"
        .TextMatrix(0, 3) = "Tipo"
    End With
End Sub

Public Sub BuscarPor_Id_Usuario()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Usuario where Id_Usuario = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No Existe Ningún Registro"
    Else
        With Crear_Usuario
            .Text5.Text = MSFlexGrid1.TextMatrix(1, 0)
            .Text1.Text = MSFlexGrid1.TextMatrix(1, 1)
            .Text2.Text = MSFlexGrid1.TextMatrix(1, 2)
            .Combo1.Text = MSFlexGrid1.TextMatrix(1, 3)
            .CmdAgregar.Visible = False
            
            .Text5.Enabled = False
            .Text1.Enabled = False
            .Text2.Enabled = False
            .Combo1.Enabled = False
            
            .Show
        End With
        Unload Me
    End If
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_Nombre()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Usuario where Nombre = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No Existe Ningún Registro"
    Else
        With Crear_Usuario
            .Text5.Text = MSFlexGrid1.TextMatrix(1, 0)
            .Text1.Text = MSFlexGrid1.TextMatrix(1, 1)
            .Text2.Text = MSFlexGrid1.TextMatrix(1, 2)
            .Combo1.Text = MSFlexGrid1.TextMatrix(1, 3)
            .CmdAgregar.Visible = False
            
            .Text5.Enabled = False
            .Text1.Enabled = False
            .Text2.Enabled = False
            .Combo1.Enabled = False
            
            .Show
        End With
        Unload Me
    End If
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_Contraseña()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Usuario where Contraseña = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_Tipo()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Usuario where Tipo = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub frank()
    
    Dim X As Integer

    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "No Existe Ningún Registro"
    Else
            Adodc1.Recordset.MoveFirst
    
            MSFlexGrid1.Rows = CInt(Adodc1.Recordset.RecordCount) + 1
            X = 1
    
            Do While X < Adodc1.Recordset.RecordCount + 1
                MSFlexGrid1.TextMatrix(X, 0) = Adodc1.Recordset.Fields("Id_Usuario")
                MSFlexGrid1.TextMatrix(X, 1) = Adodc1.Recordset.Fields("Nombre")
                MSFlexGrid1.TextMatrix(X, 2) = Adodc1.Recordset.Fields("Contraseña")
                MSFlexGrid1.TextMatrix(X, 3) = Adodc1.Recordset.Fields("Tipo")
                
                If Adodc1.Recordset.Bookmark <> Adodc1.Recordset.RecordCount Then
                    Adodc1.Recordset.MoveNext
                End If
                
                X = X + 1
            Loop
    End If

End Sub


