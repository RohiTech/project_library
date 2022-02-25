VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Libro_de_Vendedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Vendedores"
   ClientHeight    =   5685
   ClientLeft      =   2715
   ClientTop       =   1725
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10425
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   -3720
      Picture         =   "Libro_de_Vendedores.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   9495
      TabIndex        =   8
      Top             =   0
      Width           =   9495
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Left            =   9480
         ScaleHeight     =   1095
         ScaleWidth      =   15
         TabIndex        =   9
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Libro de Vendedores"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   600
         Width           =   8775
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4080
      Picture         =   "Libro_de_Vendedores.frx":1C1AD
      ScaleHeight     =   1095
      ScaleWidth      =   6375
      TabIndex        =   7
      Top             =   0
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   7800
      Picture         =   "Libro_de_Vendedores.frx":1D0B4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Buscar"
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   9000
      Picture         =   "Libro_de_Vendedores.frx":1D4F6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Reporte del libro de vendedores"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1200
      Top             =   6720
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4320
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   4320
      OleObjectBlob   =   "Libro_de_Vendedores.frx":1D938
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Libro_de_Vendedores.frx":1D99C
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "Libro_de_Vendedores.frx":1DA06
      Top             =   6720
   End
End
Attribute VB_Name = "Libro_de_Vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer

Dim identificacion As String

Public str As String
Public rst As New ADODB.Recordset

Private Sub Command1_Click()

    Dim seleccion As Integer
    
    seleccion = Combo1.ListIndex
    
    Select Case seleccion
        Case 0
            BuscarPor_Num_Cedula
        Case 1
            BuscarPor_Sexo
        Case 2
            BuscarPor_Edad
        Case 3
            BuscarPor_1erNombre
        Case 4
            BuscarPor_2doNombre
        Case 5
            BuscarPor_1erApellido
        Case 6
            BuscarPor_2doApellido
        Case 7
            BuscarPor_Estado_Civil
        Case 8
            BuscarPor_Direccion
    End Select
    
End Sub

Private Sub Command3_Click()
           
        Dim j As Variant
            j = MsgBox("Desea mostrar todos los registros", vbYesNo, "Confirmación")
    
        If j = vbYes Then
            str = "Select * from Vendedor"
        Else
            identificacion = InputBox("Digite el No. Cedula del vendedor")
            On Error GoTo error
                str = "Select * from Vendedor where Num_Cedula = '" & identificacion & "'"
error:
            If Err.Number <> 0 Then
                str = "Select * from Num_Cedula"
            End If
        End If
        
        rst.Open str, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Vendedores.DataSource = rst
        rpt_Vendedores.Show
        
End Sub

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    
    Adodc1.ConnectionString = cn
    Adodc1.RecordSource = "Select * from Vendedor"
    Adodc1.Refresh
    
    Dim francisco As Integer
    
    For francisco = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.ColWidth(francisco) = 2000
    Next francisco
    
    With Libro_de_Vendedores
        .Height = 6120
        .Width = 10515
    End With
        
    With Combo1
        .AddItem "No. Cedula"
        .AddItem "Sexo"
        .AddItem "Edad"
        .AddItem "1er Nombre"
        .AddItem "2do Nombre"
        .AddItem "1er Apellido"
        .AddItem "2do Apellido"
        .AddItem "Estado Civil"
        .AddItem "Direccion"
        .ListIndex = 0
    End With
    
    Poner_Campo
    
    frank
    
    With Libro_de_Vendedores
        .Height = 5895
        .Width = 10515
    End With
        
End Sub

Public Sub Poner_Campo()
    With MSFlexGrid1
        .TextMatrix(0, 0) = "No. Cedula"
        .TextMatrix(0, 1) = "Sexo"
        .TextMatrix(0, 2) = "Edad"
        .TextMatrix(0, 3) = "1er Nombre"
        .TextMatrix(0, 4) = "2do Nombre"
        .TextMatrix(0, 5) = "1er Apellido"
        .TextMatrix(0, 6) = "2do Apellido"
        .TextMatrix(0, 7) = "Estado Civil"
        .TextMatrix(0, 8) = "Dirección"
    End With
End Sub

Public Sub BuscarPor_Num_Cedula()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where Num_Cedula = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
    If Adodc1.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
        With Vendedor
            .MaskEdBox1.Text = MSFlexGrid1.TextMatrix(1, 0)
            .Text7.Text = MSFlexGrid1.TextMatrix(1, 1)
            .Text1.Text = MSFlexGrid1.TextMatrix(1, 2)
            .Text2.Text = MSFlexGrid1.TextMatrix(1, 3)
            .Text3.Text = MSFlexGrid1.TextMatrix(1, 4)
            .Text4.Text = MSFlexGrid1.TextMatrix(1, 5)
            .Text5.Text = MSFlexGrid1.TextMatrix(1, 6)
            .Combo1.Text = MSFlexGrid1.TextMatrix(1, 7)
            .Text6.Text = MSFlexGrid1.TextMatrix(1, 8)
            .MaskEdBox1.Enabled = False
            .CmdAgregar.Enabled = False
            .Show
        End With
        Unload Me
    End If
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_Sexo()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where Sexo = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_Edad()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where Edad = " & CInt(Text1.Text) & ""
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_1erNombre()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where I_Nombre = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_2doNombre()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where II_Nombre = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_1erApellido()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where I_Apellido = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_2doApellido()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where II_Apellido = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_Estado_Civil()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where Estado_Civil = '" & Text1.Text & "'"
    Me.Adodc1.Refresh
    Poner_Campo
    frank
error1:
    If Err.Number <> 0 Then
        MsgBox "Error al realizar busqueda", vbCritical, "Error"
    End If
End Sub

Public Sub BuscarPor_Direccion()
On Error GoTo error1
    MSFlexGrid1.Clear
    Me.Adodc1.RecordSource = "select * from Vendedor where Direccion = '" & Text1.Text & "'"
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
                MSFlexGrid1.TextMatrix(X, 0) = Adodc1.Recordset.Fields("Num_Cedula")
                MSFlexGrid1.TextMatrix(X, 1) = Adodc1.Recordset.Fields("Sexo")
                MSFlexGrid1.TextMatrix(X, 2) = Adodc1.Recordset.Fields("Edad")
                MSFlexGrid1.TextMatrix(X, 3) = Adodc1.Recordset.Fields("I_Nombre")
                MSFlexGrid1.TextMatrix(X, 4) = Adodc1.Recordset.Fields("II_Nombre")
                MSFlexGrid1.TextMatrix(X, 5) = Adodc1.Recordset.Fields("I_Apellido")
                MSFlexGrid1.TextMatrix(X, 6) = Adodc1.Recordset.Fields("II_Apellido")
                MSFlexGrid1.TextMatrix(X, 7) = Adodc1.Recordset.Fields("Estado_Civil")
                MSFlexGrid1.TextMatrix(X, 8) = Adodc1.Recordset.Fields("Direccion")
                
                If Adodc1.Recordset.Bookmark <> Adodc1.Recordset.RecordCount Then
                    Adodc1.Recordset.MoveNext
                End If
                
                X = X + 1
            Loop
    End If

End Sub


