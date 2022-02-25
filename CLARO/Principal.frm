VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de control "" CLARO """
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11145
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   11085
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11145
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   8880
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OLE OLE4 
         Class           =   "Package"
         Height          =   375
         Left            =   2520
         OleObjectBlob   =   "Principal.frx":0000
         SourceDoc       =   "D:\UNI\CLARO\mario.exe"
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.OLE OLE1 
         Class           =   "Package"
         Height          =   375
         Left            =   360
         OleObjectBlob   =   "Principal.frx":1A18
         SourceDoc       =   "D:\UNI\CLARO\Pawn2\Pawn.exe"
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE OLE2 
         Class           =   "Package"
         Height          =   375
         Left            =   1080
         OleObjectBlob   =   "Principal.frx":3430
         SourceDoc       =   "D:\UNI\CLARO\Ahorcado.EXE"
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE OLE3 
         Class           =   "Package"
         Height          =   375
         Left            =   1800
         OleObjectBlob   =   "Principal.frx":68E48
         SourceDoc       =   "D:\UNI\CLARO\RompeCabeza.exe"
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE OLE5 
         Class           =   "Package"
         Height          =   375
         Left            =   3240
         OleObjectBlob   =   "Principal.frx":16FE60
         SourceDoc       =   "D:\UNI\CLARO\Super Mario Epic\Super Mario Epic.exe"
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE OLE6 
         Class           =   "PowerPoint.Show.8"
         Height          =   375
         Left            =   6600
         OleObjectBlob   =   "Principal.frx":171A78
         SourceDoc       =   "D:\UNI\CLARO\Definiciones de los iconos.pps"
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   840
      OleObjectBlob   =   "Principal.frx":199890
      Top             =   5520
   End
   Begin VB.Menu mnuVendedor 
      Caption         =   "Vendedor"
      Index           =   1
      Begin VB.Menu mnuLibroVendedores 
         Caption         =   "Libro de Vendedores"
         Index           =   2
      End
      Begin VB.Menu mnuVerVendedor 
         Caption         =   "Ver Vendedor"
         Index           =   3
      End
      Begin VB.Menu mnuAgregarVendedor 
         Caption         =   "Agregar Vendedor"
         Index           =   4
      End
      Begin VB.Menu mnuModificarVendedor 
         Caption         =   "Modificar Vendedor"
         Index           =   5
      End
      Begin VB.Menu mnuEliminarVendedor 
         Caption         =   "Eliminar Vendedor"
         Index           =   6
      End
   End
   Begin VB.Menu mnuVenta 
      Caption         =   "Venta"
      Index           =   7
      Begin VB.Menu mnuLibroVentas 
         Caption         =   "Libro de Ventas"
         Index           =   8
      End
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "Sistema de control"
      Index           =   9
      Begin VB.Menu mnuFactura 
         Caption         =   "Realizar Factura"
         Index           =   12
      End
   End
   Begin VB.Menu mnuCambioMoneda 
      Caption         =   "Cambio"
      Index           =   40
   End
   Begin VB.Menu mnuUsuario 
      Caption         =   "Usuario"
      Index           =   13
      Begin VB.Menu mnuLibroUsuarios 
         Caption         =   "Libro de usuarios"
         Index           =   14
      End
      Begin VB.Menu mnuCrearUsuario 
         Caption         =   "Crear Usuario"
         Index           =   15
      End
      Begin VB.Menu mnuCambiarUsuario 
         Caption         =   "Cambiar Usuario"
         Index           =   16
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "Herramientas"
      Index           =   17
      Begin VB.Menu mnuCalculadora 
         Caption         =   "Calculadora"
         Index           =   18
      End
      Begin VB.Menu mnuCalendario 
         Caption         =   "Calendario"
         Index           =   19
      End
   End
   Begin VB.Menu mnuRespaldo 
      Caption         =   "Respaldo"
      Index           =   20
      Begin VB.Menu mnuGenerarBackup 
         Caption         =   "Generar Backup"
         Index           =   21
      End
      Begin VB.Menu mnuRestaurarBD 
         Caption         =   "Restaurar BD"
         Index           =   22
      End
   End
   Begin VB.Menu mnuAdministrador 
      Caption         =   "Administrador"
      Index           =   23
      Begin VB.Menu mnuLibroVentasAdmon 
         Caption         =   "Libro de Ventas"
         Index           =   24
      End
      Begin VB.Menu mnuFacturaAdmon 
         Caption         =   "Realizar Factura"
         Index           =   26
      End
   End
   Begin VB.Menu mnuEntretenimiento 
      Caption         =   "Entretenimiento"
      Index           =   27
      Begin VB.Menu mnuAhorcado 
         Caption         =   "El Ahorcado"
         Index           =   29
      End
      Begin VB.Menu mnuRompeCabeza 
         Caption         =   "Rompe Cabeza"
         Index           =   30
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Index           =   33
      Begin VB.Menu mnuAgradecimiento 
         Caption         =   "Agradecimiento"
         Index           =   35
      End
      Begin VB.Menu mnuAutor 
         Caption         =   "Autor"
         Index           =   36
      End
      Begin VB.Menu mnuContacto 
         Caption         =   "Contacto"
         Index           =   37
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de..."
         Index           =   38
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Salir"
      Index           =   39
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Skin1.ApplySkin hWnd
End Sub

Private Sub mnuAcercaDe_Click(Index As Integer)
    Acerca.Show
End Sub

Private Sub mnuAgradecimiento_Click(Index As Integer)
    Agradecimiento.Show
End Sub

Private Sub mnuAgregarVendedor_Click(Index As Integer)

    With Vendedor
        .Show
        .Option1.Value = True
        .Text7.Text = "Masculino"
        .CmdModificar.Enabled = False
        .CmdEliminar.Enabled = False
        .CmdAnterior.Enabled = False
        .CmdPrimero.Enabled = False
        .CmdSiguiente.Enabled = False
        .CmdUltimo.Enabled = False
        .Frame1.Enabled = True
        .Text1.Visible = False
        .SkinLabel3.Visible = False
        .MaskEdBox1.SetFocus
    End With
        
End Sub

Private Sub mnuAhorcado_Click(Index As Integer)
    OLE2.DoVerb
End Sub

Private Sub mnuAjedrez_Click(Index As Integer)
    OLE1.DoVerb
End Sub

Private Sub mnuAutor_Click(Index As Integer)
    Autor.Show
End Sub

Private Sub mnuCalculadora_Click(Index As Integer)
    Dim a As Variant
Dim garcia As Variant
On Error GoTo error

a = Shell("C:\WINDOWS\system32\calc.exe")
error:

If Err.Number <> garcia Then
  MsgBox "Aplicacion no disponible", vbInformation, "AVISO"
End If
End Sub

Private Sub mnuCalendario_Click(Index As Integer)
    Calendario.Show
End Sub

Private Sub mnuCambiarUsuario_Click(Index As Integer)
    cn.Close
    Unload Me
    frmconectar.Show
End Sub

Private Sub mnuCambioMoneda_Click(Index As Integer)
    Cambio_Moneda.Show
End Sub

Private Sub mnuContacto_Click(Index As Integer)
    Contacto.Show
End Sub

Private Sub mnuContenido_Click(Index As Integer)
    OLE6.DoVerb
End Sub

Private Sub mnuCrearUsuario_Click(Index As Integer)
    Crear_Usuario.Show
End Sub

Private Sub mnuEliminarVendedor_Click(Index As Integer)
        
    If Vendedor.Adodc1.Recordset.RecordCount = 0 Then
        Vendedor.Hide
        MsgBox "No Existe Ningún Registro"
    Else
        With Vendedor
            .Show
            .CmdAgregar.Enabled = False
            .CmdModificar.Enabled = False
            
            .Adodc1.RecordSource = "Select * from Vendedor"
            .Adodc1.Refresh
            
            .MaskEdBox1.DataField = "Num_Cedula"
            .Text7.DataField = "Sexo"
            .Text1.DataField = "Edad"
            .Text2.DataField = "I_Nombre"
            .Text3.DataField = "II_Nombre"
            .Text4.DataField = "I_Apellido"
            .Text5.DataField = "II_Apellido"
            .Combo1.DataField = "Estado_Civil"
            .Text6.DataField = "Direccion"
            
            .Frame1.Enabled = False
            
            .CmdAnterior.Enabled = True
            .CmdPrimero.Enabled = True
            .CmdSiguiente.Enabled = True
            .CmdUltimo.Enabled = True
            
            .MaskEdBox1.Enabled = False
        End With
    End If

End Sub

Private Sub mnuFactura_Click(Index As Integer)
    Factura.Show
End Sub

Private Sub mnuFacturaAdmon_Click(Index As Integer)
    Factura_Admon.Show
End Sub

Private Sub mnuGenerarBackup_Click(Index As Integer)
    Generar_Backup.Show
End Sub

Private Sub mnuLibroUsuarios_Click(Index As Integer)
    Libro_de_Usuarios.Show
End Sub

Private Sub mnuLibroVendedores_Click(Index As Integer)
    Libro_de_Vendedores.Show
End Sub

Private Sub mnuLibroVentas_Click(Index As Integer)
    Libro_de_Ventas.Show
End Sub

Private Sub mnuLibroVentasAdmon_Click(Index As Integer)
    Libro_de_Ventas_Admon.Show
End Sub

Private Sub mnuModificarVendedor_Click(Index As Integer)
        
    If Vendedor.Adodc1.Recordset.RecordCount = 0 Then
        Vendedor.Hide
        MsgBox "No Existe Ningún Registro"
    Else
        With Vendedor
            .Show
            .CmdAgregar.Enabled = False
            .CmdEliminar.Enabled = False
            
            .Adodc1.RecordSource = "Select * from Vendedor"
            .Adodc1.Refresh
            
            .MaskEdBox1.DataField = "Num_Cedula"
            .Text7.DataField = "Sexo"
            .Text1.DataField = "Edad"
            .Text2.DataField = "I_Nombre"
            .Text3.DataField = "II_Nombre"
            .Text4.DataField = "I_Apellido"
            .Text5.DataField = "II_Apellido"
            .Combo1.DataField = "Estado_Civil"
            .Text6.DataField = "Direccion"
            
            .Frame1.Enabled = True
            
            .CmdAnterior.Enabled = True
            .CmdPrimero.Enabled = True
            .CmdSiguiente.Enabled = True
            .CmdUltimo.Enabled = True
            
            .MaskEdBox1.Enabled = False
        End With
    End If
    
End Sub

Private Sub mnuRestaurarBD_Click(Index As Integer)
    Restaurar_BD.Show
End Sub

Private Sub mnuRompeCabeza_Click(Index As Integer)
    OLE3.DoVerb
End Sub

Private Sub mnuSalir_Click(Index As Integer)
    Dim j As Variant
    j = MsgBox("Esta seguro que desea salir", vbYesNo, "Confirmacion")
    
    If j = vbYes Then
        End
    Else
    End If
End Sub

Private Sub mnuSuperMarioBros_Click(Index As Integer)
    OLE4.DoVerb
End Sub

Private Sub mnuSuperMarioBrosEpic_Click(Index As Integer)
    OLE5.DoVerb
End Sub

Private Sub mnuVerVendedor_Click(Index As Integer)
    
    If Vendedor.Adodc1.Recordset.RecordCount = 0 Then
        Vendedor.Hide
        MsgBox "No Existe Ningún Registro"
    Else
        With Vendedor
            .Adodc1.RecordSource = "Select * from Vendedor"
            .Adodc1.Refresh
            
            .MaskEdBox1.DataField = "Num_Cedula"
            .Text7.DataField = "Sexo"
            .Text1.DataField = "Edad"
            .Text2.DataField = "I_Nombre"
            .Text3.DataField = "II_Nombre"
            .Text4.DataField = "I_Apellido"
            .Text5.DataField = "II_Apellido"
            .Combo1.DataField = "Estado_Civil"
            .Text6.DataField = "Direccion"
            
            .Frame1.Enabled = False
            .CmdAnterior.Enabled = True
            .CmdPrimero.Enabled = True
            .CmdSiguiente.Enabled = True
            .CmdUltimo.Enabled = True
            .CmdAgregar.Enabled = False
            .CmdModificar.Enabled = False
            .CmdEliminar.Enabled = False
            .Show
        End With
    End If
        
End Sub
