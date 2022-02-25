VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Libro_de_Ventas 
   Caption         =   "Libro de Ventas"
   ClientHeight    =   10530
   ClientLeft      =   8850
   ClientTop       =   990
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      Picture         =   "Libro_de_Ventas.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   7695
      TabIndex        =   44
      Top             =   0
      Width           =   7695
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Libro de Ventas"
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
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   8775
      End
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":1C1AD
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Salir"
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":1C5EF
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Cambiar de vendedor(a)"
      Top             =   6600
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   840
      OleObjectBlob   =   "Libro_de_Ventas.frx":1CA31
      Top             =   9000
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   34
      Top             =   6720
      Width           =   2535
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   46
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   38
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":59752
         TabIndex        =   35
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":597C8
         TabIndex        =   36
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":59840
         TabIndex        =   47
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ventas al Contado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   4080
      TabIndex        =   32
      Top             =   1920
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   2760
      Top             =   9960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc6"
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
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1560
      TabIndex        =   30
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Consulta de las ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   22
      Top             =   7800
      Width           =   10815
      Begin VB.CommandButton Command4 
         Height          =   615
         Left            =   9600
         Picture         =   "Libro_de_Ventas.frx":598C0
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Reporte de las ventas"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   8400
         Picture         =   "Libro_de_Ventas.frx":59CFE
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Reporte del detalle de las ventas"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   7200
         Picture         =   "Libro_de_Ventas.frx":5A140
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Buscar"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   5520
         TabIndex        =   26
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20054017
         CurrentDate     =   39128
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20054017
         CurrentDate     =   39128
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   4680
         OleObjectBlob   =   "Libro_de_Ventas.frx":5A582
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "Libro_de_Ventas.frx":5A5E4
         TabIndex        =   24
         Top             =   480
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":5A640
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   9840
      TabIndex        =   20
      Top             =   9840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   10920
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   9840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   8280
      Top             =   9840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6600
      Top             =   9840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4920
      Top             =   9840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13440
      TabIndex        =   18
      Top             =   7200
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   12480
      OleObjectBlob   =   "Libro_de_Ventas.frx":5A6A6
      TabIndex        =   17
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton CmdAgregar 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":5A716
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Agregar"
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton CmdModificar 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":5AB58
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Modificar"
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdEliminar 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":5AF9A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Eliminar"
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Detalle Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   2535
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   1680
         TabIndex        =   31
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":5B3DC
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":5B460
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":5B4D4
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":5B544
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   1200
      Picture         =   "Libro_de_Ventas.frx":5B5BA
      ScaleHeight     =   1575
      ScaleWidth      =   15255
      TabIndex        =   3
      Top             =   0
      Width           =   15255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seleccionar Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":77767
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20054017
         CurrentDate     =   39123
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":777D7
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "Libro_de_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim año As String
Dim mes As String
Dim dia As String
Dim num As Integer
Dim k As Integer
Dim i As Integer
Dim num2 As Integer

Dim tasa As Double
Dim com_vend As Double
Dim com_arq As Double
Dim sal_vend As Double
Dim sal_arq As Double

Public str1 As String
Public rst1 As New ADODB.Recordset

Public str2 As String
Public rst2 As New ADODB.Recordset

Public str3 As String
Public rst3 As New ADODB.Recordset

Public str4 As String
Public rst4 As New ADODB.Recordset

Dim moneda As Double
Dim contador As Integer
Dim booleano As Boolean

Public Sub Total_Comision()
    Calcular_Total_Comision_Vendedor_Contado
    Calcular_Total_Comision_Arqueador_Contado
    Calcular_Total_Comision_Administrador_Contado
End Sub

Private Sub CmdAgregar_Click()
    On Error GoTo error
        If Combo3.Text = "" Or Text2.Text = "" Or Text14.Text = "" Or Combo6.Text = "" Then
            MsgBox "Por favor complete el formulario detalle venta", vbCritical, "Error"
        Else
                    cn.Execute "Ingresar_Detalle_Venta_Contado '" & Combo3.Text & "'," & CInt(Text2.Text) & "," & CInt(Text14.Text) & "," & CInt(Combo6) & ",'" & Text8.Text & "','" & Combo5.Text & "'"
                    Adodc2.RecordSource = "Select * from Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc2.Refresh
                    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc3.Refresh
                    detalle_contado
                    busqueda_total_contado
                    booleano = True
                
            Frame5.Caption = "Total de comisiones del " & CStr(DTPicker1)
            estupidez
            Combo1.ListIndex = 0
            DTPicker2 = DTPicker1
            DTPicker3 = DTPicker1
            Obtener_Fecha_Inicial
            Obtener_Fecha_Final
            Command1_Click
        End If
error:
        If Err.Number <> 0 Then
            MsgBox "Error al agregar detalle venta", vbCritical, "Error"
            Text2.SetFocus
        End If
End Sub

Private Sub CmdEliminar_Click()
    On Error GoTo error
        'If Combo3.Text = "" Or Text2.Text = "" Or Text14.Text = "" Or Combo6.Text = "" Then
            'MsgBox "Por favor complete el formulario detalle venta", vbCritical, "Error"
        'Else
            
                    'If MSFlexGrid1.Rows = 2 Then
                        'MsgBox "No se puede eliminar el último registro", vbCritical, "Error"
                    'Else
                        num2 = InputBox("Dígite el Id Venta para eliminar el detalle de venta al contado")
                        cn.Execute "Eliminar_Detalle_Venta_Contado " & num2 & ",'" & Text8.Text & "','" & Combo5.Text & "'"
                        Adodc2.RecordSource = "Select * from Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                        Adodc2.Refresh
                        Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                        Adodc3.Refresh
                        detalle_contado
                        busqueda_total_contado
                        Command1_Click
                    'End If
            
            Frame5.Caption = "Total de comisiones del " & CStr(DTPicker1)
            estupidez
            inicializar_tablas
        'End If
error:
        If Err.Number <> 0 Then
            MsgBox "Error al eliminar detalle venta", vbCritical, "Error"
            MsgBox "Para eliminar asegurese de que la fecha principal que seleccionó sea la correcta", vbExclamation, "Aviso Importante"
        End If
End Sub

Private Sub CmdModificar_Click()
    On Error GoTo error
        If Combo3.Text = "" Or Text2.Text = "" Or Text14.Text = "" Or Combo6.Text = "" Then
            MsgBox "Por favor complete el formulario detalle venta", vbCritical, "Error"
        Else
            
                    num2 = InputBox("Dígite el Id Venta para modificar el detalle de venta al contado")
                    cn.Execute "Modificar_Detalle_Venta_Contado " & num2 & ",'" & Combo3.Text & "'," & CInt(Text2.Text) & "," & CInt(Text14.Text) & "," & CInt(Combo6) & ",'" & Text8.Text & "','" & Combo5.Text & "'"
                    Adodc2.RecordSource = "Select * from Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc2.Refresh
                    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc3.Refresh
                    detalle_contado
                    busqueda_total_contado
                    Command1_Click
                
            Frame5.Caption = "Total de comisiones del " & CStr(DTPicker1)
            estupidez
        End If
error:
        If Err.Number <> 0 Then
            MsgBox "Error al modificar detalle venta", vbCritical, "Error"
            MsgBox "Para modificar asegurese de que la fecha principal que seleccionó sea la correcta", vbExclamation, "Aviso Importante"
            Text2.SetFocus
        End If
End Sub

Private Sub Combo2_Click()
    
    Combo5.ListIndex = Combo2.ListIndex
    
End Sub

Private Sub Command1_Click()
    
    MSFlexGrid1.Clear
    
    Poner_Campo_Contado
    
    Dim seleccion As String
    
    seleccion = CStr(Combo1.Text)
    
    Select Case seleccion
        Case "Ambos"
            Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
            Adodc3.Refresh
        Case "5"
            Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "' And Descuento = 5"
            Adodc3.Refresh
    End Select
    
    detalle_contado
    
    Calcular_Total_Contado
    Total_Comision
    estupidez
    
End Sub

Public Sub Calcular_Total_Comision_Vendedor_Contado()
    
    Dim i As Integer
    Dim Comision As Double
    
    On Error GoTo tonto
        For i = 1 To MSFlexGrid1.Rows - 1
            Comision = Comision + CDbl(MSFlexGrid1.TextMatrix(i, 8))
        Next i
        
        Text1.Text = CStr(Comision)
tonto:
    If Err.Number <> 0 Then
        Text1.Text = "0.00"
    End If
    
End Sub

Public Sub Calcular_Total_Comision_Arqueador_Contado()
    
    Dim i As Integer
    Dim Comision As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid1.Rows - 1
            Comision = Comision + CDbl(MSFlexGrid1.TextMatrix(i, 9))
        Next i
        
        Text3.Text = CStr(Comision)
    
tonto:
    If Err.Number <> 0 Then
        Text3.Text = "0.00"
    End If
    
End Sub

Public Sub Calcular_Total_Comision_Administrador_Contado()
    
    Dim i As Integer
    Dim Comision As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid1.Rows - 1
            Comision = Comision + CDbl(MSFlexGrid1.TextMatrix(i, 10))
        Next i
        
        Text5.Text = CStr(Comision)
    
tonto:
    If Err.Number <> 0 Then
        Text3.Text = "0.00"
    End If
    
End Sub

Public Sub Calcular_Total_Contado()
    Dim i As Integer
    Dim tasa As Double
    Dim Total_Cordobas As Double
    
    On Error GoTo tonto
        For i = 1 To MSFlexGrid1.Rows - 1
            Total_Cordobas = Total_Cordobas + CDbl(MSFlexGrid1.TextMatrix(i, 6))
        Next i
        
        Text4.Text = CStr(Total_Cordobas)
        
tonto:
    If Err.Number <> 0 Then
        Text4.Text = "0.00"
        Text5.Text = "0.00"
    End If
    
End Sub

Private Sub Command2_Click()
        
        metodo_contado
        
        Detalle
        Habilitar
        
        If Principal.Text1.Text = "Arqueador(a)" Then
            CmdModificar.Enabled = False
            CmdEliminar.Enabled = False
        End If
        
        DTPicker2 = DTPicker1
        DTPicker3 = DTPicker1
        
        Obtener_Fecha_Inicial
        Obtener_Fecha_Final
        
        Command1_Click
        
        DTPicker1.Enabled = False
        Combo2.Enabled = False
        
        Actualizar_Adodc
        
        Command2.Enabled = False
        
End Sub

Public Sub metodo_contado()
   
    Dim X As Integer
    Dim id_venta As Integer
        
    Adodc2.RecordSource = "Select * from Venta_Contado where Fecha_Venta = '" & Text8.Text & "' And Num_Cedula = '" & Combo5.Text & "'"
    Adodc2.Refresh
    
    On Error GoTo error
        id_venta = CInt(Adodc2.Recordset("Id_Venta"))
        consulta_contado
error:
    If Err.Number <> 0 Then
        ingresar_contado
    End If
    
End Sub

Public Sub ingresar_contado()
    cn.Execute "Ingresar_Venta_Contado '" & Text8.Text & "','" & Combo5.Text & "'"
    Adodc2.Refresh
End Sub

Public Sub consulta_contado()
    Adodc2.RecordSource = "Select * from Venta_Contado where Num_Cedula = '" & Combo5.Text & "' and Fecha_Venta = '" & Text8.Text & "'"
    Adodc2.Refresh
End Sub

Private Sub Command3_Click()

On Error GoTo error

        str1 = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
        
        rst1.Open str1, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Detalle_Venta_Contado.DataSource = rst1
        rpt_Detalle_Venta_Contado.Show

error:
    If Err.Number <> 0 Then
    End If
    
End Sub

Private Sub Command4_Click()

Dim question As Integer

On Error GoTo error

        str3 = "Select * from Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
        
        rst3.Open str3, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Venta_Contado.DataSource = rst3
        rpt_Venta_Contado.Show
   
error:
    If Err.Number <> 0 Then
    End If
    
End Sub

Private Sub Command5_Click()
    Unload Me
    Libro_de_Ventas.Show
End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub DTPicker1_Change()
    Obtener_Fecha
    
    DTPicker2 = DTPicker1
    DTPicker3 = DTPicker1
    
    Obtener_Fecha_Inicial
    Obtener_Fecha_Final
    
End Sub

Private Sub DTPicker2_Change()
    Obtener_Fecha_Inicial
End Sub

Private Sub DTPicker3_Change()
    Obtener_Fecha_Final
End Sub

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    DTPicker1 = Format(Date, "short date")
    DTPicker2 = Format(Date, "short date")
    DTPicker3 = Format(Date, "short date")
    
    Me.Adodc1.ConnectionString = cn
    Adodc1.RecordSource = "Select * from Vendedor"
    Adodc1.Refresh
    
    Me.Adodc2.ConnectionString = cn
    Adodc2.RecordSource = "Select * from Venta_Contado"
    Adodc2.Refresh
    
    Me.Adodc3.ConnectionString = cn
    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado"
    Adodc3.Refresh
    
    busqueda_vendedor
    
    With Combo3
        .AddItem "20"
        .AddItem "30"
        .AddItem "50"
        .AddItem "100"
        .ListIndex = 0
    End With
    
    With Combo6
        .AddItem "5"
        .ListIndex = 0
    End With
    
    With Combo1
        .AddItem "5"
        .ListIndex = 0
    End With
    
    Deshabilitar
    
    Poner_Campo_Contado
    
    Obtener_Fecha
    Obtener_Fecha_Inicial
    Obtener_Fecha_Final
    
    Tamaño_Campo_Contado
    
    Frame5.Caption = "Total de comisiones del " & CStr(DTPicker1)
    
    Combo5.ListIndex = Combo2.ListIndex
    
    Combo6.Enabled = False
    Combo1.Enabled = False
    
End Sub

Public Sub inicializar_tablas()
    MSFlexGrid1.Clear
    
    Poner_Campo_Contado
    
    Tamaño_Campo_Contado
    
    detalle_contado
    busqueda_total_contado
    
End Sub

Public Sub Actualizar_Adodc()
    
    Adodc1.RecordSource = "Select * from Vendedor"
    Adodc1.Refresh
    
    Adodc2.RecordSource = "Select * from Venta_Contado"
    Adodc2.Refresh
    
    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado"
    Adodc3.Refresh
    
End Sub

Public Sub Tamaño_Campo_Contado()
    With MSFlexGrid1
        .ColWidth(0) = 790 'Id_Venta
        .ColWidth(1) = 550 'Tipo
        .ColWidth(2) = 900 'Disponible
        .ColWidth(3) = 800 'Cantidad
        .ColWidth(4) = 970 ' Descuento
        .ColWidth(5) = 750 ' Precio
        .ColWidth(6) = 880 ' SubTotal
        .ColWidth(7) = 900 ' Devolucion
        .ColWidth(8) = 1500 ' Comision_Vendedor
        .ColWidth(9) = 1500 ' Comision_Arqueador
        .ColWidth(10) = 1500 ' Comision_Admon
        .ColWidth(11) = 1000 ' Fecha_Venta
    End With
End Sub

Public Sub Poner_Campo_Contado()
    With MSFlexGrid1
        .TextMatrix(0, 0) = "Id Venta"
        .TextMatrix(0, 1) = "Tipo"
        .TextMatrix(0, 2) = "Disponible"
        .TextMatrix(0, 3) = "Cantidad"
        .TextMatrix(0, 4) = "Descuento"
        .TextMatrix(0, 5) = "Precio"
        .TextMatrix(0, 6) = "Sub-total"
        .TextMatrix(0, 7) = "Devolución"
        .TextMatrix(0, 8) = "Com. Vendedor"
        .TextMatrix(0, 9) = "Com. Arqueador"
        .TextMatrix(0, 10) = "Com. Administrador"
        .TextMatrix(0, 11) = "Fecha Venta"
    End With
End Sub

Public Sub Deshabilitar()
    Frame4.Enabled = False
    Frame6.Enabled = False
    CmdAgregar.Enabled = False
    CmdModificar.Enabled = False
    CmdEliminar.Enabled = False
End Sub

Public Sub Habilitar()
    Frame4.Enabled = True
    Frame6.Enabled = True
    CmdAgregar.Enabled = True
    CmdModificar.Enabled = True
    CmdEliminar.Enabled = True
End Sub

Public Sub busqueda_vendedor()
    
    Dim X As Integer

    If Adodc1.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc1.Recordset.MoveFirst
    
            X = 1
    
            Do While X < Adodc1.Recordset.RecordCount + 1
                Combo2.AddItem CStr(Adodc1.Recordset.Fields("I_Nombre")) & " " & CStr(Adodc1.Recordset.Fields("II_Nombre")) & " " & CStr(Adodc1.Recordset.Fields("I_Apellido")) & " " & CStr(Adodc1.Recordset.Fields("II_Apellido"))
                Combo5.AddItem CStr(Adodc1.Recordset.Fields("Num_Cedula"))
                
                If Adodc1.Recordset.Bookmark <> Adodc1.Recordset.RecordCount Then
                    Adodc1.Recordset.MoveNext
                End If
                
                X = X + 1
            Loop
            Combo2.ListIndex = 0
            Combo5.ListIndex = 0
    End If
    
End Sub

Public Sub detalle_contado()
    
    Dim X As Integer

    If Adodc3.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc3.Recordset.MoveFirst
    
            MSFlexGrid1.Rows = CInt(Adodc3.Recordset.RecordCount) + 1
            X = 1
    
            Do While X < Adodc3.Recordset.RecordCount + 1
                MSFlexGrid1.TextMatrix(X, 0) = Adodc3.Recordset.Fields("Id_Venta")
                MSFlexGrid1.TextMatrix(X, 1) = Adodc3.Recordset.Fields("Tipo")
                MSFlexGrid1.TextMatrix(X, 2) = Adodc3.Recordset.Fields("Disponible")
                MSFlexGrid1.TextMatrix(X, 3) = Adodc3.Recordset.Fields("Cantidad")
                MSFlexGrid1.TextMatrix(X, 4) = Adodc3.Recordset.Fields("Descuento")
                MSFlexGrid1.TextMatrix(X, 5) = Adodc3.Recordset.Fields("Precio")
                MSFlexGrid1.TextMatrix(X, 6) = Adodc3.Recordset.Fields("SubTotal")
                MSFlexGrid1.TextMatrix(X, 7) = Adodc3.Recordset.Fields("Devolucion")
                MSFlexGrid1.TextMatrix(X, 8) = Adodc3.Recordset.Fields("Comision_Vendedor")
                MSFlexGrid1.TextMatrix(X, 9) = Adodc3.Recordset.Fields("Comision_Arqueador")
                MSFlexGrid1.TextMatrix(X, 10) = Adodc3.Recordset.Fields("Comision_Administrador")
                MSFlexGrid1.TextMatrix(X, 11) = Adodc3.Recordset.Fields("Fecha_Venta")
                
                If Adodc3.Recordset.Bookmark <> Adodc3.Recordset.RecordCount Then
                    Adodc3.Recordset.MoveNext
                End If
                
                X = X + 1
            Loop
    End If

End Sub

Public Sub Obtener_Fecha()
    
    año = CStr(DTPicker1.Year)
    
    If DTPicker1.Month < 10 Then
        mes = "0" & CStr(DTPicker1.Month)
    Else
        mes = CStr(DTPicker1.Month)
    End If
    
    If DTPicker1.Day < 10 Then
        dia = "0" & CStr(DTPicker1.Day)
    Else
        dia = CStr(DTPicker1.Day)
    End If
    
    Text8.Text = año & "-" & mes & "-" & dia
    
End Sub

Public Sub busqueda_total_contado()
    
    Dim X As Integer

    If Adodc2.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc2.Recordset.MoveFirst
    
            X = 1
    
            Do While X < Adodc2.Recordset.RecordCount + 1
                
                'Text1.Text = CStr(Adodc2.Recordset.Fields("Total_Comision_Vendedor"))
                'Text3.Text = CStr(Adodc2.Recordset.Fields("Total_Comision_Arqueador"))
                'Text5.Text = CStr(Adodc2.Recordset.Fields("Total_Comision_Administrador"))
                
                'Text4.Text = CStr(Adodc2.Recordset.Fields("Total_C$"))
                  
                If Adodc2.Recordset.Bookmark <> Adodc2.Recordset.RecordCount Then
                    Adodc2.Recordset.MoveNext
                End If
                
                X = X + 1
            Loop
    End If
    
End Sub

Public Sub Detalle()
    
    Obtener_Fecha
    
    Adodc2.RecordSource = "Select * from Venta_Contado where Num_Cedula = '" & Combo5.Text & "' And Fecha_Venta = '" & Text8.Text & "'"
    Adodc2.Refresh
    
    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Num_Cedula = '" & Combo5.Text & "' And Fecha_Venta = '" & Text8.Text & "'"
    Adodc3.Refresh
    
    detalle_contado
    busqueda_total_contado
    
End Sub

Public Sub Obtener_Fecha_Inicial()
    
    año = CStr(DTPicker2.Year)
    
    If DTPicker2.Month < 10 Then
        mes = "0" & CStr(DTPicker2.Month)
    Else
        mes = CStr(DTPicker2.Month)
    End If
    
    If DTPicker2.Day < 10 Then
        dia = "0" & CStr(DTPicker2.Day)
    Else
        dia = CStr(DTPicker2.Day)
    End If
    
    Text12.Text = año & "-" & mes & "-" & dia
    
End Sub

Public Sub Obtener_Fecha_Final()
    
    año = CStr(DTPicker3.Year)
    
    If DTPicker3.Month < 10 Then
        mes = "0" & CStr(DTPicker3.Month)
    Else
        mes = CStr(DTPicker3.Month)
    End If
    
    If DTPicker3.Day < 10 Then
        dia = "0" & CStr(DTPicker3.Day)
    Else
        dia = CStr(DTPicker3.Day)
    End If
    
    Text13.Text = año & "-" & mes & "-" & dia
    
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If InStr("1234567890-", Chr(KeyAscii)) = 0 And KeyAscii > 13 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If InStr("1234567890-", Chr(KeyAscii)) = 0 And KeyAscii > 13 Then
    KeyAscii = 0
End If
End Sub

Public Sub estupidez()
    Text1.Enabled = True
    Text3.Enabled = True
    Text5.Enabled = True
    
    Text1.SetFocus
    Text3.SetFocus
    Text5.SetFocus
    
    Text1.Enabled = False
    Text3.Enabled = False
    Text5.Enabled = False
    
    Text14.Text = ""
    Text2.Text = ""
    
    Text2.SetFocus
    
End Sub
