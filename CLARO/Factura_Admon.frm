VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Factura_Admon 
   Caption         =   "Sistema de Facturación del Administrador"
   ClientHeight    =   10875
   ClientLeft      =   270
   ClientTop       =   405
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10875
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   12360
      TabIndex        =   34
      Top             =   10440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   39222
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   10200
      TabIndex        =   33
      Top             =   10440
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      Picture         =   "Factura_Admon.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   7575
      TabIndex        =   31
      Top             =   0
      Width           =   7575
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de facturación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   8775
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   1080
      Picture         =   "Factura_Admon.frx":2C24
      ScaleHeight     =   1575
      ScaleWidth      =   14295
      TabIndex        =   30
      Top             =   0
      Width           =   14295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Total Com. Pendientes"
      Height          =   1455
      Left            =   10440
      TabIndex        =   27
      Top             =   7560
      Width           =   2175
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Factura_Admon.frx":5848
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Total Com. Canceladas"
      Height          =   1455
      Left            =   12840
      TabIndex        =   24
      Top             =   7560
      Width           =   2175
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Factura_Admon.frx":58B2
         TabIndex        =   26
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   9360
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Top             =   5400
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Factura_Admon.frx":591C
      Top             =   9960
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seleccionar Vendedor"
      Height          =   1695
      Left            =   10440
      TabIndex        =   16
      Top             =   1920
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "Factura_Admon.frx":4263D
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39156
      End
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
      Height          =   1695
      Left            =   10440
      TabIndex        =   11
      Top             =   4680
      Width           =   3495
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39128
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39128
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Factura_Admon.frx":426A7
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Factura_Admon.frx":42709
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   11640
      Picture         =   "Factura_Admon.frx":42765
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Reporte del detalle de las ventas"
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   10440
      Picture         =   "Factura_Admon.frx":42BA7
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Buscar"
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   12840
      Picture         =   "Factura_Admon.frx":42FE9
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancelar Factura"
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   10440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   10440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   12960
      Picture         =   "Factura_Admon.frx":4342B
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cerrar"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Facturas Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   9855
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2295
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Facturas Canceladas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   9855
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2295
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   7560
      Top             =   9960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Left            =   3240
      Top             =   9960
      Visible         =   0   'False
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
      Left            =   960
      Top             =   9960
      Visible         =   0   'False
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   6480
      OleObjectBlob   =   "Factura_Admon.frx":4386D
      TabIndex        =   21
      Top             =   5520
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "Factura_Admon.frx":438FD
      TabIndex        =   23
      Top             =   9480
      Width           =   1215
   End
End
Attribute VB_Name = "Factura_Admon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim año As String
Dim mes As String
Dim dia As String

Public str1 As String
Public rst1 As New ADODB.Recordset

Public str2 As String
Public rst2 As New ADODB.Recordset

Dim num_fact2 As String
Dim fecha As String
Dim saldo As Double
Dim com_admon As Double

Private Sub Command1_Click()

    MSFlexGrid1.Clear
    MSFlexGrid2.Clear
    
    Nombrar_Campo1
    Nombrar_Campo2

    Obtener_Fecha_Inicial
    Obtener_Fecha_Final

    Adodc1.RecordSource = "Select * from Factura_Pendiente_Admon where Fecha_Factura >= '" & Text2.Text & "' And Fecha_Factura <= '" & Text3.Text & "' And Estado_Factura = 'Pendiente'"
    Adodc1.Refresh
    
    Adodc2.RecordSource = "Select * from Factura_Pendiente_Admon where Fecha_Factura >= '" & Text2.Text & "' And Fecha_Factura <= '" & Text3.Text & "' And Estado_Factura = 'Cancelado'"
    Adodc2.Refresh

    factura_pendiente
    factura_cancelada
    Totales
    
End Sub

Public Sub factura_pendiente()
   
    Dim x As Integer

    If Adodc1.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc1.Recordset.MoveFirst
    
            MSFlexGrid1.Rows = CInt(Adodc1.Recordset.RecordCount) + 1
            x = 1
    
            Do While x < Adodc1.Recordset.RecordCount + 1
                MSFlexGrid1.TextMatrix(x, 0) = Adodc1.Recordset.Fields("Num_Factura")
                MSFlexGrid1.TextMatrix(x, 1) = Adodc1.Recordset.Fields("Fecha_Factura")
                MSFlexGrid1.TextMatrix(x, 2) = Adodc1.Recordset.Fields("Saldo_Pendiente")
                MSFlexGrid1.TextMatrix(x, 3) = Adodc1.Recordset.Fields("Comision_Admon")
                
                If Adodc1.Recordset.Bookmark <> Adodc1.Recordset.RecordCount Then
                    Adodc1.Recordset.MoveNext
                End If
                
                x = x + 1
            Loop
    End If

End Sub

Public Sub factura_cancelada()
   
    Dim x As Integer

    If Adodc2.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc2.Recordset.MoveFirst
    
            MSFlexGrid2.Rows = CInt(Adodc2.Recordset.RecordCount) + 1
            x = 1
    
            Do While x < Adodc2.Recordset.RecordCount + 1
                MSFlexGrid2.TextMatrix(x, 0) = Adodc2.Recordset.Fields("Num_Factura")
                MSFlexGrid2.TextMatrix(x, 1) = Adodc2.Recordset.Fields("Fecha_Factura")
                MSFlexGrid2.TextMatrix(x, 2) = Adodc2.Recordset.Fields("Saldo_Pendiente")
                MSFlexGrid2.TextMatrix(x, 3) = Adodc2.Recordset.Fields("Comision_Admon")
                
                If Adodc2.Recordset.Bookmark <> Adodc2.Recordset.RecordCount Then
                    Adodc2.Recordset.MoveNext
                End If
                
                x = x + 1
            Loop
    End If

End Sub


Private Sub Command2_Click()
    Despues
    
    Adodc1.RecordSource = "Select * from Factura_Pendiente_Admon where Fecha_Factura = '" & Text1.Text & "'And Estado_Factura = 'Pendiente'"
    Adodc1.Refresh
    
    Adodc2.RecordSource = "Select * from Factura_Pendiente_Admon where Fecha_Factura = '" & Text1.Text & "' And Estado_Factura = 'Cancelado'"
    Adodc2.Refresh
    
    factura_pendiente
    factura_cancelada
    Totales
End Sub

Public Sub Despues()
    DTPicker1.Enabled = False
    
    Frame6.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = True
    Command4.Enabled = True
    
    Obtener_Fecha
End Sub

Private Sub Command3_Click()

Dim question As Integer

On Error GoTo error

Command1_Click

question = InputBox("1. Factura Pendiente" & vbCrLf & "2. Factura Cancelada")

Select Case question
    Case 1
        str1 = "Select * from Factura_Pendiente_Admon where Fecha_Factura >= '" & Text2.Text & "' And Fecha_Factura <= '" & Text3.Text & "' And Estado_Factura = 'Pendiente'"
        
        rst1.Open str1, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Factura_Pendiente_Admon.DataSource = rst1
        rpt_Factura_Pendiente_Admon.Show
    Case 2
        str1 = "Select * from Factura_Pendiente_Admon where Fecha_Factura >= '" & Text2.Text & "' And Fecha_Factura <= '" & Text3.Text & "' And Estado_Factura = 'Cancelado'"
        
        rst1.Open str1, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Factura_Pendiente_Admon.DataSource = rst1
        rpt_Factura_Pendiente_Admon.Show
End Select

error:
    If Err.Number <> 0 Then
    End If
    
End Sub

Private Sub Command4_Click()

    On Error GoTo error
    
        Dim numfact As String
        
        numfact = InputBox("Digite el numero de factura [#####]: ")
        
        cn.Execute "Modificar_Factura_Pendiente_Admon '" & numfact & "','" & CStr(Text7.Text) & "'"
        Adodc1.Refresh
        
        MSFlexGrid1.Clear
        Nombrar_Campo1
        Tamaño1
        
        factura_pendiente
        factura_cancelada
        Command1_Click
error:
    If Err.Number <> 0 Then
        MsgBox "Error al cancelar factura", vbCritical, "Error"
    End If
    
End Sub

Public Sub ingresar_factura_cancelada()

    cn.Execute "Ingresar_Factura_Cancelada_Admon '" & num_fact2 & "','" & fecha & "'," & saldo & "," & com_admon
    Adodc2.Refresh
    
End Sub

Public Sub eliminar_factura_pendiente()
    
    cn.Execute "Eliminar_Factura_Pendiente_Admon '" & num_fact2 & "'"
    Adodc1.Refresh
    
End Sub

Public Sub Asignar_Valores()
    
    Adodc1.RecordSource = "Select * from Factura_Pendiente_Admon where Fecha_Factura >= '" & Text2.Text & "' And Fecha_Factura <= '" & Text3.Text & "' And Num_Factura = '" & num_fact2 & "'"
    Adodc1.Refresh
    
    num_fact2 = Adodc1.Recordset("Num_Factura")
    fecha = CStr(Text1.Text)
    saldo = CDbl(Adodc1.Recordset("Saldo_Pendiente"))
    com_admon = CDbl(Adodc1.Recordset("Comision_Admon"))
    
End Sub

Private Sub Command5_Click()
    Unload Me
    Factura.Show
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


Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    
    DTPicker1 = Format(Date, "Short Date")
    DTPicker2 = Format(Date, "Short Date")
    DTPicker3 = Format(Date, "Short Date")
    
    Obtener_Fecha
    
    Me.Adodc1.ConnectionString = cn
    Adodc1.RecordSource = "Select * from Factura_Pendiente_Admon"
    Adodc1.Refresh
    
    Me.Adodc2.ConnectionString = cn
    Adodc2.RecordSource = "Select * from Factura_Pendiente_Admon"
    Adodc2.Refresh
    
    Antes
    
    Nombrar_Campo1
    Nombrar_Campo2
    
    Tamaño1
    Tamaño2
    
    With Factura_Admon
        .Height = 9645
        .Width = 14385
    End With
    
    DTPicker4 = Format(Date, "Short date")
    
    Capturar_Fecha
    
End Sub

Public Sub inicializar_tablas()
    MSFlexGrid1.Clear
    MSFlexGrid2.Clear
    
    Nombrar_Campo1
    Nombrar_Campo2
    
    Tamaño1
    Tamaño2
    
    factura_pendiente
    factura_cancelada
End Sub

Public Sub Nombrar_Campo1()
    With MSFlexGrid1
        .TextMatrix(0, 0) = "No. Factura"
        .TextMatrix(0, 1) = "Fecha Factura"
        .TextMatrix(0, 2) = "Saldo Pendiente"
        .TextMatrix(0, 3) = "Comision Admon"
    End With
End Sub

Public Sub Nombrar_Campo2()
    With MSFlexGrid2
        .TextMatrix(0, 0) = "No. Factura"
        .TextMatrix(0, 1) = "Fecha Cancelación"
        .TextMatrix(0, 2) = "Saldo Cancelado"
        .TextMatrix(0, 3) = "Comision Admon"
    End With
End Sub

Public Sub Tamaño1()
    With MSFlexGrid1
        .ColWidth(0) = 1500
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
    End With
End Sub

Public Sub Tamaño2()
    With MSFlexGrid2
        .ColWidth(0) = 1500
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
    End With
End Sub

Public Sub Antes()
    DTPicker1.Enabled = True
    
    Frame6.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = False
    Command4.Enabled = False
End Sub

Public Sub Capturar_Fecha()
    
    año = CStr(DTPicker4.Year)
    
    If DTPicker4.Month < 10 Then
        mes = "0" & CStr(DTPicker4.Month)
    Else
        mes = CStr(DTPicker4.Month)
    End If
    
    If DTPicker4.Day < 10 Then
        dia = "0" & CStr(DTPicker4.Day)
    Else
        dia = CStr(DTPicker4.Day)
    End If
    
    Text7.Text = año & "-" & mes & "-" & dia
    
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
    
    Text1.Text = año & "-" & mes & "-" & dia
    
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
    
    Text2.Text = año & "-" & mes & "-" & dia
    
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
    
    Text3.Text = año & "-" & mes & "-" & dia
    
End Sub

Public Sub Total_Saldo_Pendiente()
    Dim i As Integer
    Dim saldo1 As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid1.Rows - 1
            saldo1 = saldo1 + CDbl(MSFlexGrid1.TextMatrix(i, 2))
        Next i
        
        Text5.Text = CStr(saldo1)
    
tonto:
    If Err.Number <> 0 Then
        Text5.Text = "0.00"
    End If
End Sub

Public Sub Total_Saldo_Cancelado()
    Dim i As Integer
    Dim saldo1 As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid2.Rows - 1
            saldo1 = saldo1 + CDbl(MSFlexGrid2.TextMatrix(i, 2))
        Next i
        
        Text4.Text = CStr(saldo1)
    
tonto:
    If Err.Number <> 0 Then
        Text4.Text = "0.00"
    End If
End Sub

Public Sub Total_Comision_Admon_Pendiente()
    Dim i As Integer
    Dim comision As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid1.Rows - 1
            comision = comision + CDbl(MSFlexGrid1.TextMatrix(i, 3))
        Next i
        
        Text6.Text = CStr(comision)
    
tonto:
    If Err.Number <> 0 Then
        Text6.Text = "0.00"
    End If
End Sub

Public Sub Total_Comision_Admon_Cancelado()
    Dim i As Integer
    Dim comision As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid2.Rows - 1
            comision = comision + CDbl(MSFlexGrid2.TextMatrix(i, 3))
        Next i
        
        Text9.Text = CStr(comision)
    
tonto:
    If Err.Number <> 0 Then
        Text9.Text = "0.00"
    End If
End Sub

Public Sub Totales()
    Total_Saldo_Pendiente
    Total_Saldo_Cancelado
    Total_Comision_Admon_Pendiente
    Total_Comision_Admon_Cancelado
End Sub
