VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Libro_de_Ventas 
   Caption         =   "Libro de Ventas"
   ClientHeight    =   11010
   ClientLeft      =   8595
   ClientTop       =   990
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Consultar Factura"
      Height          =   375
      Left            =   6720
      TabIndex        =   61
      Top             =   4920
      Width           =   1935
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   5400
      TabIndex        =   60
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   255
      Left            =   4080
      OleObjectBlob   =   "Libro_de_Ventas.frx":0000
      TabIndex        =   59
      Top             =   5040
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   120
      Top             =   10440
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adodc7"
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
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":006E
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Salir"
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":04B0
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Cambiar de vendedor(a)"
      Top             =   6600
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4200
      OleObjectBlob   =   "Libro_de_Ventas.frx":08F2
      Top             =   8400
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   49
      Top             =   8520
      Width           =   2535
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   53
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   52
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3D613
         TabIndex        =   50
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3D689
         TabIndex        =   51
         Top             =   960
         Width           =   1455
      End
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
      Height          =   1575
      Left            =   240
      TabIndex        =   44
      Top             =   6720
      Width           =   2535
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   48
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   47
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3D701
         TabIndex        =   45
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3D777
         TabIndex        =   46
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ventas al Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4080
      TabIndex        =   42
      Top             =   5520
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2175
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         FixedCols       =   0
         AllowUserResizing=   1
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
      Height          =   2775
      Left            =   4080
      TabIndex        =   40
      Top             =   1920
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2175
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3836
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
      Top             =   10680
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
      TabIndex        =   38
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   120
      TabIndex        =   37
      Top             =   10680
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
      Left            =   4080
      TabIndex        =   30
      Top             =   9000
      Width           =   10815
      Begin VB.CommandButton Command4 
         Height          =   615
         Left            =   9600
         Picture         =   "Libro_de_Ventas.frx":3D7EF
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Reporte de las ventas"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   8400
         Picture         =   "Libro_de_Ventas.frx":3DC2D
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Reporte del detalle de las ventas"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   7200
         Picture         =   "Libro_de_Ventas.frx":3E06F
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Buscar"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   5520
         TabIndex        =   34
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   39128
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3240
         TabIndex        =   33
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   39128
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   4680
         OleObjectBlob   =   "Libro_de_Ventas.frx":3E4B1
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "Libro_de_Ventas.frx":3E513
         TabIndex        =   32
         Top             =   480
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3E56F
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   13440
      Top             =   10560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   12120
      Top             =   10560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Adodc4"
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
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   9840
      TabIndex        =   28
      Top             =   10560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   10920
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   8280
      Top             =   10560
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
      Top             =   10560
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
      Top             =   10560
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
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13440
      TabIndex        =   26
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   10680
      TabIndex        =   25
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13440
      TabIndex        =   22
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   10680
      TabIndex        =   21
      Top             =   4920
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   9720
      OleObjectBlob   =   "Libro_de_Ventas.frx":3E5D5
      TabIndex        =   19
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton CmdAgregar 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":3E645
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Agregar"
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton CmdModificar 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":3EA87
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Modificar"
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdEliminar 
      Height          =   615
      Left            =   2880
      Picture         =   "Libro_de_Ventas.frx":3EEC9
      Style           =   1  'Graphical
      TabIndex        =   16
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
      TabIndex        =   9
      Top             =   4080
      Width           =   2535
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   1680
         TabIndex        =   39
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3F30B
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3F38F
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3F403
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":3F473
         TabIndex        =   15
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      Picture         =   "Libro_de_Ventas.frx":3F4E9
      ScaleHeight     =   1575
      ScaleWidth      =   6735
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Libro de Ventas"
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   8775
      End
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
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":4210D
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   39123
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Libro_de_Ventas.frx":4217D
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   960
      Picture         =   "Libro_de_Ventas.frx":421E7
      ScaleHeight     =   1575
      ScaleWidth      =   14295
      TabIndex        =   0
      Top             =   0
      Width           =   14295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   12480
      OleObjectBlob   =   "Libro_de_Ventas.frx":44E0B
      TabIndex        =   20
      Top             =   5040
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   9720
      OleObjectBlob   =   "Libro_de_Ventas.frx":44E7B
      TabIndex        =   23
      Top             =   8640
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   255
      Left            =   12480
      OleObjectBlob   =   "Libro_de_Ventas.frx":44EEB
      TabIndex        =   24
      Top             =   8640
      Width           =   735
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
Dim question As Boolean

Dim fec As String

Dim f As String

Public Sub Total_Comision()
    Calcular_Total_Comision_Vendedor_Contado
    Calcular_Total_Comision_Arqueador_Contado
    Calcular_Total_Comision_Vendedor_Credito
    Calcular_Total_Comision_Arqueador_Credito
End Sub

Private Sub CmdAgregar_Click()
    On Error GoTo error
        
        If Combo3.Text = "" Or Text2.Text = "" Or Text14.Text = "" Or Combo6.Text = "" Then
            MsgBox "Por favor complete el formulario detalle venta", vbCritical, "Error"
        Else
            num = InputBox("1. Contado " & vbCrLf & "2. Crédito")
            Select Case num
                Case 1
                    cn.Execute "Ingresar_Detalle_Venta_Contado '" & Combo3.Text & "'," & CInt(Text2.Text) & "," & CInt(Text14.Text) & "," & CInt(Combo6) & ",'" & Text8.Text & "','" & Combo5.Text & "'"
                    Adodc2.RecordSource = "Select * from Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc2.Refresh
                    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc3.Refresh
                    detalle_contado
                    busqueda_total_contado
                            
                    estupidez
                    Combo1.ListIndex = 0
                    DTPicker2 = DTPicker1
                    DTPicker3 = DTPicker1
                    Obtener_Fecha_Inicial
                    Obtener_Fecha_Final
                    Command1_Click
                Case 2
                    
                    If MaskEdBox1.Text = "_____" Then
                        MsgBox "Por Favor Digite El Número De Factura", vbCritical, "Error"
                    Else
                        RealizarConsulta
                        estupidez
                        Combo1.ListIndex = 0
                        DTPicker2 = DTPicker1
                        DTPicker3 = DTPicker1
                        Obtener_Fecha_Inicial
                        Obtener_Fecha_Final
                        Command1_Click
                    End If
                    
            End Select
            
        End If
error:
        If Err.Number <> 0 Then
           MsgBox "Error al agregar detalle venta", vbCritical, "Error"
           Text2.SetFocus
        End If
End Sub

Public Sub RealizarConsulta()
    Dim ing As Boolean

    ' Obtener la fecha de la factura para comparar si es igual a la fecha actual
    On Error GoTo error
        ' La factura existe. No Ingresar. Autosumar si la fecha actual es igual a la fecha de la factura
        Adodc7.RecordSource = "Select * from Factura_Pendiente where Num_Factura = '" & MaskEdBox1.Text & "' And Num_Cedula = '" & Combo5.Text & "'"
        Adodc7.Refresh
        
        f = CStr(Adodc7.Recordset("Fecha_Factura")) '28/04/2007'
        
        If f = fec Then
            Ingresar_Venta_Credito
        Else
            MsgBox "La factura ya existe", vbCritical, "Error"
        End If
error:
    If Err.Number <> 0 Then
        ' La factura no existe. Debe Ingresar.
        
        On Error GoTo error2
            cn.Execute "Ingresar_Factura_Pendiente '" & MaskEdBox1.Text & "','" & Text8.Text & "','" & Combo5.Text & "'"
            Adodc7.Refresh
            ing = True
error2:
            If Err.Number <> 0 Then
                MsgBox "Error al ingresar factura", vbCritical, "Error"
                ing = False
            End If
                    
            If ing = True Then
                Ingresar_Venta_Credito
            End If
        
    End If
End Sub

Public Sub Ingresar_Venta_Credito()
    On Error GoTo error4
            If MaskEdBox1.Text = "_____" Then
                MsgBox "Digite el número de factura", vbExclamation, "Aviso"
            Else
                ubicar
                cn.Execute "Ingresar_Detalle_Venta_Credito '" & Combo3.Text & "'," & CInt(Text2.Text) & "," & CInt(Text14.Text) & "," & CInt(Combo6) & ",'" & Text8.Text & "','" & Combo5.Text & "','" & MaskEdBox1.Text & "'"
                Adodc4.RecordSource = "Select * from Venta_Credito where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                Adodc4.Refresh
                Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                Adodc5.Refresh
                detalle_credito
                busqueda_total_credito
            End If
error4:
            If Err.Number <> 0 Then
                MsgBox "Error al ingresar venta", vbCritical, "Error"
            End If
End Sub

Public Sub ubicar()
    Combo6.ListIndex = 0 ' Todas las ventas al crédito es del 4 %
End Sub

Public Sub ingresar_factura()

    Dim num_factura As String
    
    On Error GoTo error
    
        num_factura = InputBox("Dígite el No. de Factura [#####] : ")
        
        'Frame2.Caption = "Venta al Crédito Factura No. " & num_factura
        
        cn.Execute "Ingresar_Factura_Pendiente '" & num_factura & "','" & Text8.Text & "','" & Combo5.Text & "'"
        Adodc7.Refresh
        
        contador = contador + 1
        booleano = True
error:
    If Err.Number <> 0 Then
        MsgBox "El número de factura ya existe", vbCritical, "Error"
        contador = 0
        booleano = False
    End If
End Sub

Private Sub CmdEliminar_Click()
    On Error GoTo error
            num = InputBox("1. Contado " & vbCrLf & "2. Crédito")
            Select Case num
                Case 1
                    num2 = InputBox("Dígite el Id Venta para eliminar el detalle de venta al contado")
                    cn.Execute "Eliminar_Detalle_Venta_Contado " & num2 & ",'" & Text8.Text & "','" & Combo5.Text & "'"
                    Adodc2.RecordSource = "Select * from Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc2.Refresh
                    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc3.Refresh
                    detalle_contado
                    busqueda_total_contado
                        
                    estupidez
                    inicializar_tablas
                    Command1_Click
                Case 2
                    If MaskEdBox1.Text = "_____" Then
                        MsgBox "Por favor digite el numero de factura", vbCritical, "Error"
                    Else
                        ubicar
                        num2 = InputBox("Dígite el Id Venta para eliminar el detalle de venta al credito")
                        cn.Execute "Eliminar_Detalle_Venta_Credito " & num2 & ",'" & Text8.Text & "','" & Combo5.Text & "','" & MaskEdBox1.Text & "'"
                        Adodc4.RecordSource = "Select * from Venta_Credito where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                        Adodc4.Refresh
                        Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                        Adodc5.Refresh
                        detalle_credito
                        busqueda_total_credito
                        
                        estupidez
                        inicializar_tablas
                        Command1_Click
            
                    End If
            End Select
            
error:
        If Err.Number <> 0 Then
            MsgBox "Error al eliminar detalle venta", vbCritical, "Error"
            MsgBox "Para eliminar asegurese de que la fecha principal y el numero de factura que seleccionó sea la correcta", vbExclamation, "Aviso Importante"
        End If
End Sub

Private Sub CmdModificar_Click()
    On Error GoTo error
        If Combo3.Text = "" Or Text2.Text = "" Or Text14.Text = "" Or Combo6.Text = "" Then
            MsgBox "Por favor complete el formulario detalle venta", vbCritical, "Error"
        Else
            num = InputBox("1. Contado " & vbCrLf & "2. Crédito")
            Select Case num
                Case 1
                    num2 = InputBox("Dígite el Id Venta para modificar el detalle de venta al contado")
                    cn.Execute "Modificar_Detalle_Venta_Contado " & num2 & ",'" & Combo3.Text & "'," & CInt(Text2.Text) & "," & CInt(Text14.Text) & "," & CInt(Combo6) & ",'" & Text8.Text & "','" & Combo5.Text & "'"
                    Adodc2.RecordSource = "Select * from Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc2.Refresh
                    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                    Adodc3.Refresh
                    detalle_contado
                    busqueda_total_contado
                Case 2
                    If MaskEdBox1.Text = "_____" Then
                        MsgBox "Por favor digite el numero de factura", vbCritical, "Error"
                    Else
                        ubicar
                        num2 = InputBox("Dígite el Id Venta para modificar el detalle de venta al credito")
                        cn.Execute "Modificar_Detalle_Venta_Credito " & num2 & ",'" & Combo3.Text & "'," & CInt(Text2.Text) & "," & CInt(Text14.Text) & "," & CInt(Combo6) & ",'" & Text8.Text & "','" & Combo5.Text & "','" & MaskEdBox1.Text & "'"
                        Adodc4.RecordSource = "Select * from Venta_Credito where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                        Adodc4.Refresh
                        Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Fecha_Venta = '" & Text8.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
                        Adodc5.Refresh
                        detalle_credito
                        busqueda_total_credito
                    End If
            End Select
            
            estupidez
        End If
error:
        If Err.Number <> 0 Then
            MsgBox "Error al modificar detalle venta", vbCritical, "Error"
            MsgBox "Para modificar asegurese de que la fecha principal y el numero de factura que seleccionó sea la correcta", vbExclamation, "Aviso Importante"
            Text2.SetFocus
        End If
End Sub

Private Sub Combo2_Click()
    
    Combo5.ListIndex = Combo2.ListIndex
    
End Sub

Private Sub Command1_Click()
    
    Frame2.Caption = "Venta al Crédito"
    
    MSFlexGrid1.Clear
    MSFlexGrid2.Clear
    
    Poner_Campo_Contado
    Poner_Campo_Credito
    
    Dim seleccion As String
    
    seleccion = CStr(Combo1.Text)
    
    Select Case seleccion
        Case "Ambos"
            Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
            Adodc3.Refresh
            Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
            Adodc5.Refresh
        Case "4"
            Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "' And Descuento = 4"
            Adodc3.Refresh
            Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "' And Descuento = 4"
            Adodc5.Refresh
        Case "5"
            Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "' And Descuento = 5"
            Adodc3.Refresh
            Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "' And Descuento = 5"
            Adodc5.Refresh
        Case "6"
            Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "' And Descuento = 6"
            Adodc3.Refresh
            Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "' And Descuento = 6"
            Adodc5.Refresh
    End Select
    
    detalle_contado
    detalle_credito
    
    Calcular_Total_Contado
    Calcular_Total_Credito
    Total_Comision
    estupidez
    
End Sub

Public Sub Calcular_Total_Comision_Vendedor_Contado()
    
    Dim i As Integer
    Dim comision As Double
    
    On Error GoTo tonto
        For i = 1 To MSFlexGrid1.Rows - 1
            comision = comision + CDbl(MSFlexGrid1.TextMatrix(i, 8))
        Next i
        
        Text1.Text = CStr(comision)
tonto:
    If Err.Number <> 0 Then
        Text1.Text = "0.00"
    End If
    
End Sub

Public Sub Calcular_Total_Comision_Arqueador_Contado()
    
    Dim i As Integer
    Dim comision As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid1.Rows - 1
            comision = comision + CDbl(MSFlexGrid1.TextMatrix(i, 9))
        Next i
        
        Text3.Text = CStr(comision)
    
tonto:
    If Err.Number <> 0 Then
        Text3.Text = "0.00"
    End If
    
End Sub

Public Sub Calcular_Total_Comision_Vendedor_Credito()
    
    Dim i As Integer
    Dim comision As Double
    
    On Error GoTo tonto
        For i = 1 To MSFlexGrid2.Rows - 1
            comision = comision + CDbl(MSFlexGrid2.TextMatrix(i, 8))
        Next i
        
        Text9.Text = CStr(comision)
    
tonto:
    If Err.Number <> 0 Then
        Text9.Text = "0.00"
    End If
    
End Sub

Public Sub Calcular_Total_Comision_Arqueador_Credito()
    
    Dim i As Integer
    Dim comision As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid2.Rows - 1
            comision = comision + CDbl(MSFlexGrid2.TextMatrix(i, 9))
        Next i
        
        Text10.Text = CStr(comision)
    
tonto:
    If Err.Number <> 0 Then
        Text10.Text = "0.00"
    End If
    
End Sub

Public Sub Calcular_Total_Contado()
    Dim i As Integer
    Dim saldo_dol As Double
    Dim saldo_cor As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid1.Rows - 1
            saldo_dol = saldo_dol + CDbl(MSFlexGrid1.TextMatrix(i, 6))
            saldo_cor = saldo_cor + CDbl(MSFlexGrid1.TextMatrix(i, 11))
        Next i
        
        Text4.Text = CStr(saldo_dol)
        Text5.Text = CStr(saldo_cor)
tonto:
    If Err.Number <> 0 Then
        Text4.Text = "0.00"
        Text5.Text = "0.00"
    End If
End Sub

Public Sub Calcular_Total_Credito()
    Dim i As Integer
    Dim saldo_dol As Double
    Dim saldo_cor As Double
    
    On Error GoTo tonto
    
        For i = 1 To MSFlexGrid2.Rows - 1
            saldo_dol = saldo_dol + CDbl(MSFlexGrid2.TextMatrix(i, 6))
            saldo_cor = saldo_cor + CDbl(MSFlexGrid2.TextMatrix(i, 12))
        Next i
        
        Text6.Text = CStr(saldo_dol)
        Text7.Text = CStr(saldo_cor)
tonto:
    If Err.Number <> 0 Then
        Text6.Text = "0.00"
        Text7.Text = "0.00"
    End If
End Sub

Private Sub Command2_Click()
        
        metodo_contado
        
        metodo_credito
        
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
        
        Dim num_fact As String
        
        Adodc7.RecordSource = "Select * from Factura_Pendiente where Fecha_Factura = '" & Text8.Text & "' And Num_Cedula = '" & Combo5.Text & "'"
        Adodc7.Refresh
            
        Command2.Enabled = False
        
End Sub

Public Sub metodo_contado()
   
    Dim x As Integer
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

Public Sub metodo_credito()
   
    Dim x As Integer
    Dim id_venta As Integer
        
    Adodc4.RecordSource = "Select * from Venta_Credito where Fecha_Venta = '" & Text8.Text & "' And Num_Cedula = '" & Combo5.Text & "'"
    Adodc4.Refresh
    
    On Error GoTo error
        id_venta = CInt(Adodc4.Recordset("Id_Venta"))
        consulta_credito
error:
    If Err.Number <> 0 Then
        Ingresar_Credito
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

Public Sub Ingresar_Credito()
    cn.Execute "Ingresar_Venta_Credito '" & Text8.Text & "','" & Combo5.Text & "'"
    Adodc4.Refresh
End Sub

Public Sub consulta_credito()
    Adodc4.RecordSource = "Select * from Venta_Credito where Num_Cedula = '" & Combo5.Text & "' and Fecha_Venta = '" & Text8.Text & "'"
    Adodc4.Refresh
End Sub

Private Sub Command3_Click()

Dim question As Integer

On Error GoTo error
question = InputBox("1. Contado" & vbCrLf & "2. Crédito")

Select Case question
    Case 1
        str1 = "Select * from Detalle_Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
        
        rst1.Open str1, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Venta_Contado.DataSource = rst1
        rpt_Venta_Contado.Show
    Case 2
        str2 = "Select * from Detalle_Venta_Credito where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
        
        rst2.Open str2, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Venta_Credito.DataSource = rst2
        rpt_Venta_Credito.Show
End Select

error:
    If Err.Number <> 0 Then
    End If
    
End Sub

Private Sub Command4_Click()

Dim question As Integer

On Error GoTo error

question = InputBox("1. Contado" & vbCrLf & "2. Crédito")

Select Case question
    Case 1
        str3 = "Select * from Venta_Contado where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
        
        rst3.Open str3, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Total_Venta_Contado.DataSource = rst3
        rpt_Total_Venta_Contado.Show
    Case 2
        str4 = "Select * from Venta_Credito where Fecha_Venta >= '" & Text12.Text & "' And Fecha_Venta <= '" & Text13.Text & "' and Num_Cedula = '" & Combo5.Text & "'"
        
        rst4.Open str4, cn, adOpenDynamic, adLockOptimistic
        Set rpt_Total_Venta_Credito.DataSource = rst4
        rpt_Total_Venta_Credito.Show
End Select

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

Private Sub Command7_Click()
    MSFlexGrid1.Clear
    MSFlexGrid2.Clear
    
    Poner_Campo_Contado
    Poner_Campo_Credito
    
    Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Num_Factura = '" & MaskEdBox1.Text & "' and Num_Cedula = '" & Combo5.Text & "'" ' and Fecha_Venta >= '" & Text12.Text & "' and Fecha_Venta <= '" & Text13.Text & "'"
    Adodc5.Refresh
    detalle_credito
    
    Calcular_Total_Contado
    Calcular_Total_Credito
    Total_Comision
    estupidez
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
    
    Me.Adodc4.ConnectionString = cn
    Adodc4.RecordSource = "Select * from Venta_Credito"
    Adodc4.Refresh
    
    Me.Adodc5.ConnectionString = cn
    Adodc5.RecordSource = "Select * from Detalle_Venta_Credito"
    Adodc5.Refresh
    
    Me.Adodc6.ConnectionString = cn
    Adodc6.RecordSource = "Select * from Estado_Moneda where Id_Estado = 1"
    Adodc6.Refresh
    
    Me.Adodc7.ConnectionString = cn
    Adodc7.RecordSource = "Select * from Factura_Pendiente"
    Adodc7.Refresh
        
    busqueda_vendedor
    
    With Combo3
        .AddItem "1"
        .AddItem "1.5"
        .AddItem "3"
        .AddItem "6"
        .AddItem "12"
        .AddItem "20"
        .ListIndex = 0
    End With
    
    With Combo6
        .AddItem "4" '5
        .AddItem "5" '6
        .ListIndex = 0
    End With
    
    With Combo1
        .AddItem "Ambos"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .ListIndex = 0
    End With
    
    Deshabilitar
    
    Poner_Campo_Contado
    Poner_Campo_Credito
    
    Obtener_Fecha
    Obtener_Fecha_Inicial
    Obtener_Fecha_Final
    
    Tamaño_Campo_Contado
    Tamaño_Campo_Credito
    
    Frame5.Caption = "Total Comisión Contado"
    Frame7.Caption = "Total Comisión Crédito"
    
    Combo5.ListIndex = Combo2.ListIndex
    
    moneda = Adodc6.Recordset("Tasa_Actual")
    
    contador = 0
    booleano = False
    question = False
    
    MaskEdBox1.Mask = "#####"
    MaskEdBox1.PromptChar = "_"
    
End Sub

Public Sub inicializar_tablas()
    MSFlexGrid1.Clear
    MSFlexGrid2.Clear
    
    Poner_Campo_Contado
    Poner_Campo_Credito
    
    Tamaño_Campo_Contado
    Tamaño_Campo_Credito
    
    detalle_contado
    busqueda_total_contado
    
    detalle_credito
    busqueda_total_credito
    
End Sub

Public Sub Actualizar_Adodc()
    
    Adodc1.RecordSource = "Select * from Vendedor"
    Adodc1.Refresh
    
    
    Adodc2.RecordSource = "Select * from Venta_Contado"
    Adodc2.Refresh
    
    
    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado"
    Adodc3.Refresh
    
    
    Adodc4.RecordSource = "Select * from Venta_Credito"
    Adodc4.Refresh
    
    
    Adodc5.RecordSource = "Select * from Detalle_Venta_Credito"
    Adodc5.Refresh
    
    
    Adodc6.RecordSource = "Select * from Estado_Moneda where Id_Estado = 1"
    Adodc6.Refresh
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
        .ColWidth(10) = 1000 ' Fecha_Venta
        .ColWidth(11) = 880 ' SubTotal Cordobas
    End With
End Sub

Public Sub Tamaño_Campo_Credito()
    With MSFlexGrid2
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
        .ColWidth(10) = 1000 ' Fecha_Venta
        .ColWidth(11) = 1000 ' Num_Factura
        .ColWidth(12) = 880 ' SubTotal Cordobas
    End With
End Sub

Public Sub Poner_Campo_Contado()
    With MSFlexGrid1
        .TextMatrix(0, 0) = "Id Venta"
        .TextMatrix(0, 1) = "Tipo$"
        .TextMatrix(0, 2) = "Disponible"
        .TextMatrix(0, 3) = "Cantidad"
        .TextMatrix(0, 4) = "Descuento$"
        .TextMatrix(0, 5) = "Precio$"
        .TextMatrix(0, 6) = "Sub-total$"
        .TextMatrix(0, 7) = "Devolución"
        .TextMatrix(0, 8) = "Com. Vendedor C$"
        .TextMatrix(0, 9) = "Com. Arqueador C$"
        .TextMatrix(0, 10) = "Fecha Venta"
        .TextMatrix(0, 11) = "SubTotal C$"
    End With
End Sub

Public Sub Poner_Campo_Credito()
    With MSFlexGrid2
        .TextMatrix(0, 0) = "Id Venta"
        .TextMatrix(0, 1) = "Tipo$"
        .TextMatrix(0, 2) = "Disponible"
        .TextMatrix(0, 3) = "Cantidad"
        .TextMatrix(0, 4) = "Descuento$"
        .TextMatrix(0, 5) = "Precio$"
        .TextMatrix(0, 6) = "Sub-total$"
        .TextMatrix(0, 7) = "Devolución"
        .TextMatrix(0, 8) = "Com. Vendedor C$"
        .TextMatrix(0, 9) = "Com. Arqueador C$"
        .TextMatrix(0, 10) = "Fecha Venta"
        .TextMatrix(0, 11) = "No. Factura"
        .TextMatrix(0, 12) = "SubTotal C$"
    End With
End Sub

Public Sub Deshabilitar()
    Frame4.Enabled = False
    Frame6.Enabled = False
    CmdAgregar.Enabled = False
    CmdModificar.Enabled = False
    CmdEliminar.Enabled = False
    MaskEdBox1.Enabled = False
    Command7.Enabled = False
End Sub

Public Sub Habilitar()
    Frame4.Enabled = True
    Frame6.Enabled = True
    CmdAgregar.Enabled = True
    CmdModificar.Enabled = True
    CmdEliminar.Enabled = True
    MaskEdBox1.Enabled = True
    Command7.Enabled = True
End Sub

Public Sub busqueda_vendedor()
    
    Dim x As Integer

    If Adodc1.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc1.Recordset.MoveFirst
    
            x = 1
    
            Do While x < Adodc1.Recordset.RecordCount + 1
                Combo2.AddItem CStr(Adodc1.Recordset.Fields("I_Nombre")) & " " & CStr(Adodc1.Recordset.Fields("II_Nombre")) & " " & CStr(Adodc1.Recordset.Fields("I_Apellido")) & " " & CStr(Adodc1.Recordset.Fields("II_Apellido"))
                Combo5.AddItem CStr(Adodc1.Recordset.Fields("Num_Cedula"))
                
                If Adodc1.Recordset.Bookmark <> Adodc1.Recordset.RecordCount Then
                    Adodc1.Recordset.MoveNext
                End If
                
                x = x + 1
            Loop
            Combo2.ListIndex = 0
            Combo5.ListIndex = 0
    End If
    
End Sub

Public Sub detalle_contado()
    
    Dim x As Integer

    If Adodc3.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc3.Recordset.MoveFirst
    
            MSFlexGrid1.Rows = CInt(Adodc3.Recordset.RecordCount) + 1
            x = 1
    
            Do While x < Adodc3.Recordset.RecordCount + 1
                MSFlexGrid1.TextMatrix(x, 0) = Adodc3.Recordset.Fields("Id_Venta")
                MSFlexGrid1.TextMatrix(x, 1) = Adodc3.Recordset.Fields("Tipo")
                MSFlexGrid1.TextMatrix(x, 2) = Adodc3.Recordset.Fields("Disponible")
                MSFlexGrid1.TextMatrix(x, 3) = Adodc3.Recordset.Fields("Cantidad")
                MSFlexGrid1.TextMatrix(x, 4) = Adodc3.Recordset.Fields("Descuento")
                MSFlexGrid1.TextMatrix(x, 5) = Adodc3.Recordset.Fields("Precio")
                MSFlexGrid1.TextMatrix(x, 6) = Adodc3.Recordset.Fields("SubTotal")
                MSFlexGrid1.TextMatrix(x, 7) = Adodc3.Recordset.Fields("Devolucion")
                MSFlexGrid1.TextMatrix(x, 8) = Adodc3.Recordset.Fields("Comision_Vendedor")
                MSFlexGrid1.TextMatrix(x, 9) = Adodc3.Recordset.Fields("Comision_Arqueador")
                MSFlexGrid1.TextMatrix(x, 10) = Adodc3.Recordset.Fields("Fecha_Venta")
                MSFlexGrid1.TextMatrix(x, 11) = Adodc3.Recordset.Fields("SubTotal_Cordobas")
                
                If Adodc3.Recordset.Bookmark <> Adodc3.Recordset.RecordCount Then
                    Adodc3.Recordset.MoveNext
                End If
                
                x = x + 1
            Loop
    End If

End Sub

Public Sub detalle_credito()
    
    Dim x As Integer

    If Adodc5.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc5.Recordset.MoveFirst
    
            MSFlexGrid2.Rows = CInt(Adodc5.Recordset.RecordCount) + 1
            x = 1
    
            Do While x < Adodc5.Recordset.RecordCount + 1
                MSFlexGrid2.TextMatrix(x, 0) = Adodc5.Recordset.Fields("Id_Venta")
                MSFlexGrid2.TextMatrix(x, 1) = Adodc5.Recordset.Fields("Tipo")
                MSFlexGrid2.TextMatrix(x, 2) = Adodc5.Recordset.Fields("Disponible")
                MSFlexGrid2.TextMatrix(x, 3) = Adodc5.Recordset.Fields("Cantidad")
                MSFlexGrid2.TextMatrix(x, 4) = Adodc5.Recordset.Fields("Descuento")
                MSFlexGrid2.TextMatrix(x, 5) = Adodc5.Recordset.Fields("Precio")
                MSFlexGrid2.TextMatrix(x, 6) = Adodc5.Recordset.Fields("SubTotal")
                MSFlexGrid2.TextMatrix(x, 7) = Adodc5.Recordset.Fields("Devolucion")
                MSFlexGrid2.TextMatrix(x, 8) = Adodc5.Recordset.Fields("Comision_Vendedor")
                MSFlexGrid2.TextMatrix(x, 9) = Adodc5.Recordset.Fields("Comision_Arqueador")
                MSFlexGrid2.TextMatrix(x, 10) = Adodc5.Recordset.Fields("Fecha_Venta")
                MSFlexGrid2.TextMatrix(x, 11) = Adodc5.Recordset.Fields("Num_Factura")
                MSFlexGrid2.TextMatrix(x, 12) = Adodc5.Recordset.Fields("SubTotal_Cordobas")
                
                If Adodc5.Recordset.Bookmark <> Adodc5.Recordset.RecordCount Then
                    Adodc5.Recordset.MoveNext
                End If
                
                x = x + 1
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
    fec = dia & "/" & mes & "/" & año
    
End Sub

Public Sub busqueda_total_contado()
    
On Error GoTo error
    
    Dim x As Integer
    
        If Adodc2.Recordset.RecordCount = 0 Then
            'MsgBox "No Existe Ningún Registro"
        Else
                Adodc2.Recordset.MoveFirst
        
                x = 1
        
                Do While x < Adodc2.Recordset.RecordCount + 1
                    
                    Text1.Text = CStr(Adodc2.Recordset.Fields("Total_Comision_Vendedor"))
                    Text3.Text = CStr(Adodc2.Recordset.Fields("Total_Comision_Arqueador"))
                    
                    Text4.Text = CStr(Adodc2.Recordset.Fields("Total_$"))
                    Text5.Text = CStr(Adodc2.Recordset.Fields("Total_C$"))
                      
                    If Adodc2.Recordset.Bookmark <> Adodc2.Recordset.RecordCount Then
                        Adodc2.Recordset.MoveNext
                    End If
                    
                    x = x + 1
                Loop
        End If
error:
    If Err.Number <> 0 Then
        Text1.Text = "0"
        Text3.Text = "0"
        Text4.Text = "0"
        Text5.Text = "0"
    End If
    
End Sub

Public Sub busqueda_total_credito()
    
On Error GoTo error
    
    Dim x As Integer

    If Adodc4.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc4.Recordset.MoveFirst
    
            x = 1
    
            Do While x < Adodc4.Recordset.RecordCount + 1
                
                Text9.Text = CStr(Adodc4.Recordset.Fields("Total_Comision_Vendedor"))
                Text10.Text = CStr(Adodc4.Recordset.Fields("Total_Comision_Arqueador"))
                
                Text6.Text = CStr(Adodc4.Recordset.Fields("Total_$"))
                Text7.Text = CStr(Adodc4.Recordset.Fields("Total_C$"))
                
                If Adodc4.Recordset.Bookmark <> Adodc4.Recordset.RecordCount Then
                    Adodc4.Recordset.MoveNext
                End If
                
                x = x + 1
            Loop
    End If
    
error:
    If Err.Number <> 0 Then
        Text9.Text = "0"
        Text10.Text = "0"
        Text6.Text = "0"
        Text7.Text = "0"
    End If
    
End Sub

Public Sub Detalle()
    
    Obtener_Fecha
    
    Adodc2.RecordSource = "Select * from Venta_Contado where Num_Cedula = '" & Combo5.Text & "' And Fecha_Venta = '" & Text8.Text & "'"
    Adodc2.Refresh
    
    Adodc3.RecordSource = "Select * from Detalle_Venta_Contado where Num_Cedula = '" & Combo5.Text & "' And Fecha_Venta = '" & Text8.Text & "'"
    Adodc3.Refresh
    
    Adodc4.RecordSource = "Select * from Venta_Credito where Num_Cedula = '" & Combo5.Text & "' And Fecha_Venta = '" & Text8.Text & "'"
    Adodc4.Refresh
    
    Adodc5.RecordSource = "Select * from Detalle_Venta_Credito where Num_Cedula = '" & Combo5.Text & "' And Fecha_Venta = '" & Text8.Text & "'"
    Adodc5.Refresh
    
    detalle_contado
    busqueda_total_contado
    detalle_credito
    busqueda_total_credito
    
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
    fec = dia & "/" & mes & "/" & año
    
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
    fec = dia & "/" & mes & "/" & año
    
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
    Text9.Enabled = True
    Text10.Enabled = True
    
    Text1.SetFocus
    Text3.SetFocus
    Text9.SetFocus
    Text10.SetFocus
    
    Text1.Enabled = False
    Text3.Enabled = False
    Text9.Enabled = False
    Text10.Enabled = False
    
    Text14.Text = ""
    Text2.Text = ""
    
    Text2.SetFocus
    
End Sub

Public Sub Total_Credito()
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
