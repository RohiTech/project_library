VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Cambio_Moneda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Moneda"
   ClientHeight    =   4350
   ClientLeft      =   675
   ClientTop       =   705
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6840
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   6000
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
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton CmdAgregar 
      Height          =   615
      Left            =   5640
      Picture         =   "Cambio_Moneda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Agregar"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text3 
      DataField       =   "Cambio_Tasa"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      Picture         =   "Cambio_Moneda.frx":0442
      ScaleHeight     =   2055
      ScaleWidth      =   1815
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   6720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   77463553
      CurrentDate     =   39135
   End
   Begin VB.CommandButton CmdModificar 
      Height          =   615
      Left            =   5640
      Picture         =   "Cambio_Moneda.frx":214E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualizar"
      Top             =   3360
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5280
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   2400
      OleObjectBlob   =   "Cambio_Moneda.frx":2590
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "Fecha_Cambio"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   2040
      TabIndex        =   2
      Top             =   6600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      Picture         =   "Cambio_Moneda.frx":2606
      ScaleHeight     =   1515
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cambio de Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   840
         Width           =   5415
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Cambio_Moneda.frx":44C2
      Top             =   6000
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   2400
      OleObjectBlob   =   "Cambio_Moneda.frx":411E3
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "Cambio_Moneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fecha_actual As String
Dim tasa As String
Dim indice As Integer

Dim año As String
Dim mes As String
Dim dia As String

Private Sub CmdAgregar_Click()
    Adodc1.Recordset.AddNew
    
    Text3.Enabled = True
    
    Text3.SetFocus
    
    Obtener_Fecha
    
    CmdAgregar.Enabled = False
End Sub

Private Sub CmdModificar_Click()
    Adodc1.Recordset.Update
    'On Error GoTo error
        'tasa = InputBox("Actualice Cambio:")
        'cn.Execute "Ingresar_Cambio '" & CStr(Text1.Text) & "'," & CDbl(tasa)
        'CDbl(Text3.Text)
        'Adodc1.Recordset.AddNew
        'Adodc1.Recordset.Update
        'Adodc1.Refresh
        'Busqueda_Cambio
        'Busqueda_Estado
        'estupidez
        Unload Me
        MsgBox "Cambio efectuado sastifactoriamente", vbInformation, "CLARO"
'error:
        'If Err.Number <> 0 Then
            'MsgBox "Error al actualizar cambio", vbCritical, "Error"
        'End If
    
End Sub

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    
    DTPicker1 = Format(Date, "short date")
    
    Obtener_Fecha
    
    Me.Adodc1.ConnectionString = cn
    Adodc1.RecordSource = "Select * from Cambio where Fecha_Cambio = '" & fecha_actual & "'"
    Adodc1.Refresh
    
    Me.Adodc2.ConnectionString = cn
    Adodc2.RecordSource = "Select * from Estado_Moneda where Id_Estado = 1"
    Adodc2.Refresh
    
    Dim entero As Integer
    
    'For entero = 1 To MSFlexGrid1 - 1
        'MSFlexGrid1.ColWidth(entero) = 2000
    'Next entero
    
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.ColWidth(1) = 2000
    MSFlexGrid1.ColWidth(2) = 2000
    
    Poner_Campo
    
    With Cambio_Moneda
        .Height = 4980
        .Width = 6930
    End With
    
    Busqueda_Cambio
    Busqueda_Estado
    
    Text3.Enabled = False
    
End Sub

Public Sub Poner_Campo()
    With MSFlexGrid1
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Fecha del cambio"
        .TextMatrix(0, 2) = "Tasa"
    End With
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
    
    fecha_actual = año & "-" & mes & "-" & dia
    
End Sub

Public Sub Busqueda_Cambio()
    Dim x As Integer

    If Adodc1.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc1.Recordset.MoveFirst
    
            MSFlexGrid1.Rows = CInt(Adodc1.Recordset.RecordCount) + 1
            x = 1
    
            Do While x < Adodc1.Recordset.RecordCount + 1
                MSFlexGrid1.TextMatrix(x, 0) = Adodc1.Recordset.Fields("Id_Cambio")
                MSFlexGrid1.TextMatrix(x, 1) = Adodc1.Recordset.Fields("Fecha_Cambio")
                MSFlexGrid1.TextMatrix(x, 2) = Adodc1.Recordset.Fields("Cambio_Tasa")
                
                If Adodc1.Recordset.Bookmark <> Adodc1.Recordset.RecordCount Then
                    Adodc1.Recordset.MoveNext
                End If
                
                x = x + 1
            Loop
    End If
End Sub

Public Sub Busqueda_Estado()
    Dim x As Integer

    If Adodc2.Recordset.RecordCount = 0 Then
        'MsgBox "No Existe Ningún Registro"
    Else
            Adodc2.Recordset.MoveFirst
    
            x = 1
    
            Do While x < Adodc2.Recordset.RecordCount + 1
                
                Text2.Text = CStr(Adodc2.Recordset.Fields("Tasa_Actual"))
                
                If Adodc2.Recordset.Bookmark <> Adodc2.Recordset.RecordCount Then
                    Adodc2.Recordset.MoveNext
                End If
                
                x = x + 1
            Loop
    End If
End Sub

Public Sub estupidez()
    Text2.Enabled = True
    Text2.SetFocus
    Text2.Enabled = False
End Sub

