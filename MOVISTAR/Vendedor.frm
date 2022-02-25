VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Vendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendedor"
   ClientHeight    =   6255
   ClientLeft      =   4005
   ClientTop       =   1335
   ClientWidth     =   10320
   Icon            =   "Vendedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10320
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   3960
      Picture         =   "Vendedor.frx":0442
      ScaleHeight     =   1095
      ScaleWidth      =   6375
      TabIndex        =   32
      Top             =   0
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   -3840
      Picture         =   "Vendedor.frx":1349
      ScaleHeight     =   1095
      ScaleWidth      =   9495
      TabIndex        =   30
      Top             =   0
      Width           =   9495
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Left            =   9480
         ScaleHeight     =   1095
         ScaleWidth      =   15
         TabIndex        =   31
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedores"
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
         TabIndex        =   33
         Top             =   600
         Width           =   8775
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   4920
      Top             =   7800
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
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4320
      TabIndex        =   29
      Top             =   8280
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8760
      TabIndex        =   28
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20054017
      CurrentDate     =   39131
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Top             =   8280
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1320
      Top             =   7800
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton CmdEliminar 
      Height          =   615
      Left            =   9000
      Picture         =   "Vendedor.frx":1D4F6
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Eliminar"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton CmdModificar 
      Height          =   615
      Left            =   9000
      Picture         =   "Vendedor.frx":1D938
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Modificar"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdAgregar 
      Height          =   615
      Left            =   9000
      Picture         =   "Vendedor.frx":1DD7A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Agregar"
      Top             =   1560
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "Vendedor.frx":1E1BC
      Top             =   7800
   End
   Begin VB.CommandButton CmdUltimo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Picture         =   "Vendedor.frx":5AEDD
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Ultimo"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton CmdSiguiente 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Picture         =   "Vendedor.frx":5B31F
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Siguiente"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton CmdAnterior 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Picture         =   "Vendedor.frx":5B761
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Anterior"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton CmdPrimero 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Picture         =   "Vendedor.frx":5BBA3
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Primero"
      Top             =   4800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos personales del vendedor"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   8415
      Begin VB.TextBox Text6 
         DataSource      =   "Adodc1"
         Height          =   855
         Left            =   5880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   5880
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "F"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "M"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "Vendedor.frx":5BFE5
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "Vendedor.frx":5C059
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "Vendedor.frx":5C0C1
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "Vendedor.frx":5C129
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "Vendedor.frx":5C19D
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "Vendedor.frx":5C211
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "Vendedor.frx":5C289
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "Vendedor.frx":5C301
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "Vendedor.frx":5C379
         TabIndex        =   19
         Top             =   2040
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Vendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim año As String
Dim mes As String
Dim dia As String

Private Sub CmdAgregar_Click()

    On Error GoTo error1
        If MaskEdBox1.Text = "" Then
            MsgBox "Es necesario que introduzca la identificación", vbCritical, "Error"
        Else
            cn.Execute "Ingresar_Vendedor '" & MaskEdBox1.Text & "','" & Text7.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Combo1.Text & "','" & Text6.Text & "'"
            Adodc1.Refresh
                            
            MsgBox "Los datos del vendedor se han registrado sastifactoriamente", vbInformation, "Registrado sastifactoriamente"
            Unload Me
        End If
error1:
       If Err.Number <> 0 Then
            MsgBox "Error al registrar vendedor", vbCritical, "Error"
       End If
        
End Sub

Private Sub CmdAnterior_Click()
    If Adodc1.Recordset.RecordCount = 0 Then
        ' No hacer nada
    Else
        Adodc1.Recordset.MovePrevious
        
        If Adodc1.Recordset.BOF Then
            Adodc1.Recordset.MoveFirst
        End If
    End If
    Obtener_Sexo
End Sub

Private Sub CmdEliminar_Click()

    On Error GoTo error
        If MaskEdBox1.Text = "" Then
            MsgBox "Es necesario que introduzca la identificación", vbCritical, "Error"
        Else
            cn.Execute "Eliminar_Vendedor '" & MaskEdBox1.Text & "'"
            Adodc1.Refresh
            MsgBox "Los datos del vendedor se han eliminado sastifactoriamente", vbInformation, "Eliminado sastifactoriamente"
            Unload Me
        End If
error:
        If Err.Number <> 0 Then
            MsgBox "Error al eliminar vendedor", vbCritical, "Error"
        End If

End Sub

Private Sub CmdModificar_Click()

    On Error GoTo error
        If MaskEdBox1.Text = "" Then
            MsgBox "Es necesario que introduzca la identificación", vbCritical, "Error"
        Else
            cn.Execute "Modificar_Vendedor '" & MaskEdBox1.Text & "','" & Text7.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Combo1.Text & "','" & Text6.Text & "'"
            Adodc1.Refresh
            MsgBox "Los datos del vendedor se han modificado sastifactoriamente", vbInformation, "Modificado sastifactoriamente"
            Unload Me
        End If
error:
        If Err.Number <> 0 Then
            MsgBox "Error al modificar vendedor", vbCritical, "Error"
        End If

End Sub

Private Sub CmdPrimero_Click()
    If Adodc1.Recordset.RecordCount = 0 Then
        ' No hacer nada
    Else
        Adodc1.Recordset.MoveFirst
    End If
    
    Obtener_Sexo
End Sub

Private Sub CmdSiguiente_Click()
    If Adodc1.Recordset.RecordCount = 0 Then
        ' No hacer nada
    Else
        Adodc1.Recordset.MoveNext
        
        If Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveFirst
        End If
    End If
    Obtener_Sexo
End Sub

Private Sub CmdUltimo_Click()
    If Adodc1.Recordset.RecordCount = 0 Then
        ' No hacer nada
    Else
        Adodc1.Recordset.MoveLast
    End If
    Obtener_Sexo
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Skin1.ApplySkin hWnd
    
    Obtener_Fecha
    
    Me.Adodc1.ConnectionString = cn
    Adodc1.RecordSource = "Select * from Vendedor"
    Adodc1.Refresh
    
    With Vendedor
        .Height = 6765
        .Width = 10290
    End With
    
    With MaskEdBox1
        .Mask = "###-######-####?"
        .PromptChar = "_"
    End With
    
    With Combo1
        .AddItem "Soltero(a)"
        .AddItem "Casado(a)"
        .ListIndex = 0
    End With
    
    Obtener_Sexo
    
    If Text7.Text = "" Then
        Text7.Text = "Masculino"
        Obtener_Sexo
    End If
    
    With Vendedor
        .Height = 6720
        .Width = 10320
    End With
    
End Sub

Public Sub Obtener_Sexo()
    If Text7.Text = "Masculino" Then
        Text7.Text = "Masculino"
        Option1.Value = True
    End If
    
    If Text7.Text = "Femenino" Then
        Text7.Text = "Femenino"
        Option2.Value = True
    End If
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub

Private Sub Option1_Click()
    Text7.Text = "Masculino"
End Sub
    
Private Sub Option2_Click()
    Text7.Text = "Femenino"
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 13 Then
    KeyAscii = 0
End If

    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
    
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 13 Then
    KeyAscii = 0
End If

    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 13 Then
    KeyAscii = 0
End If

    If KeyAscii = 13 Then
        Text5.SetFocus
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 13 Then
    KeyAscii = 0
End If

    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ -+*/%#$\^<>,.;:_()¿?¡!&@{}[]ºª|123456789", Chr(KeyAscii)) = 0 And KeyAscii > 13 Then
    KeyAscii = 0
End If

    If KeyAscii = 13 Then
        On Error GoTo error
            CmdAgregar.SetFocus
error:
            If Err.Number <> 0 Then
                CmdModificar.SetFocus
            End If
    End If
        
End Sub
