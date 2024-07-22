VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMantenedorConcesionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Concesionarios"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmMantenedorConcesionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5280
      Left            =   60
      TabIndex        =   3
      Top             =   375
      Width           =   6615
      Begin MSDataListLib.DataCombo dtcComuna 
         Bindings        =   "frmMantenedorConcesionario.frx":0442
         Height          =   315
         Left            =   1035
         TabIndex        =   16
         Top             =   1620
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcCiudad 
         Bindings        =   "frmMantenedorConcesionario.frx":045A
         Height          =   315
         Left            =   4095
         TabIndex        =   15
         Top             =   1275
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.TextBox txtDireccion 
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   13
         Top             =   915
         Width           =   5475
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   4125
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2010
         Width           =   2355
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2010
         Width           =   2340
      End
      Begin MSComctlLib.ListView lvwMarcas 
         Height          =   2535
         Left            =   60
         TabIndex        =   7
         Top             =   2655
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Codigo"
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Descripcion"
            Text            =   "Descripción"
            Object.Width           =   9701
         EndProperty
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   2
         Top             =   555
         Width           =   4560
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1050
         MaxLength       =   25
         TabIndex        =   0
         Top             =   195
         Width           =   2595
      End
      Begin VB.CheckBox chkVigencia 
         Alignment       =   1  'Right Justify
         Caption         =   "Vigente:"
         Height          =   195
         Left            =   5520
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dtcPais 
         Bindings        =   "frmMantenedorConcesionario.frx":0472
         Height          =   315
         Left            =   1035
         TabIndex        =   19
         Top             =   1275
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datPais 
         Height          =   330
         Left            =   2160
         Top             =   1275
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
      Begin MSAdodcLib.Adodc datComuna 
         Height          =   330
         Left            =   2145
         Top             =   1635
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
      Begin MSAdodcLib.Adodc datCiudad 
         Height          =   330
         Left            =   5235
         Top             =   1275
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   1305
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comuna :"
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   18
         Top             =   1650
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   5
         Left            =   3480
         TabIndex        =   17
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fax :"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   12
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   2055
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marcas Relacionadas"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2445
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   3120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":0488
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":059A
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":06AC
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":07BE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":08D0
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":09E2
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":0AF4
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":0C06
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":0D18
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":0E2A
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":0F3C
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":104E
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":1160
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":1272
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":1384
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":1496
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":15A8
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":19FA
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConcesionario.frx":1E4C
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro (Ctrl+G)"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
            ImageKey        =   "Cancelar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Registro (Ctrl+D)"
            ImageKey        =   "Borrar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro (Ctrl+B)"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
            ImageKey        =   "Primero"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
            ImageKey        =   "Siguiente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro (Ctrl+U)"
            ImageKey        =   "Ultimo"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad:"
      Height          =   195
      Index           =   8
      Left            =   0
      TabIndex        =   21
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "frmMantenedorConcesionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset

Dim mstrSql As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mblnSW As Boolean
Const mcNombreTabla = "Glbl_Concesionarios"
Const mcCampoCodigo = "Id_Concesionario"
Const mcCampoNombre = "Razon_Social"
Sub FillPais()
'    dtcPais.Enabled = True
    mstrSql = "Select Id_PAis as CODIGO, Descripcion as Nombre from Glbl_Pais where VIGENCIA = 'S' order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datPais
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcPais.ListField = "Nombre"
                dtcPais.BoundColumn = "Codigo"
'                If .Recordset.RecordCount < 2 Then
'                    dtcPais.BoundText = .Recordset!Codigo
'                    dtcPais.Enabled = False
'                End If
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub


Sub FillCiudad(strPais As String)
'    dtcCiudad.Enabled = True
    mstrSql = "Select Id_Ciudad as CODIGO, Descripcion as Nombre from Glbl_Ciudad where VIGENCIA = 'S' AND ID_PAIS = '" & strPais & "' order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datCiudad
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcCiudad.ListField = "Nombre"
                dtcCiudad.BoundColumn = "Codigo"
'                If .Recordset.RecordCount < 2 Then
'                    dtcPais.BoundText = .Recordset!CODIGO
'                    dtcCiudad.Enabled = False
'                End If
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Sub FillComuna(strPais As String, strCiudad As String)
'    dtcComuna.Enabled = True
    mstrSql = "Select Id_Comuna as CODIGO, Descripcion as Nombre from Glbl_Comuna where VIGENCIA = 'S' AND ID_PAIS = '" & strPais & "' AND ID_CIUDAD = '" & strCiudad & "' order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datComuna
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcComuna.ListField = "Nombre"
                dtcComuna.BoundColumn = "Codigo"
'                If .Recordset.RecordCount < 2 Then
'                    dtcComuna.BoundText = .Recordset!Codigo
'                    dtcComuna.Enabled = False
'                End If
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub


Sub MarcasConcesionario(strConcesionario As String)
    mstrSql = "SELECT Id_Marca FROM Glbl_Concesionarios_Vs_Marca WHERE Id_Concesionario = '" & strConcesionario & "' "
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveLast: .MoveFirst
                While Not .EOF
                    Set lvwMarcas.SelectedItem = lvwMarcas.FindItem(CStr(!Id_Marca), , , 1)
                    lvwMarcas.SelectedItem.Checked = True
                    .MoveNext
                Wend
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

'Private Sub CHECK_OFF()
'Dim V As Integer
'
'For V = 1 To lvwMarcas.ListItems.Count
'    Set lvwMarcas.SelectedItem = lvwMarcas.ListItems(V)
'    lvwMarcas.SelectedItem.Checked = False
'Next
'End Sub
Sub FillMarcas()
Dim Item As ListItem
    
lvwMarcas.ListItems.Clear
mstrSql = "SELECT Id_Marca, Descripcion FROM Glbl_Marca WHERE Vigencia = 'S' order by descripcion"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst
            While Not .EOF
                Set Item = lvwMarcas.ListItems.Add(, , !Id_Marca)
                Item.SubItems(1) = !Descripcion
                .MoveNext
            Wend
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal

End Sub

Private Sub dtcCiudad_Change()
If dtcCiudad.BoundText <> "" Then
    dtcComuna.Text = ""
    FillComuna dtcPais.BoundText, dtcCiudad.BoundText
End If
End Sub

Private Sub dtcPais_Change()
If dtcPais.BoundText <> "" Then
    dtcCiudad.Text = ""
    dtcComuna.Text = ""
    FillCiudad dtcPais.BoundText
End If
End Sub

Private Sub Form_Load()
    mblnSW = True
End Sub



Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            AgregarRegistro
        Case "Grabar"
            GrabarRegistro
        Case "Cancelar"
            CancelarAgregaRegistro
        Case "Borrar"
            BorrarRegistro
        Case "Buscar"
            BuscarRegistro
        Case "Imprimir"
            ImprimirInforme
        Case "Primero"
            PrimerRegistro
        Case "Anterior"
            RegistroAnterior
        Case "Siguiente"
            RegistroSiguiente
        Case "Ultimo"
            UltimoRegistro
        Case "Renovar"
            Renovar
        Case "Cerrar"
            CerrarSalir
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0030", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        FillMarcas
        FillPais
        If gapAccion = apcrear Then
           AgregarRegistro
           txtCodigo = gstrBusca
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost adoPrincipal
            End If
            txtCodigo.Enabled = False
            Me.SetFocus
        End If
        If gapAccion = apninguno Then
           Renovar
        End If
    End If
    gapAccion = apninguno
    mblnSW = False
    txtNombre.SetFocus
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            CancelarAgregaRegistro
        Case 14 And tlbBarraHerramientas.Buttons.Item("Crear").Enabled
            KeyAscii = 0
            AgregarRegistro
        Case 7 And tlbBarraHerramientas.Buttons.Item("Grabar").Enabled
            KeyAscii = 0
            GrabarRegistro
        Case 4 And tlbBarraHerramientas.Buttons.Item("Borrar").Enabled
            KeyAscii = 0
            BorrarRegistro
        Case 2 And tlbBarraHerramientas.Buttons.Item("Buscar").Enabled
            KeyAscii = 0
            BuscarRegistro
        Case 9 And tlbBarraHerramientas.Buttons.Item("Imprimir").Enabled
            KeyAscii = 0
            ImprimirInforme
        Case 16 And tlbBarraHerramientas.Buttons.Item("Primero").Enabled
            KeyAscii = 0
            PrimerRegistro
        Case 1 And tlbBarraHerramientas.Buttons.Item("Anterior").Enabled
            KeyAscii = 0
            RegistroAnterior
        Case 19 And tlbBarraHerramientas.Buttons.Item("Siguiente").Enabled
            KeyAscii = 0
            RegistroSiguiente
        Case 21 And tlbBarraHerramientas.Buttons.Item("Ultimo").Enabled
            KeyAscii = 0
            UltimoRegistro
        Case 18 And tlbBarraHerramientas.Buttons.Item("Renovar").Enabled
            KeyAscii = 0
            Renovar
        Case 3 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    txtCodigo.SetFocus
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost adoPrincipal
    txtNombre.SetFocus
End Sub
Private Sub GrabarRegistro()
    If Not Validacion() Then
        Exit Sub
    End If

    If Me.Tag = "Crear" Then
        mstrSql = "INSERT INTO " & mcNombreTabla & " (" & mcCampoCodigo & ", " & mcCampoNombre & ", vigencia, "
        mstrSql = mstrSql & "usr_id, usr_fecha, Id_Comuna, Id_Ciudad, Id_Pais, Direccion, Telefono, Fax) "
        mstrSql = mstrSql & "values ('" & Trim(txtCodigo) & "', '" & Trim(txtNombre) & "', '" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSql = mstrSql & "'" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "', '" & dtcComuna.BoundText & "' , '" & dtcCiudad.BoundText & "' , '" & dtcPais.BoundText & "' , '" & txtDireccion & "', '" & txtTelefono & "' , '" & txtFax & "' )"
    Else                                                                                                                                'Id_Comuna,                                                                                                      Id_Ciudad, Id_Pais, Direccion, Telefono, Fax) "
        mstrSql = "UPDATE " & mcNombreTabla & " SET " & mcCampoNombre & "='" & Trim(txtNombre) & "', vigencia='" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSql = mstrSql & "usr_id='" & gstrUsuario & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "', Id_Comuna = '" & dtcComuna.BoundText & "' , Id_Ciudad = '" & dtcCiudad.BoundText & "', Id_Pais = '" & dtcPais.BoundText & "', Direccion = '" & txtDireccion & "', Telefono = '" & txtTelefono & "', Fax = '" & txtFax & "' "
        mstrSql = mstrSql & " where " & mcCampoCodigo & "='" & Trim(txtCodigo) & "'"
    End If
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
    
     GuardaMarcas Trim(txtCodigo)
    
End Sub

Sub GuardaMarcas(strConcesionario As String)
Dim X As Integer

If lvwMarcas.ListItems.Count > 0 Then
    For X = 1 To lvwMarcas.ListItems.Count
        Set lvwMarcas.SelectedItem = lvwMarcas.ListItems(X)
        If lvwMarcas.SelectedItem.Checked = True Then
            If VerificaMarcaConcesionario(strConcesionario, lvwMarcas.SelectedItem) = False Then
                mstrSql = "INSERT INTO Glbl_Concesionarios_Vs_Marca ( ID_Marca, ID_Concesionario )"
                mstrSql = mstrSql & " VALUES('" & lvwMarcas.SelectedItem & "' , '" & strConcesionario & "' ) "
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
            End If
        Else
            If VerificaMarcaConcesionario(strConcesionario, lvwMarcas.SelectedItem) = True Then
                mstrSql = "DELETE FROM Glbl_Concesionarios_Vs_Marca WHERE ID_Marca='" & lvwMarcas.SelectedItem & "' AND ID_Concesionario ='" & strConcesionario & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
            End If
        End If
    Next '///////////////AQUI GRABA LAS NUEVAS Y LAS QUE ESTABAN
End If
End Sub

Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        mstrSql = "DELETE FROM " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtCodigo & "'"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo
                    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                            LeerCampos
                        Else
                            mblnTablaVacia = True
                            LimpiaCampos
                        End If
                    End If
                End If
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
End Sub
Private Sub BuscarRegistro()
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption)
    If gstrBusca <> "" Then
        mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
    Me.SetFocus
End Sub
Private Sub ImprimirInforme()
   ' FormVol1.ImprimirRegistros Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption, gstrPathReporte, "APCARROC.RPT", gstrUSUARIO, gstrCodigoEmpresa
End Sub
Private Sub PrimerRegistro()
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroAnterior()
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo & " DESC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroSiguiente()

    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub UltimoRegistro()
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo & " DESC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub Renovar()
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        VerificaTablaVacia
        ActivaBotones
        If Not mblnTablaVacia Then
            PrimerRegistro
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
    txtCodigo.Enabled = False
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
        .Item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoEditar, True, False))
        .Item("Cancelar").Enabled = False
        .Item("Borrar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoBorrar, True, False))
        .Item("Buscar").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Imprimir").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoImprimir, True, False))
        .Item("Primero").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Anterior").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Siguiente").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Ultimo").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Renovar").Enabled = True
        .Item("Cerrar").Enabled = True
    End With
End Sub
Private Sub DesactivaBotones()
    txtCodigo.Enabled = True
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = False
        .Item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
        .Item("Cancelar").Enabled = True
        .Item("Borrar").Enabled = False
        .Item("Buscar").Enabled = False
        .Item("Imprimir").Enabled = False
        .Item("Primero").Enabled = False
        .Item("Anterior").Enabled = False
        .Item("Siguiente").Enabled = False
        .Item("Ultimo").Enabled = False
        .Item("Renovar").Enabled = False
        .Item("Cerrar").Enabled = True
    End With
End Sub
Private Sub VerificaTablaVacia()
    If (Not adoPrincipal.BOF And Not adoPrincipal.EOF) And adoPrincipal.RecordCount > 0 Then
        mblnTablaVacia = False
    Else
        mblnTablaVacia = True
        LimpiaCampos
        MsgBox "La tabla no contiene registros...", vbInformation, "Advertencia"
'        Me.lvwMarcas.Enabled = False
    End If
End Sub
Private Sub LeerCampos()

    If mblnTablaVacia Then
        LimpiaCampos
        Exit Sub
    End If

    With adoPrincipal
        txtCodigo.Text = ValorNulo(.Fields(mcCampoCodigo))
        If IsNull(!vigencia) Then
            chkVigencia.Value = vbUnchecked
        Else
            If !vigencia = "S" Then
                chkVigencia.Value = vbChecked
            Else
                chkVigencia.Value = vbUnchecked
            End If
        End If
        txtNombre.Text = ValorNulo(.Fields(mcCampoNombre))
        txtDireccion.Text = ValorNulo(.Fields("Direccion"))
        dtcPais.BoundText = !Id_Pais
        dtcCiudad.BoundText = !Id_Ciudad
        dtcComuna.BoundText = !ID_Comuna
        txtTelefono.Text = ValorNulo(.Fields("Telefono"))
        txtFax.Text = ValorNulo(.Fields("Fax"))
        SetCheckOff lvwMarcas
        MarcasConcesionario .Fields(mcCampoCodigo)
        
    End With
End Sub
Private Sub LimpiaCampos()
    txtCodigo.Text = ""
    chkVigencia.Value = vbUnchecked
    txtNombre.Text = ""
    txtDireccion.Text = ""
    txtFax.Text = ""
    txtTelefono.Text = ""
    If datPais.Recordset.RecordCount > 1 Then
        dtcPais.BoundText = ""
    Else
        dtcPais.Enabled = False
    End If
    dtcCiudad.BoundText = ""
    dtcComuna.BoundText = ""
    
End Sub
Private Sub ValoresporDefecto()
    With adoPrincipal
        chkVigencia.Value = vbChecked
    End With
End Sub
Private Function Validacion() As Boolean
    Validacion = True
    If txtCodigo = "" Then
        MsgBox "El código debe contener un valor...", vbInformation, "Advertencia"
        txtCodigo.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtNombre = "" Then
        MsgBox "La descripción debe contener un valor...", vbInformation, "Advertencia"
        txtNombre.SetFocus
        Validacion = False
        Exit Function
    End If
  
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As ADODB.Recordset
        mstrSql = "select " & mcCampoCodigo & ", " & mcCampoNombre & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtCodigo & "'"
        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(adoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", adoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                Validacion = False
                txtCodigo.SetFocus
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMantenedorCompañiaSeguro = Nothing
    gstrBusca = txtCodigo.Text
End Sub
Private Sub RevizaAtributos()
    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True
End Sub
