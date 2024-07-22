VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmMantenedorServiciosServiteca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicios de Serviteca"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmMantenedorServiciosServiteca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   6615
      Begin VB.TextBox txtComision 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1485
         TabIndex        =   10
         Top             =   1665
         Width           =   1455
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1470
         TabIndex        =   9
         Top             =   1305
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dtcConceptos 
         Bindings        =   "frmMantenedorServiciosServiteca.frx":179A
         Height          =   315
         Left            =   1470
         TabIndex        =   8
         Top             =   945
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datConceptos 
         Height          =   330
         Left            =   2070
         Top             =   945
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
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
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1485
         MaxLength       =   25
         TabIndex        =   0
         Top             =   270
         Width           =   2595
      End
      Begin VB.CheckBox chkVigencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Activo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   4560
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Comisi�n :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Valor :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   1305
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripci�n:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C�digo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   675
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
            Object.ToolTipText     =   "Cerrar (Ctrl+Q)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":17B5
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":18C7
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":19D9
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":1AEB
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":1BFD
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":1D0F
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":1E21
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":1F33
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":2045
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":2157
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":2269
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":237B
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":248D
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":259F
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":26B1
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":27C3
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":28D5
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":2D27
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":3179
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":328B
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":33E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":3543
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":369F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":37FB
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":42C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":471B
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":487F
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":4CDB
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":4E37
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":6143
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":66DF
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":683B
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":6997
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":6CEB
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":703F
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":7393
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":76E7
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":7A3B
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":7F7F
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":8091
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":83E5
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":8739
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":884D
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":8BA1
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":8CB3
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorServiciosServiteca.frx":9007
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantenedorServiciosServiteca"
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
Dim mstrD_P As String
Const mcNombreTabla = "Srvt_Servicios"
Const mcCampoCodigo = "Id_Servicio"
Const mcCampoNombre = "Descripcion"
Private Sub Form_Load()

    mblnSW = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
gintProcedencia = 0
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
        
        LlenaConceptos
        
        If Not Atributos("Glbl", "Tllr_10_0160", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If

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
        Case 17 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
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
        mstrSql = mstrSql & "valor,factor_comision,id_concepto_servicio,usr_id, usr_fecha) "
        mstrSql = mstrSql & " values ('" & Trim(txtCodigo) & "', '" & Trim(txtNombre) & "', '" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSql = mstrSql & CDbl(txtValor) & "," & CDbl(txtComision) & ", '" & dtcConceptos.BoundText & "', '" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "' )"
    Else
        mstrSql = "UPDATE " & mcNombreTabla & " SET " & mcCampoNombre & "='" & Trim(txtNombre) & "', vigencia='" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSql = mstrSql & " usr_id='" & gstrUsuario & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "' , "
        mstrSql = mstrSql & " Valor = " & CDbl(txtValor) & " , Factor_Comision = " & CDbl(txtComision) & ", id_concepto_servicio = '" & dtcConceptos.BoundText & "'"
        mstrSql = mstrSql & " where " & mcCampoCodigo & "='" & Trim(txtCodigo) & "'"
    End If
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
    
    If gintProcedencia = 1 Then
        'frmMantenedorServiciosServiteca.LlenaConceptos
        'frmMantenedorCompa�iaSeguro.Conceptos_Compa�ia frmMantenedorCompa�iaSeguro.txtCodigo
        Unload Me
    End If
End Sub
Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("� Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
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
'    Set FormVol1 = New APFORM1.APFORM
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
    With rptMantenedor
        .ReportFileName = gstrPathReporte & "\APSERVICIOSERVITECA.RPT"
        .Formulas(0) = "Titulo='Listado Servicios de Serviteca'"
        .Formulas(1) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(2) = "Ruc='" & gstrIdEmpresa & "'"
        .Formulas(3) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(4) = "Usuario='" & gstrUsuario & "'"
        .Formulas(5) = "Marcamodulo='ElisaTaller'"
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Connect = cnnAux.ConnectionString
        .Action = True
    End With

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
    Set adoPrincipal = New ADODB.Recordset
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
        txtValor.Text = IIf(Not IsNull(!Valor), !Valor, 0)
        txtComision.Text = IIf(Not IsNull(!factor_comision), !factor_comision, 0)
        dtcConceptos.BoundText = !Id_Concepto_Servicio
    End With
End Sub
Private Sub LimpiaCampos()
    txtCodigo.Text = ""
    chkVigencia.Value = vbUnchecked
    txtNombre.Text = ""
    txtValor.Text = ""
    txtComision.Text = ""
End Sub
Private Sub ValoresporDefecto()
    With adoPrincipal
        chkVigencia.Value = vbChecked
    End With
End Sub
Private Function Validacion() As Boolean
    Validacion = True
    If txtCodigo = "" Then
        MsgBox "El c�digo debe contener un valor...", vbInformation, "Advertencia"
        txtCodigo.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtNombre = "" Then
        MsgBox "La descripci�n debe contener un valor...", vbInformation, "Advertencia"
        txtNombre.SetFocus
        Validacion = False
        Exit Function
    End If
  
    If txtValor = "" Then
        MsgBox "Valor del servicio Incorrecto...", vbInformation, "Advertencia"
        txtValor.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If txtComision = "" Then
        MsgBox "El valor de la Comisi�n es Incorrecto...", vbInformation, "Advertencia"
        txtComision.SetFocus
        Validacion = False
        Exit Function
    End If
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim AdoTemp As New ADODB.Recordset
        mstrSql = "select " & mcCampoCodigo & ", " & mcCampoNombre & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtCodigo & "'"
        If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoTemp.BOF And Not AdoTemp.EOF Then
                MsgBox "Este c�digo ya esta registrado con la descripci�n " & Chr(13) & "[" & IIf(IsNull(AdoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", AdoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                Validacion = False
                txtCodigo.SetFocus
            End If
        End If
        Conexion.CloseHost AdoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMantenedorServiciosServiteca = Nothing
    gstrBusca = txtCodigo.Text
End Sub
Private Sub RevizaAtributos()
'    mblnAccesoCrear = rsUsuarios!OPC_AUXILIAR_CREAR
'    mblnAccesoEditar = rsUsuarios!OPC_AUXILIAR_EDITAR
'    mblnAccesoBorrar = rsUsuarios!OPC_AUXILIAR_BORRAR
'    mblnAccesoImprimir = rsUsuarios!OPC_AUXILIAR_IMPRIMIR

    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True

End Sub

Sub LlenaConceptos()
    dtcConceptos.Enabled = True
    mstrSql = "Select Id_concepto_Servicio as CODIGO, Descripcion as Nombre from Srvt_Concepto_Servicio where VIGENCIA = 'S' order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datConceptos
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                '.Recordset.MoveFirst
                dtcConceptos.ListField = "Nombre"
                dtcConceptos.BoundColumn = "Codigo"
                dtcConceptos.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

