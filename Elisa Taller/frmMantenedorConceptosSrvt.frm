VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmMantenedorConceptosSrvt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos Servicio Serviteca"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmMantenedorConceptosSrvt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   6615
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
         Left            =   1470
         MaxLength       =   25
         TabIndex        =   0
         Top             =   240
         Width           =   2595
      End
      Begin VB.CheckBox chkVigencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Activo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5280
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   4680
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
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
         Caption         =   "Código:"
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
            Picture         =   "frmMantenedorConceptosSrvt.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":18AC
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":19BE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":1AD0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":1BE2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":1CF4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":1E06
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":1F18
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":202A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":213C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":224E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":2360
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":2472
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":2584
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":2696
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":27A8
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":28BA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":2D0C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":315E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":3270
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":33CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":3528
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":3684
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":37E0
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":42AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":4700
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":4864
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":4CC0
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":4E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":6128
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":66C4
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":6820
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":697C
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":6CD0
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":7024
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":7378
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":76CC
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":7A20
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":7F64
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":8076
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":83CA
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":871E
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":8832
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":8B86
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":8C98
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorConceptosSrvt.frx":8FEC
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantenedorConceptosSrvt"
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
Const mcNombreTabla = "Srvt_Concepto_Servicio"
Const mcCampoCodigo = "Id_Concepto_Servicio"
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
        If Not Atributos("Glbl", "Tllr_10_0150", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
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
        mstrSql = mstrSql & "usr_id, usr_fecha) "
        mstrSql = mstrSql & " values ('" & Trim(txtCodigo) & "', '" & Trim(txtNombre) & "', '" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSql = mstrSql & " '" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "' )"
    Else
        mstrSql = "UPDATE " & mcNombreTabla & " SET " & mcCampoNombre & "='" & Trim(txtNombre) & "', vigencia='" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSql = mstrSql & " usr_id='" & gstrUsuario & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "' "
        'mstrSql = mstrSql & " D_P = '" & mstrD_P & "' , Orden = " & updOrden.Value & " "
        mstrSql = mstrSql & " where " & mcCampoCodigo & "='" & Trim(txtCodigo) & "'"
    End If
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
    
    If gintProcedencia = 1 Then
  '      frmMantenedorCompañiaSeguro.Fill_Conceptos
  '      frmMantenedorCompañiaSeguro.Conceptos_Compañia frmMantenedorCompañiaSeguro.txtCodigo
        Unload Me
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
        .ReportFileName = gstrPathReporte & "\APCONCEPTOSERVITECA.RPT"
        .Formulas(0) = "Titulo='Listado Conceptos de Serviteca'"
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
    End With
End Sub
Private Sub LimpiaCampos()
    txtCodigo.Text = ""
    chkVigencia.Value = vbUnchecked
    txtNombre.Text = ""
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
        Dim AdoTemp As New ADODB.Recordset
        mstrSql = "select " & mcCampoCodigo & ", " & mcCampoNombre & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtCodigo & "'"
        If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoTemp.BOF And Not AdoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(AdoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", AdoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                Validacion = False
                txtCodigo.SetFocus
            End If
        End If
        Conexion.CloseHost AdoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMantenedorConceptosSrvt = Nothing
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
