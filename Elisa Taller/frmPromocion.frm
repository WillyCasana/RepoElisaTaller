VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmPromocion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Campañas"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmPromocion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   111935489
         CurrentDate     =   44278
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   111935489
         CurrentDate     =   44278
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   6120
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Final:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
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
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
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
            Object.Visible         =   0   'False
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro (Ctrl+B)"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
      Top             =   120
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
            Picture         =   "frmPromocion.frx":038A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":049C
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":05AE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":06C0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":07D2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":08E4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":09F6
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":0B08
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":0C1A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":0D2C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":0E3E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":0F50
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":1062
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":1174
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":1286
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":1398
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":14AA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":18FC
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":1D4E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":1E60
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":1FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":2118
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":2274
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":23D0
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":2E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":32F0
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":3454
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":38B0
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":3A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":4D18
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":52B4
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":5410
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":556C
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":58C0
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":5C14
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":5F68
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":62BC
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":6610
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":6B54
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":6C66
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":6FBA
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":730E
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":7422
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":7776
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":7888
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPromocion.frx":7BDC
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPromocion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoPrincipal As New ADODB.Recordset
Dim mstrSQL As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mblnSW As Boolean
Dim mstrD_P As String

Private Sub Form_Activate()
If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0022", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
'FillPromocion
        If gapAccion = apcrear Then
           AgregarRegistro
           txtCodigo = gstrBusca
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSQL = "select * from Promocion WHERE Id_Promo='" & gstrBusca & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
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
    Screen.MousePointer = vbDefault
End Sub

Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    Me.txtDescripcion.SetFocus
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

Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    mstrSQL = "select TOP 1 * from Promocion WHERE Id_Promo >'" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSQL = "select TOP 1 * from Promocion WHERE Id_Promo <'" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
            If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost AdoPrincipal
    cboMes.SetFocus
End Sub
Private Sub GrabarRegistro()
    If Not validacion() Then
        Exit Sub
    End If

    If Me.Tag = "Crear" Then
        mstrSQL = "INSERT INTO Promocion ( Descripcion,"
        mstrSQL = mstrSQL & "id_usr, fecha_usr , fecha_inicio, fecha_final,Vigencia,Id_Empresa,Id_Sucursal) "
        mstrSQL = mstrSQL & " values ( '" & Trim(Me.txtDescripcion.Text) & "', "
        mstrSQL = mstrSQL & " '" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "','" & Me.dtpInicio.Value & "', '" & Me.dtpFinal.Value & " ', '" & IIf(chkActivo.Value = vbChecked, "S", "N") & "',"
        mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "')"
    Else
        mstrSQL = "UPDATE Promocion SET Descripcion ='" & Trim(txtDescripcion.Text) & "', "
        mstrSQL = mstrSQL & " id_usr='" & gstrUsuario & "', fecha_usr='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "' ,"
        mstrSQL = mstrSQL & " fecha_inicio= '" & Me.dtpInicio.Value & "', "
        mstrSQL = mstrSQL & " fecha_final= '" & Me.dtpFinal.Value & "', "
        mstrSQL = mstrSQL & " Vigencia= '" & IIf(chkActivo.Value = vbChecked, "S", "N") & "',"
        mstrSQL = mstrSQL & " Id_Empresa= '" & gstrIdEmpresa & "',"
        mstrSQL = mstrSQL & " Id_Sucursal= '" & gstrIdSucursal & "'"
        mstrSQL = mstrSQL & " where Id_Promo ='" & Trim(txtCodigo) & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    End If
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
End Sub
Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        mstrSQL = "DELETE FROM Promocion where Id_Promo ='" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
            mstrSQL = "select TOP 1 * from Promocion WHERE Id_Promo >'" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
            If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    mstrSQL = "select TOP 1 * from Promocion WHERE Id_Promo <'" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
                    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                            LeerCampos
                        Else
                            mblnTablaVacia = True
                            LimpiaCampos
                        End If
                    End If
                End If
            End If
        End If
        Conexion.CloseHost AdoPrincipal
    End If
End Sub

Private Sub PrimerRegistro()
    
    mstrSQL = "select TOP 1 * from Promocion Where ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroAnterior()
    
    mstrSQL = "select TOP 1 * from Promocion WHERE Id_Promo <'" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo DESC"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroSiguiente()

    mstrSQL = "select TOP 1 * from Promocion WHERE Id_Promo>'" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub UltimoRegistro()
    mstrSQL = "select TOP 1 * from Promocion Where ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo DESC"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub Renovar()
    Set AdoPrincipal = New ADODB.Recordset
    mstrSQL = "select TOP 1 * from Promocion Where ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Id_Promo"
    
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        VerificaTablaVacia
        ActivaBotones
        If Not mblnTablaVacia Then
            PrimerRegistro
        End If
    End If
    Conexion.CloseHost AdoPrincipal
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
    If (Not AdoPrincipal.BOF And Not AdoPrincipal.EOF) And AdoPrincipal.RecordCount > 0 Then
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
    With AdoPrincipal
        txtCodigo.Text = ValorNulo(!Id_Promo)
        Me.txtDescripcion.Text = ValorNulo(!Descripcion)
        Me.dtpInicio.Value = ValorNulo(!Fecha_Inicio)
        Me.dtpFinal.Value = ValorNulo(!fecha_final)
        If IsNull(!vigencia) Then
        chkActivo.Value = vbUnchecked
        Else
            If !vigencia = "S" Then
                chkActivo.Value = vbChecked
            Else
                chkActivo.Value = vbUnchecked
            End If
        End If
        
    End With
End Sub
Private Sub LimpiaCampos()
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    dtpInicio.Value = BOM(Date)
    dtpFinal.Value = EOM(Date)
    chkActivo.Value = vbChecked
   
End Sub
Private Sub ValoresporDefecto()
    With AdoPrincipal
        
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    dtpInicio.Value = BOM(Date)
    dtpFinal.Value = EOM(Date)
    chkActivo.Value = vbChecked
    End With
End Sub
Private Function validacion() As Boolean
    validacion = True
    If txtDescripcion = "" Then
        MsgBox "La descripcion debe contener un valor...", vbInformation, "Advertencia"
        txtDescripcion.SetFocus
        validacion = False
        Exit Function
    End If
    
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As New ADODB.Recordset
        mstrSQL = "select Id_Promo, Descripcion from Promocion where Id_Promo='" & txtCodigo & "' And ID_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' "
        If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(adoTemp.Fields(Descripcion)), "SIN DESCRIPCION", adoTemp.Fields(Descripcion)) & "]", vbInformation, "Advertencia"
                validacion = False
                txtCodigo.SetFocus
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmPromocion = Nothing
    gstrBusca = txtCodigo.Text
End Sub
