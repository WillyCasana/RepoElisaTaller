VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Gestión Taller  - Elisa Taller"
   ClientHeight    =   3525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10755
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "TallerPro"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   10695
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   10755
      Begin VB.TextBox txtComando 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3225
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   "EMPRESA"
            Object.ToolTipText     =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "SUCURSAL"
            Object.ToolTipText     =   "Sucursal"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Key             =   "USUARIO"
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "MAYÚS"
            Object.ToolTipText     =   "Indicador de Mayúscula"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NÚM"
            Object.ToolTipText     =   "Indicador de Numérico"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "29/06/2024"
            Object.ToolTipText     =   "Fecha del Sistema"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1646
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":234E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3056
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuPopup 
      Caption         =   "MenuPopup"
      Visible         =   0   'False
      Begin VB.Menu popup 
         Caption         =   "Cambiar Item"
         Index           =   1
      End
      Begin VB.Menu popup 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu popup 
         Caption         =   "Descuentos"
         Index           =   3
      End
      Begin VB.Menu popup 
         Caption         =   "Cargo"
         Index           =   4
      End
      Begin VB.Menu popup 
         Caption         =   "Mecánico"
         Index           =   5
      End
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "Configuración"
      Begin VB.Menu itemSistema 
         Caption         =   "Usuarios"
         Index           =   10
      End
      Begin VB.Menu itemSistema 
         Caption         =   "Roles de Usuario"
         Index           =   20
      End
      Begin VB.Menu itemSistema 
         Caption         =   "Cambio de Contraseña"
         Index           =   30
      End
      Begin VB.Menu itemSistema 
         Caption         =   "-"
         Index           =   40
      End
      Begin VB.Menu itemSistema 
         Caption         =   "Opciones del sistema"
         Index           =   50
      End
      Begin VB.Menu itemSistema 
         Caption         =   "-"
         Index           =   60
      End
      Begin VB.Menu itemSistema 
         Caption         =   "Salir"
         Index           =   70
      End
   End
   Begin VB.Menu mnuGlobales 
      Caption         =   "Mantenedores"
      Begin VB.Menu itemGlobales 
         Caption         =   "Marcas"
         Index           =   10
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Modelos"
         Index           =   20
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Clientes y Proveedores"
         Index           =   40
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Compañias de Seguro"
         Index           =   50
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Costo de Mano de Obra"
         Index           =   60
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Conceptos de Inventario"
         Index           =   70
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Conos"
         Index           =   80
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Tipos de OT"
         Index           =   90
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Tipos de Cargo"
         Index           =   100
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "-"
         Index           =   110
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Mecanicos"
         Index           =   120
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Especialidades"
         Index           =   130
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Horas Mecanico"
         Index           =   140
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "-"
         Index           =   150
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Temparios"
         Index           =   160
         Begin VB.Menu itemTemparios 
            Caption         =   "Servicios Generales"
            Index           =   10
         End
         Begin VB.Menu itemTemparios 
            Caption         =   "Actividades Generales"
            Index           =   20
         End
         Begin VB.Menu itemTemparios 
            Caption         =   "Temparios"
            Index           =   30
         End
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "-"
         Index           =   170
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Vehículos de Clientes"
         Index           =   180
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Tipo de Cambio Moneda"
         Index           =   190
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Promociones y/o Campañas"
         Index           =   200
      End
      Begin VB.Menu itemGlobales 
         Caption         =   "Tipo de Trabajo"
         Index           =   210
      End
   End
   Begin VB.Menu mnuOperaciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuReservaHoras 
         Caption         =   "Reserva de Atención"
      End
      Begin VB.Menu mnuRecordatorio 
         Caption         =   "Recordatorio de Servicio"
      End
      Begin VB.Menu ro1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecepcionMecCar 
         Caption         =   "Generar Orden de Trabajo"
      End
      Begin VB.Menu mnuGenOrdenesTrabajo 
         Caption         =   "Liquidación Orden de Trabajo"
      End
      Begin VB.Menu mnuGenerarPresupuestos 
         Caption         =   "Generar Presupuestos"
      End
      Begin VB.Menu ro2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPresupuestoMantenciones 
         Caption         =   "Presupuestos de Mantención"
         Visible         =   0   'False
      End
      Begin VB.Menu ro3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovEmiOrdComTer 
         Caption         =   "Ordenes de Compra a Terceros"
      End
      Begin VB.Menu rayahorasreales 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHorasActividades 
         Caption         =   "Horas Reales por Actividades"
      End
      Begin VB.Menu raya3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAsignacionTurnos 
         Caption         =   "Asignación de Turnos"
      End
      Begin VB.Menu mnuAusenciaMecanicos 
         Caption         =   "Ausencia de Mecánicos"
      End
      Begin VB.Menu raya4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAsignacionRecursos 
         Caption         =   "Asignación de Recursos"
      End
      Begin VB.Menu mnuRelojTareas 
         Caption         =   "Cronómetro de Tareas"
      End
      Begin VB.Menu Raya5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReemplazaVinPlaca 
         Caption         =   "Reemplaza Vin por Placa"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFacturarInternos 
         Caption         =   "Facturar Cargos Internos"
      End
      Begin VB.Menu mnuPrueba 
         Caption         =   "Prueba"
      End
   End
   Begin VB.Menu mnuServiteca 
      Caption         =   "Serviteca"
      Visible         =   0   'False
      Begin VB.Menu mnuOT 
         Caption         =   "Orden de Trabajo"
      End
      Begin VB.Menu mnuOpcServiteca 
         Caption         =   "Opciones de Serviteca"
      End
      Begin VB.Menu mnuConsOTSrvt 
         Caption         =   "Consulta OT Serviteca"
      End
      Begin VB.Menu mnuMaestroComisionesSrvt 
         Caption         =   "Factores Para Comisiones"
      End
      Begin VB.Menu mnuInformeComicionesSrvt 
         Caption         =   "Informe de Comisiones"
      End
      Begin VB.Menu mnuResumenServiteca 
         Caption         =   "Resumen Servicios Serviteca"
      End
   End
   Begin VB.Menu mnuInformes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuCMT 
         Caption         =   "Consulta Multipropósito de Taller"
      End
      Begin VB.Menu mnuEmisionPresupuestos 
         Caption         =   "Consulta de Presupuestos"
         Shortcut        =   ^P
      End
      Begin VB.Menu ri1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfLisOrdTra 
         Caption         =   "Listado de Ordenes de Trabajo"
      End
      Begin VB.Menu mnuInfOrdCom 
         Caption         =   "Reporte de Ordenes de Compra"
      End
      Begin VB.Menu ri2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfMargRep 
         Caption         =   "Margen de Repuestos por OT"
      End
      Begin VB.Menu mnuRepuestosGeneral 
         Caption         =   "Margen de Repuestos General"
      End
      Begin VB.Menu mnuReservaRepuestos 
         Caption         =   "Reserva de Repuestos"
      End
      Begin VB.Menu mnuServiceRate 
         Caption         =   "Service Rate Reserva Repuestos"
      End
      Begin VB.Menu ri3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfResValOT 
         Caption         =   "Resumen Valorizado por O/T"
      End
      Begin VB.Menu mnuInfRentOT 
         Caption         =   "Rentabilidad por O/T"
      End
      Begin VB.Menu mnuInfResDed 
         Caption         =   "Resumen de Deducibles"
      End
      Begin VB.Menu mnuInfResumenProveedores 
         Caption         =   "Resumen de Proveedores"
      End
      Begin VB.Menu mnuRentaCarroceria 
         Caption         =   "Rentabilidad Sección Carrocería"
      End
      Begin VB.Menu mnuAsignadasMecanicos 
         Caption         =   "Tareas Asignadas por Mecanicos"
      End
      Begin VB.Menu MnuHoras 
         Caption         =   "Consulta de Tareas"
      End
      Begin VB.Menu ri4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfProdMec 
         Caption         =   "Productividad por Mecánico"
      End
      Begin VB.Menu ri5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistoricoPatente 
         Caption         =   "Historico de Atenciones Por Placa"
      End
      Begin VB.Menu mnuFactInt 
         Caption         =   "Facturación Interna"
      End
      Begin VB.Menu mnuDaiTaller 
         Caption         =   "Informe Daily Taller"
      End
      Begin VB.Menu mnuReservaAten 
         Caption         =   "Reporte Reserva de Atención"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Visible         =   0   'False
      Begin VB.Menu mnuAyudaTaller 
         Caption         =   "Ayuda Elisa Taller"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
'kjcv 07.04.15
Dim j As Integer
Const maxCargos = 9

Private Sub itemGlobales_Click(Index As Integer)
Screen.MousePointer = vbHourglass

Select Case Index
    Case 10
        Libreria.Marcas Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 20
        Libreria.Modelos Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 30
        '-
    Case 40
        Dim lstrIdCliente As String
        Dim lstrDescripCliente As String
        Libreria.ClienteNuevo Conexion, lstrIdCliente, lstrDescripCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 50
        Libreria.CompañiasDeSeguro Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 60
        gstrBusca = CostoManoObra(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 70
        Libreria.ConceptosInventario Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 80
        Libreria.Conos Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 90
        gstrBusca = TipoGarantias(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 100
        gstrBusca = TipoCargo(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 110
        '-
    Case 120
        gstrBusca = Mecanicos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 130
        Libreria.EspecialidadMecanico Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 140
        gstrBusca = HorasMecanico(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 150
        '-
    Case 160
        '- temparios
    Case 170
        '-
    Case 180
        gstrProcedencia = "Mantenedor"
        gstrBusca = Vehiculos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 190
        Libreria.TipoCambio Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 200
        gstrBusca = Promocion(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
        
    Case 210
        gstrBusca = Trabajos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Select

Screen.MousePointer = vbDefault
End Sub

Private Sub itemSistema_Click(Index As Integer)
Screen.MousePointer = vbHourglass

Select Case Index
    Case 10
        Libreria.Usuarios Conexion, gstrPathReporte, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 20
        Libreria.RolesDeUsuario Conexion, gstrPathReporte, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
    Case 30
        frmCambioContraseña.Show
    Case 40
        '-
    Case 50
        frmMantenedorParametros.Show
    Case 60
        '-
    Case 70
        Unload Me
End Select

Screen.MousePointer = vbDefault

End Sub

Private Sub itemTemparios_Click(Index As Integer)
Screen.MousePointer = vbHourglass

Select Case Index
    Case 10
        gstrBusca = ServiciosGenerales(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 20
        gstrBusca = ActividadesGenerales(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
    Case 30
        gstrBusca = TemparioServicios(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Select

Screen.MousePointer = vbDefault
End Sub

Private Sub MDIForm_Activate()
        '//Rescata parametros para apServer.
        If Not SW Then
           SW = True
           Call LoginConsola
        End If
End Sub

Private Sub MDIForm_Load()
Me.Caption = "ElisaTaller - versión " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'apServer.apLogout
    Dim lintResponde As Integer
    
    lintResponde = MsgBox("Está seguro de abandonar ElisaTaller?", vbYesNo + vbQuestion, "ElisaTaller")
    If lintResponde = vbYes Then
        Conexion.DisconnectHost
        'SERVER.apLogout
        End
    Else
        Cancel = 1
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuActividadesGenerales_Click()
gstrBusca = ActividadesGenerales(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuAsignacionRecursos_Click()
frmAsignacionRecursos.Show
End Sub

Private Sub mnuAsignacionTurnos_Click()
frmAsignacionTurnos.Show
End Sub

Private Sub mnuAsignadasMecanicos_Click()
frmTareasasignadasMecanico.Show
End Sub

Private Sub mnuAusenciaMecanicos_Click()
frmAusenciaMecanicos.Show
End Sub

Private Sub mnuCompañiaSeguro_Click()
gstrProcedencia = "Mantenedor"
gstrBusca = CompañiaSeguro(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuConceptoInventario_Click()
gstrBusca = ConceptosInventario(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuConceptosServiteca_Click()
gstrProcedencia = "Mantenedor"
gstrBusca = ConceptoSrvt(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuConDesPin_Click()
gstrBusca = ConceptosDyP(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuConos_Click()
gstrBusca = TipoConos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuCMT_Click()
Libreria.ConsultaMultipropositoTaller Conexion, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
End Sub

Private Sub mnuConsOTSrvt_Click()
With frmBuscaOTSrvt
    .cmdSeleccionar.Enabled = False
    .Show vbModal
End With
End Sub

Private Sub mnuDaiTaller_Click()
frmInformeDayli.Show
End Sub

Private Sub mnuEmisionPresupuestos_Click()
gstrBusca = Presupuestos(Conexion, gstrIdUsuario, "Tllr", " ", gstrIdEmpresa, gstrPathReporte, "", apninguno)
gstrImpresion = "O"
gstrProcedencia = "Presupuesto"
End Sub

Private Sub mnuEspecialidad_Click()
gstrBusca = Especialidades(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuFactInt_Click()
frmInformeFacturacionInterna.Show
End Sub

Private Sub mnuFacturarInternos_Click()
frmFacturarCargosInternos.Show
End Sub

Private Sub mnuGenerarPresupuestos_Click()
gstrBusca = GeneraPresupuestos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, gstrPathReporte, "", apninguno)
gstrImpresion = "O"
gstrProcedencia = "Presupuestos"
End Sub

Private Sub mnuGenOrdenesTrabajo_Click()
gstrBusca = OrdenesdeTrabajo(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, gstrPathReporte, "", apninguno)
gstrImpresion = "O"
gstrProcedencia = "Movimientos"
End Sub

Private Sub mnuHistoricoPatente_Click()
frmHistoricoPatente.Show
End Sub

Private Sub MnuHoras_Click()
FrmConsultadeHoras.Show
End Sub

Private Sub mnuHorasActividades_Click()
frmIngresaHorasMecanicoActividades.Show
End Sub

Private Sub mnuHorasMecanico_Click()
gstrBusca = HorasMecanico(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuInfDayTllr_Click()
frmInformeDayli.Show
End Sub

Private Sub mnuInfLisOrdTra_Click()
With frmBuscaOT
    .cmdSeleccionar.Enabled = False
    .Show vbModal
End With
End Sub

Private Sub mnuInfMargRep_Click()
frmMargenRep.Show
End Sub

Private Sub mnuInfOrdCom_Click()
With frmInfOrdCom
    .cmdSeleccionar.Enabled = False
    .Show vbModal
End With
End Sub

Private Sub mnuInformeComicionesSrvt_Click()
frmComisionesServiteca.Show
End Sub

Private Sub mnuInfProdMec_Click()
frmInfProdMec.Show
End Sub

Private Sub mnuInfRentOT_Click()
frmRentabilidadOT.Show
End Sub

Private Sub mnuInfResDed_Click()
frmResumenDeducibles.Show
End Sub

Private Sub mnuInfResumenProveedores_Click()
frmResumenProveedores.Show
End Sub

Private Sub mnuInfResValOT_Click()
frmResumenValorizadoOt.Show
End Sub

Private Sub mnuMaestroComisionesSrvt_Click()
frmMaestroFactores.Show 1
End Sub

Private Sub mnuManoObra_Click()
gstrBusca = CostoManoObra(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuMecanicos_Click()
gstrBusca = Mecanicos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuModeloVehiculo_Click()
    If Not Atributos("Glbl", "Tllr_10_0170", True, True, True, True) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Exit Sub
    Else
        gstrBusca = apfFormulario.Modelo(Conexion, "Tllr", "", "", gstrIdUsuario, apninguno, "", "", gstrIdEmpresa, gstrIdSucursal)
    End If
End Sub

Private Sub mnuMotivoAusencia_Click()
gstrBusca = MotivoAusencia(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuMovEmiOrdComTer_Click()
gstrBusca = OrdenesCompra(Conexion, gstrIdUsuario, "Tllr", "TLLR_20_0040", gstrIdEmpresa, gstrPathReporte, "", apninguno)
End Sub

Private Sub mnuOpcServiteca_Click()
frmOpcionesServiteca.Show 1
End Sub

Private Sub mnuOT_Click()
gstrBusca = ""
frmOtServiteca.Show
End Sub

Private Sub mnuParametrosSistema_Click()
frmMantenedorParametros.Show
End Sub

Private Sub mnuPresupuestoMantenciones_Click()
frmPresupuestoMantenciones.Show
End Sub

Private Sub mnuProveedorServicio_Click()
If Not Atributos("Glbl", "Tllr_10_0110_0010", True, True, True, True) Then
    MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
    Exit Sub
Else
    gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, gstrPathReporte, "", "", apcrear, "Proveedor", gstrIdSucursal)
End If
End Sub



Private Sub mnuRecepcionMecCar_Click()
gstrBusca = Recepciones(Conexion, gstrIdUsuario, "Tllr", " ", gstrIdEmpresa, LetConnectionString("TLLR", "RPT", "AUTOPRO", 256), "", apcrear)
gstrImpresion = "R"  'IMPRESION DE RECEPCION
gstrProcedencia = "Recepcion"
End Sub

Private Sub mnuRecordatorio_Click()
frmRecordatorioServicio.Show
End Sub

Private Sub mnuReemplazaVinPlaca_Click()
frmReemplazaVinxPatente.Show vbModal
End Sub

Private Sub mnuRelojTareas_Click()
FrmIngtareas.Show vbModal
End Sub

Private Sub mnuRentaCarroceria_Click()
frmRentabilidadCarroceria.Show
End Sub

Private Sub mnuRepuestosGeneral_Click()
frmRentabilidadRepuestos.Show
End Sub

Private Sub mnuRepuestosMecanico_Click()

End Sub

Private Sub mnuReservaAten_Click()
frmRptReservaAten.Show
End Sub

Private Sub mnuReservaHoras_Click()
gstrBusca = ReservaDeHoras(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, gstrPathReporte, "", apninguno)
gstrProcedencia = "ReservaHora"
End Sub

Private Sub mnuReservaRepuestos_Click()
frmInfReservaRepuestos.Show
End Sub

Private Sub mnuResumenServiteca_Click()
frmResumenServiteca.Show
End Sub

Private Sub mnuSalirSistema_Click()
Unload Me
End Sub

Private Sub mnuServiceRate_Click()
frmInfServiceRateReserva.Show
End Sub

Private Sub mnuServiciosGenerales_Click()
gstrBusca = ServiciosGenerales(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuServiciosServiteca_Click()
gstrProcedencia = "Mantenedor"
gstrBusca = ServiciosSrvt(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuTempServActRepVsMarMod_Click()
gstrBusca = TemparioServicios(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuTpoCargo_Click()
gstrBusca = TipoCargo(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuTpoOT_Click()
gstrBusca = TipoGarantias(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuTurnos_Click()
gstrBusca = Turnos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuVehiculos_Click()
gstrProcedencia = "Mantenedor"
gstrBusca = Vehiculos(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub mnuVehiculosPropios_Click()
gstrProcedencia = "MantenedorPropio"
gstrBusca = VehiculosPropios(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", apninguno)
End Sub

Private Sub popup_Click(Index As Integer)
Dim i As Integer
Dim dblTotalInicial As Double
Dim dblDescuento As Double

'kjcv 28.08.14

Const CargoGtiaFab As String = "04"
Dim gTipoCargoActual As String

Dim Sql As String
Dim AdoCargo As New ADODB.Recordset
Dim Cargos(9) As String
Dim ldblCont As Integer

'kjcv 07.04.15
Sql = "SELECT Id_Cargo FROM Tllr_Mecanicos_Cargo WHERE Id_Empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Mecanico='" & gstrIdUsuario & "'"
If Conexion.SendHost(Sql, AdoCargo, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If AdoCargo.EOF = False And AdoCargo.BOF = False Then
        ldblCont = 1
        AdoCargo.MoveFirst
        While AdoCargo.EOF = False
            Cargos(ldblCont) = ValorNulo(AdoCargo.Fields("Id_Cargo"))
            ldblCont = ldblCont + 1
            AdoCargo.MoveNext
        Wend
    End If
End If
Conexion.CloseHost AdoCargo


Select Case Index
    Case 3 'descuentos
        gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
        If IsNumeric(gstrBusca) Then
            If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
    
                If gstrProcedenciaBotonDerecho = "Mecanica" Then
                    For i = 1 To frmRecepcion.lvwServiciosMecanica.ListItems.Count
                        If frmRecepcion.lvwServiciosMecanica.ListItems(i).Selected Then
                            dblTotalInicial = Round(CDbl(frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(2)) * CDbl(frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(3)), 2)
                            frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                            frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                            frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                        End If
                    Next
                    frmRecepcion.AsignaTotal mcFichaMecanica, frmRecepcion.stbTotalMec
                    frmRecepcion.TotalFinal
                ElseIf gstrProcedenciaBotonDerecho = "Otros" Then
                    For i = 1 To frmRecepcion.lvwOtrosServicios.ListItems.Count
                        If frmRecepcion.lvwOtrosServicios.ListItems(i).Selected Then
                            dblTotalInicial = Round(CDbl(frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(2)) * CDbl(frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(3)), 2)
                            frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                            frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                            frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                        End If
                    Next
                    frmRecepcion.AsignaTotal mcFichaOtros, frmRecepcion.stbTotalOtros
                    frmRecepcion.TotalFinal
                ElseIf gstrProcedenciaBotonDerecho = "Carroceria" Then
                    For i = 1 To frmRecepcion.lvwServiciosCarroceria.ListItems.Count
                        If frmRecepcion.lvwServiciosCarroceria.ListItems(i).Selected Then
                            If Trim(frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(5)) <> "0.0" Then
                                dblTotalInicial = Round(CDbl(frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(5)) * CDbl(frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(9)), 2)
                                frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(11) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                                frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(16) = FormatoValor(dblTotalInicial - CDbl(frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(11)), "", gintDecimalesMoneda)
                                frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(10) = FormatoValor(Val(gstrBusca), "", 2)
                            End If
                        End If
                    Next
                    frmRecepcion.AsignaTotal mcFichaCarroceria, frmRecepcion.stbTotalCarroceria
                    frmRecepcion.TotalFinal
                ElseIf gstrProcedenciaBotonDerecho = "Terceros" Then
                    For i = 1 To frmRecepcion.lvwServiciosTerceros.ListItems.Count
                        If frmRecepcion.lvwServiciosTerceros.ListItems(i).Selected Then
                            If Trim(frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(6)) <> "0.0" Then
                                dblTotalInicial = Round(CDbl(frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(6)) * CDbl(frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(9)), 2)
                                frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(11) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                                frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(12) = FormatoValor(dblTotalInicial - CDbl(frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(11)), "", gintDecimalesMoneda)
                                frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(10) = FormatoValor(Val(gstrBusca), "", 2)
                            End If
                        End If
                    Next
                    frmRecepcion.AsignaTotal mcFichaTerceros, frmRecepcion.stbTotalTerceros
                    frmRecepcion.TotalFinal
                ElseIf gstrProcedenciaBotonDerecho = "Repuestos" Then
                    If Val(gstrBusca) <= gintDescuentoMaximo Then
'kjcv 13.03.17
'                    If Val(gstrBusca) <= gintDescuentoMaximo Or (Val(gstrBusca) <= gintDescuentoMaximoCIA) Then
                        For i = 1 To frmRecepcion.lvwRepuestos.ListItems.Count
                            If frmRecepcion.lvwRepuestos.ListItems(i).Selected Then
                                dblTotalInicial = Round(CDbl(frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(2)) * CDbl(frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(3)), 2)
                                dblDescuento = ValorPorcentaje(dblTotalInicial, Val(gstrBusca))
                                If gblnValidaCostoRepuestos = True Then
                                    If CostoRepuesto(frmRecepcion.lvwRepuestos.ListItems.Item(i), CDbl(frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(2))) < (dblTotalInicial - dblDescuento) Then
                                        frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(5) = FormatoValor(dblDescuento, "", gintDecimalesMoneda)
                                        frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(8) = FormatoValor(dblTotalInicial - CDbl(frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                                        frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                                    Else
                                        If MsgBox("El Valor Venta del Repuesto " & frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(1) & " " & Chr(13) & _
                                                  "Es menor que el Precio de Costo " & Chr(13) & "Desea Continuar...", vbQuestion + vbYesNo, "Confirma Precio Venta Repuesto") = vbYes Then
                                            
                                            Screen.MousePointer = 1
                                            gblnDescuentoRepuesto = True
                                            frmPermisoDiasHabiles.Show 1
                                            
                                            If NoEsLaPassword(gstrVerificacion, gstrMecanicoDiasHabiles) Then
                                                frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(5) = FormatoValor(dblDescuento, "", gintDecimalesMoneda)
                                                frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(8) = FormatoValor(dblTotalInicial - CDbl(frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                                                frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                                            Else
                                                MsgBox "Lo Siento, La passWord ingresada no es Correcta", vbExclamation, "Password"
                                            End If
                                            gblnDescuentoRepuesto = False
                                        End If
                                    End If
                                Else
                                    frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(5) = FormatoValor(dblDescuento, "", gintDecimalesMoneda)
                                    frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(8) = FormatoValor(dblTotalInicial - CDbl(frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                                    frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                                End If
                            End If
                        Next
                        frmRecepcion.AsignaTotal mcFichaRepuestos, frmRecepcion.stbTotalRepuestos
                        frmRecepcion.TotalFinal
                    Else
                        MsgBox "El descuento ingresado es mayor que el permitido", vbExclamation, "Advertencia"
                    End If
                End If
                
            Else
                MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
            End If
        Else
            MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
        End If
            
    Case 4 'Cargo
        'kjcv 24.03.20
    gstrBusca = ""
    frmTipoCargo.Show vbModal
'        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Tipo_Cargo", "Id_Tipo_cargo", "Descripcion", "Buscar Cargo OT")
'        Dim Cargos
'        Cargos = CargosMecanicos(gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario)
        If gstrBusca = "03" Then
            frmCentroCosto.Show vbModal
        End If
        If gstrBusca <> "" Then
            If gstrProcedenciaBotonDerecho = "Mecanica" Then
                For i = 1 To frmRecepcion.lvwServiciosMecanica.ListItems.Count
                    If frmRecepcion.lvwServiciosMecanica.ListItems(i).Selected Then
                    ' inicio kjcv 07.04.15
                        For j = 1 To maxCargos
                            If gstrBusca = Cargos(j) Then
                                frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(6) = gstrBusca
                                frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(7) = TraeCargoDes(gstrBusca)
                                ValidaCostoCargoMecanica gstrBusca, frmRecepcion.lvwServiciosMecanica, i
                            End If
                        Next j
                    
                    End If
                Next
                frmRecepcion.AsignaTotal mcFichaMecanica, frmRecepcion.stbTotalMec
                frmRecepcion.TotalFinal
            ElseIf gstrProcedenciaBotonDerecho = "Otros" Then
                For i = 1 To frmRecepcion.lvwOtrosServicios.ListItems.Count
                    If frmRecepcion.lvwOtrosServicios.ListItems(i).Selected Then
                    ' inicio kjcv 07.04.15
                        For j = 1 To maxCargos
                            If gstrBusca = Cargos(j) Then
                                frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(6) = gstrBusca
                                frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(7) = TraeCargoDes(gstrBusca)
                                ValidaCostoCargoMecanica gstrBusca, frmRecepcion.lvwOtrosServicios, i
                            End If
                        Next j
                    'fin kjcv 07.04.15
                    
'                        frmRecepcion.lvwOtrosServicios.ListItems.item(i).SubItems(6) = gstrBusca
'                        frmRecepcion.lvwOtrosServicios.ListItems.item(i).SubItems(7) = TraeCargoDes(gstrBusca)
'                        ValidaCostoCargoMecanica gstrBusca, frmRecepcion.lvwOtrosServicios, i
                    End If
                Next
                frmRecepcion.AsignaTotal mcFichaOtros, frmRecepcion.stbTotalOtros
                frmRecepcion.TotalFinal
            ElseIf gstrProcedenciaBotonDerecho = "Carroceria" Then
                For i = 1 To frmRecepcion.lvwServiciosCarroceria.ListItems.Count
                    If frmRecepcion.lvwServiciosCarroceria.ListItems(i).Selected Then
                    ' inicio kjcv 07.04.15
                        For j = 1 To maxCargos
                            If gstrBusca = Cargos(j) Then
                                frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(13) = gstrBusca
                                frmRecepcion.lvwServiciosCarroceria.ListItems.Item(i).SubItems(12) = TraeCargoDes(gstrBusca)
                                ValidaCostoCargoCarroceria gstrBusca, i
                            End If
                        Next j
                    'fin kjcv 07.04.15
                    
''                        frmRecepcion.lvwServiciosCarroceria.ListItems.item(i).SubItems(13) = gstrBusca
''                        frmRecepcion.lvwServiciosCarroceria.ListItems.item(i).SubItems(12) = TraeCargoDes(gstrBusca)
''                        ValidaCostoCargoCarroceria gstrBusca, i
                    End If
                Next
                frmRecepcion.AsignaTotal mcFichaCarroceria, frmRecepcion.stbTotalCarroceria
                frmRecepcion.TotalFinal
            ElseIf gstrProcedenciaBotonDerecho = "Terceros" Then
                For i = 1 To frmRecepcion.lvwServiciosTerceros.ListItems.Count
                    If frmRecepcion.lvwServiciosTerceros.ListItems(i).Selected Then
                    ' inicio kjcv 07.04.15
                        For j = 1 To maxCargos
                            If gstrBusca = Cargos(j) Then
                                frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(14) = gstrBusca
                                frmRecepcion.lvwServiciosTerceros.ListItems.Item(i).SubItems(13) = TraeCargoDes(gstrBusca)
                                ValidaCostoCargoTerceros gstrBusca, i
                            End If
                        Next j
                    'fin kjcv 07.04.15
                    
''                        frmRecepcion.lvwServiciosTerceros.ListItems.item(i).SubItems(14) = gstrBusca
''                        frmRecepcion.lvwServiciosTerceros.ListItems.item(i).SubItems(13) = TraeCargoDes(gstrBusca)
''                        ValidaCostoCargoTerceros gstrBusca, i
                    End If
                Next
                frmRecepcion.AsignaTotal mcFichaTerceros, frmRecepcion.stbTotalTerceros
                frmRecepcion.TotalFinal
            ElseIf gstrProcedenciaBotonDerecho = "Repuestos" Then
                For i = 1 To frmRecepcion.lvwRepuestos.ListItems.Count
                    If frmRecepcion.lvwRepuestos.ListItems(i).Selected Then
                        'kjcv 27.08.14 evalua
                        gTipoCargoActual = frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(7)
                        ' inicio kjcv 07.04.15
                        For j = 1 To maxCargos
                            If gstrBusca = Cargos(j) Then
                                frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(7) = gstrBusca
                                frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(6) = TraeCargoDes(gstrBusca)
                                ValidaCostoCargoRepuestos gstrBusca, i, frmRecepcion.lvwRepuestos.ListItems.Item(i)
                            End If
                        Next j
                        
                        If gstrBusca = "03" Then
                        frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(14) = gCentroCosto
                        Else
                        frmRecepcion.lvwRepuestos.ListItems.Item(i).SubItems(14) = ""
                        End If
                        'fin kjcv 07.04.15
                        
''                        If gstrBusca = CargoGtiaFab Or gTipoCargoActual = CargoGtiaFab Then
''                            MsgBox "Cambio de Cargo NO Permitido..", vbCritical, "Elisa Taller"
''                            Exit Sub
''                        End If
''                        frmRecepcion.lvwRepuestos.ListItems.item(i).SubItems(7) = gstrBusca
''                        frmRecepcion.lvwRepuestos.ListItems.item(i).SubItems(6) = TraeCargoDes(gstrBusca)
''                        ValidaCostoCargoRepuestos gstrBusca, i, frmRecepcion.lvwRepuestos.ListItems.item(i)
                    End If
                Next
                frmRecepcion.AsignaTotal mcFichaRepuestos, frmRecepcion.stbTotalRepuestos
                frmRecepcion.TotalFinal
            End If
        End If
    
    Case 5 'Mecánico
        
        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(select * from Tllr_Mecanicos where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "' And Vigencia='S') as Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Buscar Mecánico")
        If gstrBusca <> "" Then
            
            If gstrProcedenciaBotonDerecho = "Mecanica" Then
                For i = 1 To frmRecepcion.lvwServiciosMecanica.ListItems.Count
                    If frmRecepcion.lvwServiciosMecanica.ListItems(i).Selected Then
                        frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(8) = gstrBusca
                        frmRecepcion.lvwServiciosMecanica.ListItems.Item(i).SubItems(9) = TraeNombreMecanico(gstrBusca)
                    End If
                Next
            ElseIf gstrProcedenciaBotonDerecho = "Otros" Then
                For i = 1 To frmRecepcion.lvwOtrosServicios.ListItems.Count
                    If frmRecepcion.lvwOtrosServicios.ListItems(i).Selected Then
                        frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(8) = gstrBusca
                        frmRecepcion.lvwOtrosServicios.ListItems.Item(i).SubItems(9) = TraeNombreMecanico(gstrBusca)
                    End If
                Next
            End If
        End If
    
End Select
End Sub
Sub ValidaCostoCargoMecanica(strIdCargo As String, lvwObjeto As ListView, Indice As Integer)
Dim lstrCostea As String
Dim lstrSQL As String
Dim dblTotalInicial As Double
Dim recAux As New ADODB.Recordset

If strIdCargo <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & strIdCargo & "'", gcdynamic)
    If lstrCostea = "S" Then
        
        If gblnPreciosMarca = True Then
            'trae costo de hora por marca
            lstrSQL = "SELECT CostoManoObra, CostoMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & frmRecepcion.lblIdMarca & "')"
            If Conexion.SendHost(lstrSQL, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not recAux.BOF And Not recAux.EOF Then
'                    lvwObjeto.ListItems.item(Indice).SubItems(3) = FormatoValor(IIf(strIdCargo = gstrCargoGtiaFabrica, recAux!CostoMOGarantia, recAux!CostoManoObra), "", gintDecimalesMoneda)
'kjcv 07.10.15 modificacion de cargo GARANTIA FABRICA 29.10.15 se agrego cortesia comercial
                    If strIdCargo = gstrCargoGtiaFabrica Or strIdCargo = "06" Or strIdCargo = "07" Or strIdCargo = "08" Then
                        lvwObjeto.ListItems.Item(Indice).SubItems(3) = FormatoValor(recAux!CostoMOGarantia, "", gintDecimalesMoneda)
                    'kjcv 09.09.16
                    ElseIf strIdCargo = "02" Then
                        lvwObjeto.ListItems.Item(Indice).SubItems(3) = Round(traeValorHoraCS(gstrIdCompañiaSeg, gstrIdEmpresa) * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa) = 0, traeParidadMoneda("02"), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa)), 2)
                    Else
                        lvwObjeto.ListItems.Item(Indice).SubItems(3) = FormatoValor(recAux!CostoManoObra, "", gintDecimalesMoneda)
                    End If
                End If
            End If
        Else
            lvwObjeto.ListItems.Item(Indice).SubItems(3) = FormatoValor(gcurCostoManoObra, "", gintDecimalesMoneda)
        End If
        
        lvwObjeto.ListItems.Item(Indice).SubItems(10) = FormatoValor(CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(2)) * CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(3)), "", gintDecimalesMoneda)
        If lvwObjeto.ListItems.Item(Indice).SubItems(4) <> "" And lvwObjeto.ListItems.Item(Indice).SubItems(2) <> "" Then
            dblTotalInicial = CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(2)) * CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(3))
            lvwObjeto.ListItems.Item(Indice).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, CSng(lvwObjeto.ListItems.Item(Indice).SubItems(4))), "", gintDecimalesMoneda)
            lvwObjeto.ListItems.Item(Indice).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(5)), "", gintDecimalesMoneda)
        Else
            lvwObjeto.ListItems.Item(Indice).SubItems(10) = FormatoValor(CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(3)) * CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(2)), "", gintDecimalesMoneda)
            lvwObjeto.ListItems.Item(Indice).SubItems(4) = 0
            lvwObjeto.ListItems.Item(Indice).SubItems(5) = 0
        End If
    Else
        
        If gblnPreciosMarca = True Then
            'trae costo de hora por marca
            lstrSQL = "SELECT VentaManoObra, VentaMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & frmRecepcion.lblIdMarca & "')"
            If Conexion.SendHost(lstrSQL, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not recAux.BOF And Not recAux.EOF Then
'                    lvwObjeto.ListItems.item(Indice).SubItems(3) = FormatoValor(IIf(strIdCargo = gstrCargoGtiaFabrica, recAux!VentaMOGarantia, recAux!VentaManoObra), "", gintDecimalesMoneda)
'kjcv 07.10.15 modificacion por garantia fabrica 29.10.15 se agrego Cortesia Comercial
                    If strIdCargo = gstrCargoGtiaFabrica Or strIdCargo = "06" Or strIdCargo = "07" Or strIdCargo = "08" Then
                        lvwObjeto.ListItems.Item(Indice).SubItems(3) = FormatoValor(recAux!VentaMOGarantia, "", gintDecimalesMoneda)
                    'kjcv 09.09.16
                    ElseIf strIdCargo = "02" Then
                        lvwObjeto.ListItems.Item(Indice).SubItems(3) = Round(traeValorHoraCS(gstrIdCompañiaSeg, gstrIdEmpresa) * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa) = 0, traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa)), 2)
                    Else
                        lvwObjeto.ListItems.Item(Indice).SubItems(3) = FormatoValor(recAux!VentaManoObra, "", gintDecimalesMoneda)
                    End If
                End If
            End If
        Else
'            lvwObjeto.ListItems.item(Indice).SubItems(3) = IIf(strIdCargo = gstrCargoGtiaFabrica, FormatoValor(Retorna_Valor_General("Select PrecioManoOBraGarantia from Tllr_Parametro Where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'", gcdynamic), "", gintDecimalesMoneda), FormatoValor(gcurPrecioManoObra, "", gintDecimalesMoneda))
            'kjcv 07.10.15
'            lvwObjeto.ListItems.item(Indice).SubItems(3) = IIf(strIdCargo = gstrCargoGtiaFabrica, FormatoValor(Retorna_Valor_General("Select PrecioManoOBraGarantia from Tllr_Parametro Where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'", gcdynamic), "", gintDecimalesMoneda), FormatoValor(gcurPrecioManoObra, "", gintDecimalesMoneda))
'kjcv 29.10.15 Se Agrego cargo "Cortesia Comercial"
            If strIdCargo = gstrCargoGtiaFabrica Or strIdCargo = "06" Or strIdCargo = "07" Or strIdCargo = "08" Then
                'lvwObjeto.ListItems.item(Indice).SubItems(3) = FormatoValor(Retorna_Valor_General("Select PrecioManoOBraGarantia from Tllr_Parametro Where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'", gcdynamic), "", gintDecimalesMoneda)
                'kjcv 09.05.20 toma valores de tabla Tllr_MO
                lvwObjeto.ListItems.Item(Indice).SubItems(3) = FormatoValor(Retorna_Valor_General("select ValorMOGarantia from Tllr_mo where id_empresa='" & gstrIdEmpresa & "' and Id_Marca = '" & frmRecepcion.lblIdMarca & "'", gcdynamic), "", gintDecimalesMoneda)
                
            'kjcv 09.09.16
            'kjcv 09.09.16
            ElseIf strIdCargo = "02" Then
                lvwObjeto.ListItems.Item(Indice).SubItems(3) = traeValorHoraCS(gstrIdCompañiaSeg, gstrIdEmpresa) * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa) = 0, traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa))
            Else
                lvwObjeto.ListItems.Item(Indice).SubItems(3) = FormatoValor(gcurPrecioManoObra, "", gintDecimalesMoneda)
            End If
            
        End If
    
        lvwObjeto.ListItems.Item(Indice).SubItems(10) = FormatoValor(CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(2)) * CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(3)), "", gintDecimalesMoneda)
        If lvwObjeto.ListItems.Item(Indice).SubItems(4) <> "" And lvwObjeto.ListItems.Item(Indice).SubItems(2) <> "" Then
            dblTotalInicial = CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(2)) * CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(3))
            lvwObjeto.ListItems.Item(Indice).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, CSng(lvwObjeto.ListItems.Item(Indice).SubItems(4))), "", gintDecimalesMoneda)
            lvwObjeto.ListItems.Item(Indice).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(lvwObjeto.ListItems.Item(Indice).SubItems(5)), "", gintDecimalesMoneda)
        End If
    End If
End If
End Sub

Private Sub ValidaCostoCargoCarroceria(strIdCargo As String, Indice As Integer)
Dim lstrCostea As String
Dim dblMtoinicial As Double

If strIdCargo <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & strIdCargo & "'", gcdynamic)
    If lstrCostea = "S" Then
        frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(16) = FormatoValor(CDbl(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5)) * CDbl(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6)), "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(7) = FormatoValor(0, "", 2)
        frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(8) = FormatoValor(0, "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(10) = FormatoValor(0, "", 2)
        frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(11) = FormatoValor(0, "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9) = frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(16)
    Else
        frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(16) = FormatoValor(CDbl(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5)) * CDbl(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6)), "", gintDecimalesMoneda)
        
        'recargo
        dblMtoinicial = 0
        If Trim(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(7)) <> "" Then
            dblMtoinicial = CDbl(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5), "0")) * CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6), "0"))
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(8) = FormatoValor(ValorPorcentaje(CDbl(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6), "0")), CSng(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(7) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(7), "0"))), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9) = FormatoValor(CDbl(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6), "0")) + CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(8) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(8), "0")), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(16) = FormatoValor(CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9), "0")) * CDbl(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5), "0")), "", gintDecimalesMoneda)
        Else
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(7) = FormatoValor(0, "", 2)
        End If
        
        'descuento
        dblMtoinicial = 0
        If Trim(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(10)) <> "" Then
            dblMtoinicial = CDbl(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5), "0")) * CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6), "0"))
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9) = FormatoValor(CDbl(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(6), "0")) + CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(8) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(8), "0")), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(11) = FormatoValor(ValorPorcentaje(CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9), 0)) * CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5), 0)), CSng(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(10) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(10), "0"))), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(16) = FormatoValor((CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(9), "0")) * CDbl(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(5), "0"))) - CCur(IIf(frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(11) <> "", frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(11), 0)), "", gintDecimalesMoneda)
        Else
            frmRecepcion.lvwServiciosCarroceria.ListItems(Indice).SubItems(10) = "0"
        End If
    End If
End If
End Sub
Private Sub ValidaCostoCargoTerceros(strIdCargo As String, Indice As Integer)
Dim lstrCostea As String
Dim dblMtoinicial As Double

If strIdCargo <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & strIdCargo & "'", gcdynamic)
    If lstrCostea = "S" Then
        frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(12) = FormatoValor(CDbl(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6)) * CDbl(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5)), "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(7) = FormatoValor(0, "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(8) = FormatoValor(0, "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(10) = FormatoValor(0, "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(11) = FormatoValor(0, "", gintDecimalesMoneda)
        frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9) = frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(12)
    Else
        frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(12) = FormatoValor(CDbl(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6)) * CDbl(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5)), "", gintDecimalesMoneda)
        
        'recargo
        dblMtoinicial = 0
        If Trim(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(7)) <> "" Then
            dblMtoinicial = CDbl(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6), "0")) * CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5), "0"))
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(8) = FormatoValor(ValorPorcentaje(CDbl(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5), "0")), CSng(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(7) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(7), "0"))), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9) = FormatoValor(CDbl(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5), "0")) + CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(8) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(8), "0")), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(12) = FormatoValor(CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9), "0")) * CDbl(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6), "0")), "", gintDecimalesMoneda)
        Else
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(7) = "0"
        End If
        
        'descuento
        dblMtoinicial = 0
        If Trim(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(10)) <> "" Then
            dblMtoinicial = CDbl(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6), "0")) * CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5), "0"))
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9) = FormatoValor(CDbl(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(5), "0")) + CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(8) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(8), "0")), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(11) = FormatoValor(ValorPorcentaje(CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9), 0)) * CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6), 0)), CSng(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(10) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(10), "0"))), "", gintDecimalesMoneda)
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(12) = FormatoValor((CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(9), "0")) * CDbl(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(6), "0"))) - CCur(IIf(frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(11) <> "", frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(11), 0)), "", gintDecimalesMoneda)
        Else
            frmRecepcion.lvwServiciosTerceros.ListItems(Indice).SubItems(10) = "0"
        End If
    
    End If
End If
End Sub
Private Sub ValidaCostoCargoRepuestos(strIdCargo As String, Indice As Integer, strIdItem As String)
Dim lstrCostea As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset
Dim lstrTipo As String
Dim gPrecioUnitario As Double

If gstrIdCargo <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & strIdCargo & "'", gcdynamic)
    lstrTipo = Retorna_Valor_General("Select Tipo from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & strIdCargo & "'", gcdynamic)
    If lstrCostea = "S" Then
        '//LREYES...
        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(4) = FormatoValor(0, "", 2)
        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(5) = FormatoValor(0, "", gintDecimalesMoneda)
        'error grande multiplicaba como 4 veces el precio unitario
        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(CostoRepuesto(strIdItem, CDbl(frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(2))) * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(8) = FormatoValor(CDbl(frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3)) * CDbl(frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(2)), "", gintDecimalesMoneda)
    Else
        '//LREYES...
'        strSql = "select isnull(precio_venta,0) as precio_venta from stck_item where id_item='" & strIdItem & "'"
        'kjcv 25.08.14 se agrega CUP
'        strSql = "select isnull(precio_venta,0) as precio_venta, isnull(CUP,0) as CUP ,isnull(precio_costo,0) as precio_costo from stck_item where id_item='" & strIdItem & "'"
        'kjcv 09.11.15
        If lstrTipo = "C" Then
            strSql = "select isnull(precio_venta,0) as precio_venta, isnull(Precio_Venta_Por_Mayor,0) as precio_taller ,isnull(Precio_Venta_CIA,0) as precio_CIA from stck_item where id_item='" & strIdItem & "'"
        Else
            strSql = " Select top 1 isnull(Precio_Unitario,0) as Precio_Costo From Stck_Mayor_auxiliar where Tipo_Movto = 'E' and id_tipo_docto in ('FC','RG') and Id_Item='" & strIdItem & "' Order By Fecha desc "
        End If

        If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(4) = FormatoValor(0, "", 2)
                frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(5) = FormatoValor(0, "", gintDecimalesMoneda)
'                'kjcv 01.03.13 cambio de taller a dolares.
'                'frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(adoTemp!Precio_Venta, "", gintDecimalesMoneda)
                'kjcv 06.11.12 Por cargo a Compañia de seguros y Deducibles
                If strIdCargo = "02" Then
'                    frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(adoTemp!Precio_Venta * traeParidadMoneda("02"), "", gintDecimalesMoneda)
                    frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(adoTemp!Precio_CIA * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa) = 0, traeParidadMoneda("02"), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa)), "", gintDecimalesMoneda)
                    gPrecioUnitario = adoTemp!Precio_CIA
                    frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(15) = gPrecioUnitario
                Else
                    If strIdCargo = "04" Or strIdCargo = "06" Or strIdCargo = "07" Then
'                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor((1 + (gstrPorPrecioGtia / 100)) * adoTemp!CUP * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
                        'kjcv 21.10.15 tipo cambio gtia fabrica
                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor((1 + (gstrPorPrecioGtia / 100)) * adoTemp!Precio_Costo * traeParidadMonedaMesGarantia("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
                        gPrecioUnitario = Round((1 + (gstrPorPrecioGtia / 100)) * adoTemp!Precio_Costo, 2)
                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(15) = gPrecioUnitario
                        'kjcv 18.08.15 quitar el mark up
'                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(adoTemp!CUP * traeParidadMonedaMesGarantia("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
'kjcv 29.10.15 cargo Cortesi Comercial
                    ElseIf strIdCargo = "08" Then
                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(adoTemp!Precio_Costo * traeParidadMonedaMesGarantia("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
                        gPrecioUnitario = adoTemp!Precio_Costo
                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(15) = gPrecioUnitario
                    Else
'                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(adoTemp!Precio_Venta * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
                        'kjcv 10.11.15
                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor(adoTemp!Precio_taller * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
                        gPrecioUnitario = (1 + (gstrPorPrecioGtia / 100)) * adoTemp!Precio_taller
                        frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(15) = gPrecioUnitario
                    End If
                End If
                'kjcv 25.08.14
'                If strIdCargo = "04" Then
'                    frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) = FormatoValor((adoTemp!CUP / 0.95) * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), "", gintDecimalesMoneda)
'                End If

                
                frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(8) = FormatoValor(frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(3) * CDbl(frmRecepcion.lvwRepuestos.ListItems(Indice).SubItems(2)), "", gintDecimalesMoneda)
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End If
End Sub

