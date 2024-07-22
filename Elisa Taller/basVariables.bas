Attribute VB_Name = "basVariables"
Option Explicit
Public strIdMarcaDefecto As String
Public blnContinuar As Boolean
'Public Server As New apServer
'Public Perfil As New APPERFIL1.APPERFIL
Public Conexion As New APCONADO.ConnectionAdo
Public cnnAux As New ADODB.Connection
Public Libreria As New ElisaLibs.Llamadas
'Public apfFormulario As New APFORM1.APFORM
Public apfFormulario As New APFORM2.APFORM
'Public apfLogin As New apLogin.Login
Public Const gcTiempoEspera = 30
Public gintTimeOut As Long
Public strRetorno As String
Public intTamaño As Integer
Public intValido As Integer
Public strConnect As String
Public strMode As String '///////crear nuevo, modificar, etc
Public gadoPrincipal As New ADODB.Recordset
Public gmstrProcedencia As String
'KJCV 09.07.15
Public gstrPorPrecioGtia As Single
Public gstrIdPerfil As String
Public gflag As Boolean
'//////////////////////////VARIABLE DE DOMINIO >
Public gstrRazonSocial As String
Public gstrPassWordLiquidador As String
Public gstrVerificacion As Integer
Public gstrVerificaMecanico As String
Public gstrArchivoIni As String
Public gstrEmpresa As String
Public gstrSucursal As String
Public gstrDirSuc As String
Public gstrUsuario As String
Public gstrIdEmpresa As String
Public gstrIdSucursal As String
Public gstrIdUsuario As String
Public gstrIdEmpleado As String
Public gstrTelefono As String
Public gstrFax As String
Public gintDiasHabiles As Integer
Public entraRecepcion As Boolean
Public gstrCodigoAcceso As String
Public gstrPrefijoSistema As String
Public gstrInicial As String * 1
Public gstrBusca As String
Public gstrPathReporte As String
Public gstrRutaApclient As String
Public gstrMailRepuestosFallidos As String
Public gstrCodigoLubricantes As String
Public gstrCodigoMateriales As String
Public gstrCodigoInsumos As String
Public glsiItem As ListItem
Public gintProcedencia As Integer
Public gstrIdCargo As String
Public gcurPrecioManoObra As Currency
Public gcurCostoManoObra As Currency
Public gcurInsumo As Currency
Public gcurMaterialesMO As Currency
Public gcurMaterialesPesos As Currency
Public gcurMateriales As Currency
Public gcurLubricantes As Currency
Public gdblNroHorOblg As Currency
Public gdblValorExistencia As Currency
Public gintNumeroLineasRecepcion As Integer
Public gblnDescuentoRepuesto As Boolean
Public gstrMonedaLocal As String
Public gintDecimalesMoneda As Integer
Public gcurTotalNeto As Currency
Public gcurTotalIVA As Currency
Public gcurTotalNetoMasIVA As Currency
Public gcurSeguroTaller As Currency
Public gintNroRecDefectoQry As Integer
Public gstrIdCargoDefecto As String
Public gstrIdTipoOtDefecto As String

Public gstrIdTipoOtDefectoElegir As String
Public gstrMecanicoDefectoSecMec  As String
Public gstrMecanicoDefectoSecCar As String
Public gstrMecanicoDefectoSecDes As String
Public gstrMecanicoDefectoSecPin As String
Public gstrMecanicoDiasHabiles As String
Public gapColor As apColor
Public gapAccion As apAccion
Public gstrProcedencia As String
Public gstrProcedenciaBotonDerecho As String
Public gstrProcedenciaRptos As String
Public gstrSql As String
Public gstrSql2 As String
Public gstrSeccion As String * 1
Public gstrEstadoOT As String * 1
Public gintFila As Integer
Public gintColumna As Integer
Public gitmActual As Integer
Public gstrImpresion As String
Public gstrEstadoProdMecanico As String
Public gblnEnviaMailBodega As Boolean
Public gblnTraspasaRepuestos As Boolean
Public gblnImprimeImagen As Boolean
Public gblnValidaCostoRepuestos As Boolean
Public gblnCierraLiq As Boolean
Public gblnPreciosMarca As Boolean
Public gstrServiciosMarca As String
Public gblnBloqueaSubtotalRep As Boolean
Public gblnValidaServiciosCero  As Boolean
Public gintOtExistente As Integer
Public gstrCargoDeducibleMas As String
Public gstrCargoDeducibleMenos As String
Public gstrAsignaRecursos As String
Public gstrCargoGtiaFabrica As String
Public gstrIdCargoInterno As String
Public gintHoraInicio As Integer
Public gintHoratermino As Integer
Public gintIntervaloMinutos As Integer
Public gintDescuentoMaximo As Integer
Public gintDescuentoMaximoCIA As Integer
Public gcRecepcionMecanica As String
Public gcRecepcionCarroceria As String
Public gstrNotaPresupuesto As String
Public gstrNotaRecepcion As String
Public gstrNombreRut As String
Public gstrValidaRut As String
Public gstrEditaRut As String
Public gstrNombrePatente As String
Public gstrIdMecanico As String
Public gstrValidaPatente As String
Public gstrNombreIva As String
Public gstrNombreDP As String
Public gstrNombreComuna As String
Public gstrNombreCiudad As String
Public gstrNombreSucursal As String
Public gstrNombreBodega As String
Public gObjListView As ListView
Public blnPassLogin As Boolean
Public gstrIdCompañiaSeg As String
Public gstrNombreRecepcionista As String
Public gstrNombreRecepLlamado As String
Public swActivateRecorda As Boolean
Public ReporteRuta As String
Public Act As Integer
Public gCentroCosto As String
Public gUsr_Activacion As String
Public gstrEstado As String

Public gCliente As String

'////// picoro vin existencia
Type TIPO_PARAMETROS_CONTABLES
    Cont_id_Tipo_Comprobante_Diario As String
    Cont_id_Tipo_Comprobante_Mes_Anterior As String
    Cont_id_Tipo_Docto As String
    Cont_Iva As String
    Cont_Proveedor As String
    Cont_id_Tipo_Auxiliar As String
End Type

Public MesManoObra(1 To 12)

Public Declare Function tapiRequestMakecall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)

'//Para enviar datos a apClient...
Public Type Cliente
    Comando As String * 25
    CodigoEmpresa As String * 25
    NombreEmpresa As String * 50
    CodigoSucursal As String * 25
    NombreSucursal As String * 50
    CodigoUsuario As String * 25
    NombreUsuario As String * 50
    Modulo As String * 25
    hwnd As String * 25
    ArchivoINI As String * 50
    
End Type
Public DatosCliente As Cliente

Public Type VentaRepuestos
    Repuestos As Double
    Descuentos As Double
    
End Type

Public Type HojaRecurso
    Id_Sucursal As String
    Id_Turno As String
    Id_Item As String
    Id_Mecanico As String
    Id_Fecha As String
    Horas As Double
End Type

Public Type TIPO_DATOS_REPORTE
    NombreEmpresa As String
    RutEmpresa As String
    DireccionEmpresa As String
    iD_Usuario As String
    TituloReporte As String
End Type

Type VariablesGlobales
    str_Resultado_CODIGO As String
    str_Resultado_DESCRIPCION As String
    lsvPasaListaParaConfigurar As ListView
    strPasaTitulo As String
    strPasaNombreItemIni As String
    strPasaStringDefecto As String
    blnPassLogin As Boolean
    Buscar_RESULTADO_CODIGO As String
    Buscar_RESULTADO_DESCRIPCION As String
    Buscar_CAMPO_CODIGO As String
    Buscar_CAMPO_CODIGO_PARAM1 As String
    Buscar_CAMPO_CODIGO_PARAM2 As String
    Buscar_CODIGO_PARAM1 As String
    Buscar_CODIGO_PARAM2 As String
    Buscar_CAMPO_DESCRIPCION As String
    Buscar_TABLA As String
    Buscar_TITULO As String
End Type


Public gstrPresionoEnter As String

Global VGlob As VariablesGlobales
