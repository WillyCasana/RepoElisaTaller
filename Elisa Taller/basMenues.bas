Attribute VB_Name = "basMenues"

Public Function Mecanicos(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuMecanicos"
    frmMantenedorMecanicos.Show vbModal
    Mecanicos = gstrBusca
End Function

Public Function ActividadesGenerales(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuActividadesGenerales"
    frmMantenedorActividades.Show vbModal
    ActividadesGenerales = gstrBusca
End Function
Public Function CostoManoObra(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuActividadesGenerales"
    frmMantenedorCostoManoObra.Show vbModal
    CostoManoObra = gstrBusca
End Function

Public Function Promocion(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuActividadesGenerales"
    frmPromocion.Show vbModal
    Promocion = gstrBusca
End Function

Public Function Trabajos(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa

    gstrBusca = strCodigoInicial
    gapAccion = Accion

    frmTrabajo.Show vbModal
    Trabajos = gstrBusca
End Function

Public Function CompañiaSeguro(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuCompañiaSeguro"
    frmMantenedorCompañiaSeguro.Show vbModal
    CompañiaSeguro = gstrBusca
End Function
Public Function Campañas(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    frmMantenedorCampañas.Show vbModal
    Campañas = gstrBusca
End Function


Public Function ServiciosGenerales(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuServiciosGenerales"
    frmMantenedorServicios.Show vbModal
    ServiciosGenerales = gstrBusca
End Function


Public Function Vehiculos(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuVehiculos"
    frmMantenedorVehiculoCliente.Show vbModal
    Vehiculos = gstrBusca
End Function
Public Function VehiculosPropios(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuVehiculos"
    frmMantenedorVehiculosPropios.Show vbModal
    VehiculosPropios = gstrBusca
End Function

Public Function ConceptoSrvt(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuVehiculos"
    frmMantenedorConceptosSrvt.Show vbModal
    ConceptoSrvt = gstrBusca
End Function
Public Function ServiciosSrvt(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuVehiculos"
    frmMantenedorServiciosServiteca.Show vbModal
    ServiciosSrvt = gstrBusca
End Function

Public Function ConceptosInventario(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuConceptosInventario"
    frmMantenedorConceptosInventario.Show vbModal
    ConceptosInventario = gstrBusca
    
End Function
Public Function Concesionarios(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuConcesionarios"
    frmMantenedorConcesionario.Show vbModal
    Concesionarios = gstrBusca
    
End Function


Public Function TipoConos(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuTipoConos"
    frmMantenedorTipoCono.Show vbModal
    TipoConos = gstrBusca
End Function

Public Function Especialidades(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuEspecialidades"
    frmMantenedorEspecialidad.Show vbModal
    Especialidades = gstrBusca
End Function
Public Function HorasMecanico(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    frmMantenedorHorasMecanicos.Show vbModal
    HorasMecanico = gstrBusca
End Function
Public Function Turnos(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    frmMantenedorTurnos.Show vbModal
    Turnos = gstrBusca
End Function
Public Function MotivoAusencia(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    frmMantenedorMotivoAusencia.Show vbModal
    MotivoAusencia = gstrBusca
End Function

Public Function ParametrosSistema(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    
    frmMantenedorParametros.Show
    ParametrosSistema = gstrBusca
End Function


Public Function PartePieza(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuPartePieza"
    frmMantenedorParteYPieza.Show vbModal
    PartePieza = gstrBusca
End Function

Public Function TemparioServicios(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    gstrProcedencia = "Temparios"
    frmTempServiciosMarMod.Show
    TemparioServicios = gstrBusca
End Function

Public Function TemparioCompañiasSeguro(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuTemparioCompañiasSeguro"
    frmTempCiaSeguro.Show
    TemparioCompañiasSeguro = gstrBusca
End Function


Public Function ConceptosDyP(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuConceptosDyP"
    frmMantenedorConceptosDyP.Show
    ConceptosDyP = gstrBusca
End Function
'frmMantenedorProveedorServTerc
Public Function ProveedorServTercero(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuProveedorServTercero"
    frmMantenedorProveedorServTerc.Show vbModal
    ProveedorServTercero = gstrBusca
End Function
Public Function ServiciosdeTercero(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuServiciosdeTercero"
    frmMantenedorServiciosTerceros.Show vbModal
    ServiciosdeTercero = gstrBusca
End Function
Public Function Recepciones(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    gstrIdCargo = gstrCargoDeducibleMas
    With frmRecepcion
        .Caption = "Recepción"
        .lblCorrelativo.Caption = "Recepción Nº :"
        .lblEstadoOT.Visible = False
        .lblEstadoOTValor.Visible = False
        .lblTipo.Visible = False
        .txtTipo.Visible = False
        .stbTotalMec.Visible = False
        .stbTotalCarroceria.Visible = False
        .stbTotalDesabolladura.Visible = False
        .stbTotalPintura.Visible = False
        .stbTotalArmeyDesarme.Visible = False
        .stbTotalOtros.Visible = False
        .stbTotalTerceros.Visible = False
        .stbTotalRepuestos.Visible = False
        .stbTotalMateriales.Visible = False
        .StbLubricantes.Visible = False
        .stbTotalOT.Visible = False
        .stbSeguroTaller.Visible = False
        .tlbAddRep.Visible = True
        .cmdConsultaStock.Visible = False
        .tlbBarraHerramientas.Buttons(13).Visible = False
        .tlbBarraHerramientas.Buttons(14).Visible = False
        .tlbBarraHerramientas.Buttons(15).Visible = False
        .tlbBarraHerramientas.Buttons(16).Visible = False
        .tlbBarraHerramientas.Buttons(21).Visible = False
        .tlbBarraHerramientas.Buttons(22).Visible = False
        .tlbBarraHerramientas.Buttons(23).Visible = False
        .tlbBarraHerramientas.Buttons(24).Visible = False
        .tlbBarraHerramientas.Buttons(25).Visible = False
        .cmdAnularReserva.Visible = False 'True
        .cmdReserva.Visible = False 'True
        .cmdConsultaSaldo.Visible = True
        .Show
    End With
    Recepciones = gstrBusca
End Function
Public Function Presupuestos(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    gstrIdCargo = 1
    With frmPresupuesto
        .Caption = "Consulta Presupuesto de Reparación"
        .lblCorrelativo.Caption = "Presupuesto Nº :"
        .lblEstadoOT.Visible = False
        .lblEstadoOTValor.Visible = False
        '.stbTotalMec.Visible = False
        '.stbTotalCarroceria.Visible = False
        '.stbTotalDesabolladura.Visible = False
        '.stbTotalPintura.Visible = False
        '.stbTotalOtros.Visible = False
        '.stbTotalTerceros.Visible = False
        '.stbTotalRepuestos.Visible = False
        '.stbTotalMateriales.Visible = False
        '.stbTotalOT.Visible = False
        .tlbBarraHerramientas.Buttons(13).Visible = False
        .tlbBarraHerramientas.Buttons(14).Visible = False
        .tlbBarraHerramientas.Buttons(15).Visible = False
        .tlbBarraHerramientas.Buttons(16).Visible = False
        .Show
    End With
    Presupuestos = gstrBusca
End Function

Public Function OrdenesdeTrabajo(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    With frmRecepcion
        .Caption = "Orden de Trabajo"
        .lblCorrelativo.Caption = "OT Nº :"
        .stbServicios.TabVisible(4) = True
        .stbServicios.TabVisible(5) = True
        .stbServicios.TabVisible(6) = True
        .stbTotalMec.Visible = True
        .stbTotalCarroceria.Visible = True
        .stbTotalDesabolladura.Visible = True
        .stbTotalPintura.Visible = True
        .stbTotalOtros.Visible = True
        .stbTotalTerceros.Visible = True
        .stbTotalRepuestos.Visible = True
        .stbTotalMateriales.Visible = True
        .stbTotalOT.Visible = True
        .lblEstadoOT.Visible = True
        .lblEstadoOTValor.Visible = True
        .tlbAddRep.Visible = True
        .cmdConsultaStock.Visible = False
        .tlbBarraHerramientas.Buttons(13).Visible = True
        .tlbBarraHerramientas.Buttons(14).Visible = True
        .tlbBarraHerramientas.Buttons(15).Visible = True
        .tlbBarraHerramientas.Buttons(16).Visible = True
        .cmdAnularReserva.Visible = False 'True
        .cmdReserva.Visible = False 'True
        .cmdConsultaSaldo.Visible = True
        .Show
    End With
    OrdenesdeTrabajo = gstrBusca
End Function
Public Function GeneraPresupuestos(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    gcurInsumo = 0
    With frmRecepcion
        .Caption = "Presupuesto de Reparación"
        .lblCorrelativo.Caption = "Presupuesto Nº :"
        .stbServicios.TabVisible(4) = True
        .stbServicios.TabVisible(5) = True
        .stbServicios.TabVisible(6) = True
        .stbTotalMec.Visible = True
        .stbTotalCarroceria.Visible = True
        .stbTotalDesabolladura.Visible = True
        .stbTotalPintura.Visible = True
        .stbTotalOtros.Visible = True
        .stbTotalTerceros.Visible = True
        .stbTotalRepuestos.Visible = True
        .stbTotalMateriales.Visible = True
        .stbTotalOT.Visible = True
        .tlbAgregarRepuestos.Visible = False
        .lblEstadoOT.Visible = True
        .lblEstadoOTValor.Visible = True
        .tlbBarraHerramientas.Buttons(13).Visible = True
        .tlbBarraHerramientas.Buttons(14).Visible = True
        .tlbBarraHerramientas.Buttons(15).Visible = True
        .tlbBarraHerramientas.Buttons(16).Visible = True
        .Label(30).Visible = False
        .txtFolioGarantia.Visible = False
        .Label2.Visible = False
        .lblDocumentos.Visible = False
        .Label5.Visible = False
        .txtNReferencia.Visible = False
        .Label(0).Visible = False
        .lblPresupuesto.Visible = False
        .Label3.Visible = False
        .lblFechaLiquidacion.Visible = False
        .tlbAddRep.Visible = True
        .cmdConsultaStock.Visible = True
        '.txtKilAct.Locked = True
        .cmdAnularReserva.Visible = False
        .cmdReserva.Visible = False
        .cmdConsultaSaldo.Visible = False
        .Show
    End With
    gstrProcedencia = "Presupuesto"
    GeneraPresupuestos = gstrBusca
End Function

Public Function ReservaDeHoras(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    With frmReservadeHoras
'        .Caption = "Orden de Trabajo"
'        .lblCorrelativo.Caption = "OT Nº :"
'        .stbServicios.TabVisible(4) = True
'        .stbServicios.TabVisible(5) = True
'        .stbServicios.TabVisible(6) = True
'        .stbTotalMec.Visible = True
'        .stbTotalCarroceria.Visible = True
'        .stbTotalDesabolladura.Visible = True
'        .stbTotalPintura.Visible = True
'        .stbTotalOtros.Visible = True
'        .stbTotalTerceros.Visible = True
'        .stbTotalRepuestos.Visible = True
'        .stbTotalMateriales.Visible = True
'        .stbTotalOT.Visible = True
'        .lblEstadoOT.Visible = True
'        .lblEstadoOTValor.Visible = True
'        .tlbBarraHerramientas.Buttons(13).Visible = True
'        .tlbBarraHerramientas.Buttons(14).Visible = True
'        .tlbBarraHerramientas.Buttons(15).Visible = True
'        .tlbBarraHerramientas.Buttons(16).Visible = True
        .Show
    End With
    ReservaDeHoras = gstrBusca
End Function
Public Function TipoGarantias(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
'    gstrProcedencia = "mnuTipoGarantias"
    frmMantenedorGarantias.Show vbModal
    TipoGarantias = gstrBusca
End Function


Public Function TipoCargo(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
'    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    frmMantenedorTipoCargo.Show vbModal
    TipoCargo = gstrBusca
End Function

Public Function OrdenesCompra(ByRef apConexion As APCONADO.ConnectionAdo, strId_Usuario As String, strPrefijoSistema As String, strCodigoAcceso As String, strCodigoEmpresa As String, strPathReporte As String, ByRef strCodigoInicial As String, Accion As apAccion) As String
'//Parametros estandar...
    Set Conexion = apConexion
    gstrIdUsuario = strId_Usuario
    gstrPrefijoSistema = strPrefijoSistema
    gstrCodigoAcceso = strCodigoAcceso
    gstrIdEmpresa = strCodigoEmpresa
    gstrPathReporte = strPathReporte
    gstrBusca = strCodigoInicial
    gapAccion = Accion
    frmEmisionOrdCom.Show
    OrdenesCompra = gstrBusca
End Function
