VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmEditaAsignacionOt 
   Caption         =   "Edicion Recurso Asignado"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   Icon            =   "frmEditaAsignacionOt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   3480
      Width           =   855
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3105
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   5477
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "item"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sucursal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "OT"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Seccion"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cod. Tarea"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Servicio"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Horas Asignadas"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Mecanico Asignado"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "idservicio"
         Object.Width           =   882
      EndProperty
   End
End
Attribute VB_Name = "frmEditaAsignacionOt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql As String
Dim adoTemp As New ADODB.Recordset
Dim lstrIdTarea As String
Dim lstrIdServicio As String
Dim lsiItem As ListItem

Private Sub cmdCancelar_Click()
frmAsignacionRecursos.Text1.Tag = ""
Unload Me
End Sub

Private Sub cmdQuitar_Click()
Dim dblHorasAsignadas As Double
Dim dblHorasDisponibles As Double
Dim intContador As Integer

For intContador = 1 To lvDetalle.ListItems.Count
    Set Me.lvDetalle.SelectedItem = Me.lvDetalle.ListItems(intContador)
    If Me.lvDetalle.ListItems(intContador).Checked = True Then
        'eliminar de hoja de recursos detalle
        strSql = "Delete from Tllr_Hoja_Recursos_Detalle where"
        strSql = strSql & " Id_empresa='" & gstrIdEmpresa & "'"
        strSql = strSql & " And Id_Sucursal='" & frmAsignacionRecursos.Text1 & "'"
        strSql = strSql & " And Id_Mecanico='" & frmAsignacionRecursos.Text4 & "'"
        strSql = strSql & " And Id_Turno='" & frmAsignacionRecursos.Text2 & "'"
        strSql = strSql & " And Id_Item='" & frmAsignacionRecursos.Text3 & "'"
        strSql = strSql & " And Id_Fecha='" & frmAsignacionRecursos.Text5 & "'"
        strSql = strSql & " And Id_Servicio='" & Me.lvDetalle.SelectedItem.SubItems(8) & "'"
        Conexion.SendHost strSql, , , , gcTiempoEspera
        
        'Actualizar hoja de recursos(encabezado)
        
        'rescato horas asignadas y disponibles
        strSql = "Select Horas_Asignadas,Horas_Disponibles from Tllr_Hoja_Recursos where id_empresa='" & gstrIdEmpresa & "'"
        strSql = strSql & " And Id_Sucursal='" & frmAsignacionRecursos.Text1 & "'"
        strSql = strSql & " And Id_Mecanico='" & frmAsignacionRecursos.Text4 & "'"
        strSql = strSql & " And Id_Turno='" & frmAsignacionRecursos.Text2 & "'"
        strSql = strSql & " And Id_Item='" & frmAsignacionRecursos.Text3 & "'"
        strSql = strSql & " And Id_Fecha='" & frmAsignacionRecursos.Text5 & "'"
        If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                dblHorasAsignadas = adoTemp!Horas_Asignadas - CDbl(Me.lvDetalle.SelectedItem.SubItems(6))
                dblHorasDisponibles = adoTemp!horas_disponibles + CDbl(Me.lvDetalle.SelectedItem.SubItems(6))
            End If
        End If
            
        strSql = "Update Tllr_Hoja_Recursos set Horas_Asignadas=" & dblHorasAsignadas & ", "
        strSql = strSql & "Horas_Disponibles=" & dblHorasDisponibles & " where"
        strSql = strSql & " Id_empresa='" & gstrIdEmpresa & "'"
        strSql = strSql & " And Id_Sucursal='" & frmAsignacionRecursos.Text1 & "'"
        strSql = strSql & " And Id_Mecanico='" & frmAsignacionRecursos.Text4 & "'"
        strSql = strSql & " And Id_Turno='" & frmAsignacionRecursos.Text2 & "'"
        strSql = strSql & " And Id_Item='" & frmAsignacionRecursos.Text3 & "'"
        strSql = strSql & " And Id_Fecha='" & frmAsignacionRecursos.Text5 & "'"
        Conexion.SendHost strSql, , , , gcTiempoEspera
        
        'actualizar Tllr_Mecanica_ot
        strSql = "Update Tllr_Mecanica_Ot set mecanico_designado='" & gstrMecanicoDefectoSecMec & "',"
        strSql = strSql & " estado_tarea='', HorasReales=0 where id_tarea='" & Me.lvDetalle.SelectedItem.SubItems(4) & "'"
        Conexion.SendHost strSql, , , , gcTiempoEspera
        
        'actualizar Tllr_otro_ot
        strSql = "Update Tllr_Otro_Ot set mecanico_Asignado='" & gstrMecanicoDefectoSecMec & "',"
        strSql = strSql & " estado_tarea='', HorasReales=0 where Id_Tarea='" & Me.lvDetalle.SelectedItem.SubItems(4) & "'"
        Conexion.SendHost strSql, , , , gcTiempoEspera
        
        'eliminar de horas reales
        strSql = "Delete from Tllr_Ordenes_Trabajo where id_tarea='" & Me.lvDetalle.SelectedItem.SubItems(4) & "'"
        Conexion.SendHost strSql, , , , gcTiempoEspera
        
        frmAsignacionRecursos.Text1.Tag = "QUITO"
    
    End If
Next
Unload Me
'eliminar de hoja de recursos detalle
'strSql = "Delete from Tllr_Hoja_Recursos_Detalle where"
'strSql = strSql & " Id_empresa='" & gstrIdEmpresa & "'"
'strSql = strSql & " And Id_Sucursal='" & frmAsignacionRecursos.Text1 & "'"
'strSql = strSql & " And Id_Mecanico='" & frmAsignacionRecursos.Text4 & "'"
'strSql = strSql & " And Id_Turno='" & frmAsignacionRecursos.Text2 & "'"
'strSql = strSql & " And Id_Item='" & frmAsignacionRecursos.Text3 & "'"
'strSql = strSql & " And Id_Fecha='" & frmAsignacionRecursos.Text5 & "'"
'strSql = strSql & " And Id_Servicio='" & lstrIdServicio & "'"
'Conexion.SendHost strSql, , , , gcTiempoEspera
'
''Actualizar hoja de recursos(encabezado)
'
''rescato horas asignadas y disponibles
'strSql = "Select Horas_Asignadas,Horas_Disponibles from Tllr_Hoja_Recursos where id_empresa='" & gstrIdEmpresa & "'"
'strSql = strSql & " And Id_Sucursal='" & frmAsignacionRecursos.Text1 & "'"
'strSql = strSql & " And Id_Mecanico='" & frmAsignacionRecursos.Text4 & "'"
'strSql = strSql & " And Id_Turno='" & frmAsignacionRecursos.Text2 & "'"
'strSql = strSql & " And Id_Item='" & frmAsignacionRecursos.Text3 & "'"
'strSql = strSql & " And Id_Fecha='" & frmAsignacionRecursos.Text5 & "'"
'If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
'    If Not adoTemp.BOF And Not adoTemp.EOF Then
'        dblHorasAsignadas = adoTemp!Horas_Asignadas - CDbl(Me.lblHorasAsignadas)
'        dblHorasDisponibles = adoTemp!horas_disponibles + CDbl(Me.lblHorasAsignadas)
'    End If
'End If
'
'strSql = "Update Tllr_Hoja_Recursos set Horas_Asignadas=" & dblHorasAsignadas & ", "
'strSql = strSql & "Horas_Disponibles=" & dblHorasDisponibles & " where"
'strSql = strSql & " Id_empresa='" & gstrIdEmpresa & "'"
'strSql = strSql & " And Id_Sucursal='" & frmAsignacionRecursos.Text1 & "'"
'strSql = strSql & " And Id_Mecanico='" & frmAsignacionRecursos.Text4 & "'"
'strSql = strSql & " And Id_Turno='" & frmAsignacionRecursos.Text2 & "'"
'strSql = strSql & " And Id_Item='" & frmAsignacionRecursos.Text3 & "'"
'strSql = strSql & " And Id_Fecha='" & frmAsignacionRecursos.Text5 & "'"
'Conexion.SendHost strSql, , , , gcTiempoEspera
'
''actualizar Tllr_Mecanica_ot
'strSql = "Update Tllr_Mecanica_Ot set mecanico_designado='" & gstrMecanicoDefectoSecMec & "',"
'strSql = strSql & " estado_tarea='', HorasReales=0 where id_tarea='" & lstrIdTarea & "'"
'Conexion.SendHost strSql, , , , gcTiempoEspera
'
''actualizar Tllr_otro_ot
'strSql = "Update Tllr_Otro_Ot set mecanico_Asignado='" & gstrMecanicoDefectoSecMec & "',"
'strSql = strSql & " estado_tarea='', HorasReales=0 where Id_Tarea='" & lstrIdTarea & "'"
'Conexion.SendHost strSql, , , , gcTiempoEspera
'
''eliminar de horas reales
'strSql = "Delete from Tllr_Ordenes_Trabajo where id_tarea='" & lstrIdTarea & "'"
'Conexion.SendHost strSql, , , , gcTiempoEspera
'
'frmAsignacionRecursos.Text1.Tag = "QUITO"
'
'Unload Me
End Sub

Private Sub Form_Load()
    CargaDatosOt
End Sub

Private Sub CargaDatosOt()
Dim ExisteMecanica As Boolean
Dim lstrestadoTarea As String
Dim AdoAux As New ADODB.Recordset
Dim i As Integer

i = 1

strSql = "Select * from Tllr_Hoja_Recursos_Detalle where"
strSql = strSql & " Id_Empresa='" & gstrIdEmpresa & "'"
strSql = strSql & " And Id_Sucursal='" & frmAsignacionRecursos.Text1 & "'"
strSql = strSql & " And Id_Mecanico='" & frmAsignacionRecursos.Text4 & "'"
strSql = strSql & " And Id_Turno='" & frmAsignacionRecursos.Text2 & "'"
strSql = strSql & " And Id_Item='" & frmAsignacionRecursos.Text3 & "'"
strSql = strSql & " And Id_Fecha='" & frmAsignacionRecursos.Text5 & "'"
If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoTemp.BOF And Not adoTemp.EOF Then
        While Not adoTemp.EOF
            lstrIdTarea = ValorNulo(adoTemp!Id_tarea)
            lstrIdServicio = ValorNulo(adoTemp!Id_servicio)
            ExisteMecanica = False
            
            strSql = "Select id_Sucursal,Id_ot,Seccion_Ot,Id_Tarea,Estado_Tarea,Id_Servicio,Horas,Mecanico_Designado from Tllr_Mecanica_OT Where id_Tarea='" & lstrIdTarea & "'"
            If Conexion.SendHost(strSql, AdoAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not AdoAux.BOF And Not AdoAux.EOF Then
                    Set lsiItem = Me.lvDetalle.ListItems.Add(, , i)
                    lsiItem.SubItems(1) = NombreSucursal(gstrIdEmpresa, AdoAux!Id_Sucursal)
                    lsiItem.SubItems(2) = AdoAux!Id_OT
                    lsiItem.SubItems(3) = IIf(AdoAux!Seccion_OT = "M", "MECANICA", "CARROCERIA")
                    lsiItem.SubItems(4) = AdoAux!Id_tarea
                    lsiItem.SubItems(5) = Retorna_Valor_General("Select descripcion from Tllr_Servicio where id_Servicio='" & AdoAux!Id_servicio & "'", gcdynamic)
                    lsiItem.SubItems(6) = AdoAux!Horas
                    lsiItem.SubItems(7) = TraeNombreMecanico(AdoAux!mecanico_designado)
                    lsiItem.SubItems(8) = lstrIdServicio
                    lstrestadoTarea = ValorNulo(AdoAux!estado_tarea)
                    ExisteMecanica = True
                    i = i + 1
                End If
            End If
            Conexion.CloseHost AdoAux
            
            If ExisteMecanica = False Then
                strSql = "Select id_Sucursal,Id_ot,Seccion_Ot,Id_Tarea,Estado_Tarea,Descripcion_otro,Horas,Mecanico_Asignado from Tllr_Otro_OT Where id_Tarea='" & lstrIdTarea & "'"
                If Conexion.SendHost(strSql, AdoAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                    If Not AdoAux.BOF And Not AdoAux.EOF Then
                        Set lsiItem = Me.lvDetalle.ListItems.Add(, , i)
                        lsiItem.SubItems(1) = NombreSucursal(gstrIdEmpresa, AdoAux!Id_Sucursal)
                        lsiItem.SubItems(2) = AdoAux!Id_OT
                        lsiItem.SubItems(3) = IIf(AdoAux!Seccion_OT = "M", "MECANICA", "CARROCERIA")
                        lsiItem.SubItems(4) = AdoAux!Id_tarea
                        lsiItem.SubItems(5) = AdoAux!Descripcion_Otro
                        lsiItem.SubItems(6) = AdoAux!Horas
                        lsiItem.SubItems(7) = TraeNombreMecanico(AdoAux!Mecanico_Asignado)
                        lsiItem.SubItems(8) = lstrIdServicio
                        lstrestadoTarea = ValorNulo(AdoAux!estado_tarea)
                        i = i + 1
                    End If
                End If
                Conexion.CloseHost AdoAux
            End If
            adoTemp.MoveNext
        Wend
    End If
End If
Conexion.CloseHost adoTemp
    
End Sub
