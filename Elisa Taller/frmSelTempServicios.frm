VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelTempServicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Servicios"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmSelTempServicios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Criterios de Busqueda"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox optCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   375
         Width           =   990
      End
      Begin VB.ComboBox cboCoincidir 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSelTempServicios.frx":038A
         Left            =   4245
         List            =   "frmSelTempServicios.frx":039A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1245
         Width           =   2220
      End
      Begin VB.ComboBox cboSeccion 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSelTempServicios.frx":03ED
         Left            =   1320
         List            =   "frmSelTempServicios.frx":03FA
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1245
         Width           =   1770
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   795
         Width           =   5130
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   345
         Width           =   2235
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   540
         Left            =   4080
         TabIndex        =   2
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   953
         ButtonWidth     =   1217
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "Agregar"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Cerrar"
               ImageIndex      =   27
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox optCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sección"
         Height          =   195
         Left            =   375
         TabIndex        =   8
         Top             =   1275
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coincidir en :"
         Height          =   195
         Left            =   3285
         TabIndex        =   7
         Top             =   1305
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView lvwServicios 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción Servicio"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Horas"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sección"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   6000
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":041B
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":052D
            Key             =   "Menos"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":0985
            Key             =   "Mas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":0DDD
            Key             =   "Persona"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1235
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1347
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1459
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":156B
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":167D
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":178F
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":18A1
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":19B3
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1AC5
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1BD7
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1CE9
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1DFB
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":1F0D
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":201F
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":2131
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":2243
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":2695
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":2AE7
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":2BF9
            Key             =   "Vaciar"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":304D
            Key             =   "Confirmar"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":3369
            Key             =   "LiquidarPres"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":37C1
            Key             =   "AnularPres"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempServicios.frx":3C15
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelTempServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSql As String
Dim strWhere As String
Dim lsiItem As ListItem
Dim adoPrincipal As New ADODB.Recordset


Sub FillServicios(strCondicion As String, strOrden As String)

lvwServicios.ListItems.Clear
mstrSql = "SELECT Id_Servicio, Descripcion, Horas, Valor, Seccion From Tllr_Servicio  " & strCondicion & strOrden
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwServicios.ListItems.Add(, , !Id_servicio)
                    lsiItem.SubItems(1) = !Descripcion
                    lsiItem.SubItems(2) = !Horas
                    lsiItem.SubItems(3) = Format(!Valor, "###,##0")
                    lsiItem.SubItems(4) = IIf(!Seccion = "M", "MECANICA", "CARROCERIA")
                    .MoveNext
                Wend
            End If
        End With
    End If

End Sub



Private Sub cboCoincidir_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then cboCoincidir.Text = ""
End Sub

Private Sub cboSeccion_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then cboSeccion.Text = ""
End Sub

Private Sub Form_Load()
cboSeccion.ListIndex = 2
cboCoincidir.ListIndex = 0

CargaServiciosMarca gstrServiciosMarca
End Sub

Private Sub optCriterios_Click(Index As Integer)
With Me
Select Case Index
    Case 0
        If .optCriterios(0).Value = 1 Then ' codigo
            .optCriterios(1).Value = 0
            .txtDes.Enabled = False
            .txtDes.Text = ""
            .txtCodigo.Enabled = True
            .txtCodigo.SetFocus
        Else
            .txtCodigo.Enabled = False
            .txtCodigo.Text = ""
        End If
    Case 1 '////////////////---------------descripcion
        If .optCriterios(1).Value = 1 Then
            .optCriterios(0).Value = 0
            .txtDes.Enabled = True
            .txtCodigo.Enabled = False
            .txtCodigo.Text = ""
            .txtDes.SetFocus
        Else
            .txtDes.Enabled = False
            .txtDes.Text = ""
        
        End If
End Select
End With
End Sub

Private Sub tlbBotones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim intContador As Integer
Dim itmFound As ListItem
Dim itmLista As ListItem

Select Case Button.Key
    Case "Agregar"
        For intContador = 1 To lvwServicios.ListItems.Count
            Set lvwServicios.SelectedItem = lvwServicios.ListItems(intContador)
            If lvwServicios.ListItems(intContador).Checked = True Then
                Set itmFound = frmTempServiciosMarMod.lvwServicios.FindItem(lvwServicios.SelectedItem, lvwText, , 0)
                If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
'                    MsgBox "No se ha encontrado ninguna coincidencia"
                    Set itmFound = frmTempServiciosMarMod.lvwServicios.ListItems.Add(, , lvwServicios.ListItems(intContador))
                    itmFound.SubItems(1) = lvwServicios.ListItems(intContador).SubItems(1)
                    itmFound.SubItems(2) = lvwServicios.ListItems(intContador).SubItems(2)
                    itmFound.SubItems(3) = lvwServicios.ListItems(intContador).SubItems(3)
                    itmFound.SubItems(4) = lvwServicios.ListItems(intContador).SubItems(4)
                    '/*//////////////////////////
                    mstrSql = "INSERT INTO TLLR_SERVICIO_MODELO ( Id_Marca, Id_Modelo, Id_Servicio, Valor, Horas ) "
                    mstrSql = mstrSql & " VALUES( '" & frmTempServiciosMarMod.dtcMarca.BoundText & "' , "
                    mstrSql = mstrSql & " '" & frmTempServiciosMarMod.dtcModelo.BoundText & "' , "
                    mstrSql = mstrSql & " '" & lvwServicios.ListItems(intContador) & "' , " & CCur(Format(lvwServicios.ListItems(intContador).SubItems(3), "####0")) & "," & CCur(lvwServicios.ListItems(intContador).SubItems(2)) & ") "
                    Conexion.SendHost mstrSql, , , , gcTiempoEspera
                End If
            End If
        Next
    Case "Buscar"
            strWhere = ""
            If optCriterios(0).Value = 1 Then '/////////////// codigo
                Select Case cboSeccion.Text
                    Case Is = "Carrocería"
                        strWhere = " Where id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' AND seccion = 'C' "
                    Case Is = "Mecánica"
                        strWhere = " Where id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' AND seccion = 'M' "
                    Case Is = "Ambas"
                        strWhere = " Where id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
                    Case Else
                        strWhere = " Where id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
                End Select
                    If gstrServiciosMarca = "S" Then
                        strWhere = strWhere & " and (id_marca = '" & frmTempServiciosMarMod.dtcMarca.BoundText & "')"
                    End If
                    FillServicios strWhere, " Order by Id_Servicio"
            ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
                Select Case cboSeccion.Text
                    Case Is = "Carrocería"
                        strWhere = " Where Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' AND seccion = 'C' "
                    Case Is = "Mecánica"
                        strWhere = " Where Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' AND seccion = 'M' "
                    Case Is = "Ambas"
                        strWhere = " Where Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
                    Case Else
                        strWhere = " Where Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
                End Select
                If gstrServiciosMarca = "S" Then
                    strWhere = strWhere & " and (id_marca = '" & frmTempServiciosMarMod.dtcMarca.BoundText & "')"
                End If
                FillServicios strWhere, " Order by Descripcion"
            Else
                If gstrServiciosMarca = "S" Then
                    If strWhere <> "" Then
                        strWhere = strWhere & " and id_marca = '" & frmTempServiciosMarMod.dtcMarca.BoundText & "'"
                    Else
                        strWhere = "where id_marca = '" & frmTempServiciosMarMod.dtcMarca.BoundText & "'"
                    End If
                End If

                FillServicios strWhere, "order by id_Servicio"
            End If
    Case "Cerrar"
        Unload Me
End Select
End Sub
Private Sub CargaServiciosMarca(ByRef strParametro As String)
    '//Verifica si utiliza servicios a nivel de marca...
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    
    strParametro = "N"
    strSql = "select isnull(ServiciosMarca,'N') as serviciosmarca from tllr_parametro where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            strParametro = IIf(UCase(adoTemp!ServiciosMarca) = "S", "S", "N")
        End If
    End If
    Conexion.CloseHost adoTemp
End Sub
