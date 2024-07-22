VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelTempActividades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Actividades"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmSelTempActividades.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwActividades 
      Height          =   3825
      Left            =   -15
      TabIndex        =   1
      Top             =   1815
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   6747
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
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción Actividad"
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
         Text            =   "Especialidad"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "IdEspec"
         Object.Width           =   18
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criterios de Busqueda"
      Height          =   1800
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   6750
      Begin VB.CheckBox optCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   315
         Width           =   990
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         TabIndex        =   6
         Top             =   285
         Width           =   2235
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Top             =   735
         Width           =   5130
      End
      Begin VB.ComboBox cboCoincidir 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSelTempActividades.frx":179A
         Left            =   4245
         List            =   "frmSelTempActividades.frx":17AA
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1185
         Width           =   2220
      End
      Begin MSDataListLib.DataCombo dtcEspecialidad 
         Bindings        =   "frmSelTempActividades.frx":17FD
         Height          =   315
         Left            =   1305
         TabIndex        =   2
         Top             =   1200
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datEspecialidad 
         Height          =   330
         Left            =   1830
         Top             =   1215
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
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   540
         Left            =   4380
         TabIndex        =   3
         Top             =   120
         Width           =   2340
         _ExtentX        =   4128
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
               ImageIndex      =   23
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
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Especialidad"
         Height          =   195
         Left            =   375
         TabIndex        =   8
         Top             =   1275
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coincidir en :"
         Height          =   195
         Left            =   3285
         TabIndex        =   7
         Top             =   1245
         Width           =   915
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   5715
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":181B
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":192D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":1D85
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":21DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2635
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2747
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2859
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":296B
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2A7D
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2B8F
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2CA1
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2DB3
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2EC5
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":2FD7
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":30E9
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":31FB
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":330D
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":341F
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":3531
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":3643
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":3A95
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":3EE7
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempActividades.frx":3FF9
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelTempActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSql As String
Dim strSino As String
Dim strWhere As String
Dim lsiItem As ListItem
Dim mblnSW As Boolean
Dim adoPrincipal As New ADODB.Recordset




Private Sub Form_Activate()
If mblnSW Then
    strSino = "%"
    FillEspecialidades
    mblnSW = False
End If

End Sub

Private Sub Form_Load()
mblnSW = True
End Sub

Sub FillActividades(strCondicion As String, strOrden As String)
lvwActividades.ListItems.Clear
mstrSql = "SELECT Tllr_Actividad.Id_Actividad AS CODIGO,"
mstrSql = mstrSql & " Tllr_Actividad.Descripcion AS NOMBRE,"
mstrSql = mstrSql & " Tllr_Actividad.Horas AS HORAS,"
mstrSql = mstrSql & " Tllr_Actividad.Valor AS VALOR,"
mstrSql = mstrSql & " Tllr_Especialidad.Descripcion AS ESPECIALIDAD,"
mstrSql = mstrSql & " Tllr_Actividad.Id_Especialidad AS IDESPEC"
mstrSql = mstrSql & " FROM Tllr_Actividad LEFT OUTER JOIN Tllr_Especialidad ON"
mstrSql = mstrSql & " Tllr_Actividad.Id_Especialidad = Tllr_Especialidad.Id_Especialidad"
mstrSql = mstrSql & strCondicion & strOrden

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwActividades.ListItems.Add(, , !Codigo)
                    lsiItem.SubItems(1) = !Nombre
                    lsiItem.SubItems(2) = !Horas
                    lsiItem.SubItems(3) = Format(!Valor, "###,##0")
                    lsiItem.SubItems(4) = !ESPECIALIDAD
                    lsiItem.SubItems(5) = !IDESPEC
                    .MoveNext
                Wend
            End If
        End With
    End If

End Sub

Sub FillEspecialidades()
    mstrSql = "SELECT Id_Especialidad AS Codigo, Descripcion AS Nombre FROM Tllr_Especialidad WHERE Vigencia = 'S'  order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datEspecialidad
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcEspecialidad.ListField = "Nombre"
                dtcEspecialidad.BoundColumn = "Codigo"
'                dtcEspecialidad.BoundText = .Recordset!CODIGO
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
    
    cboCoincidir.ListIndex = 0
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
        For intContador = 1 To lvwActividades.ListItems.Count
            Set lvwActividades.SelectedItem = lvwActividades.ListItems(intContador)
            If lvwActividades.ListItems(intContador).Checked = True Then
                Set itmFound = frmTempServiciosMarMod.lvwActividades.FindItem(lvwActividades.SelectedItem, lvwText, , 0)
                If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                    'MsgBox "No se ha encontrado ninguna coincidencia"
                    Set itmFound = frmTempServiciosMarMod.lvwActividades.ListItems.Add(, , lvwActividades.ListItems(intContador))
                    itmFound.SubItems(1) = lvwActividades.ListItems(intContador).SubItems(1)
                    itmFound.SubItems(2) = lvwActividades.ListItems(intContador).SubItems(2)
                    itmFound.SubItems(3) = lvwActividades.ListItems(intContador).SubItems(3)
                    itmFound.SubItems(4) = lvwActividades.ListItems(intContador).SubItems(4)
                    '/*//////////////////////////
                    mstrSql = "INSERT INTO Tllr_Actividad_Servicio_Modelo ( Id_Marca, Id_Modelo, Id_Servicio, ID_ACTIVIDAD, Horas, Valor ) "
                    mstrSql = mstrSql & " VALUES( '" & frmTempServiciosMarMod.dtcMarca.BoundText & "' , "
                    mstrSql = mstrSql & " '" & frmTempServiciosMarMod.dtcModelo.BoundText & "' , '" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' , "
                    mstrSql = mstrSql & " '" & lvwActividades.ListItems(intContador) & "' ," & lvwActividades.ListItems(intContador).SubItems(2) & ", " & CCur(Format(lvwActividades.ListItems(intContador).SubItems(3), "####0")) & ") "
                    Conexion.SendHost mstrSql, , , , gcTiempoEspera
                Else
                    MsgBox "Verifique... esta actividad ya existe"
                End If
            End If
        Next
    
    Case "Buscar"
        If optCriterios(0).Value = 1 Then '/////////////// codigo
            If dtcEspecialidad.BoundText <> "" Then
                strWhere = " Where id_Actividad LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' AND Tllr_Actividad.Id_Especialidad = '" & dtcEspecialidad.BoundText & "' "
            Else
                strWhere = " Where id_Actividad LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
            End If
            FillActividades strWhere, " Order by Id_Actividad"
        ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
            If dtcEspecialidad.BoundText <> "" Then
                strWhere = " Where Tllr_Actividad.Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' AND Tllr_Actividad.Id_Especialidad = '" & dtcEspecialidad.BoundText & "' "
            Else
                strWhere = " Where Tllr_Actividad.Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
            End If
            FillActividades strWhere, " Order by Tllr_Actividad.Descripcion"
        Else
            If dtcEspecialidad.BoundText <> "" Then
                strWhere = " Where Tllr_Actividad.Id_Especialidad = '" & dtcEspecialidad.BoundText & "' "
            Else
                strWhere = " "
            End If
            FillActividades strWhere, ""
        End If
    Case "Cerrar"
        Unload Me
        
End Select

End Sub
