VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCopiaServiciosMarMod 
   Caption         =   "Copiar Servicios, Actividades, Repuestos"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      Begin VB.Frame Frame2 
         Caption         =   "Modelos"
         Height          =   4095
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5535
         Begin MSComctlLib.ListView lvwModelos 
            Height          =   3465
            Left            =   80
            TabIndex        =   8
            Top             =   240
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   6112
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
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "Codigo"
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "Des"
               Text            =   "Descripción"
               Object.Width           =   10019
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   3720
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   582
            ButtonWidth     =   1773
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Todos"
                  Key             =   "SelectAll"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   21
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Ninguno"
                  Key             =   "UnSelectAll"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   8
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImgBarraHerramienta 
            Left            =   3000
            Top             =   3600
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   22
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":0000
                  Key             =   "Crear"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":0112
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":056A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":09C2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":0E1A
                  Key             =   "Editar"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":0F2C
                  Key             =   "Grabar"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":103E
                  Key             =   "Cancelar"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1150
                  Key             =   "Borrar"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1262
                  Key             =   "Buscar"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1374
                  Key             =   "Imprimir"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1486
                  Key             =   "Cerrar"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1598
                  Key             =   "Ayuda"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":16AA
                  Key             =   "Primero"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":17BC
                  Key             =   "Anterior"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":18CE
                  Key             =   "Siguiente"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":19E0
                  Key             =   "Ultimo"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1AF2
                  Key             =   "Renovar"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1C04
                  Key             =   "SortAsc"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1D16
                  Key             =   "SortDesc"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":1E28
                  Key             =   "Seleccion"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":227A
                  Key             =   "Seleccion1"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCopiaServiciosMarMod.frx":26CC
                  Key             =   "Copiar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.CheckBox chkCopiarRepuestos 
         Caption         =   "Copiar Repuestos de las Actividades"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   5400
         Width           =   2895
      End
      Begin VB.CheckBox chkCopiaActividades 
         BackColor       =   &H8000000A&
         Caption         =   "Copiar Actividades del Servicio"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   5040
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo dtcMarca 
         Bindings        =   "frmCopiaServiciosMarMod.frx":27DE
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datMarcas 
         Height          =   330
         Left            =   1440
         Top             =   360
         Visible         =   0   'False
         Width           =   2970
         _ExtentX        =   5239
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
      Begin VB.Label Label1 
         Caption         =   "Marca      :"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   6000
      Width           =   975
   End
End
Attribute VB_Name = "frmCopiaServiciosMarMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoPrincipal As New ADODB.Recordset
Dim AdoModelo As New ADODB.Recordset
Dim lsiItem As ListItem
Dim mstrSql As String

Private Sub cmdAceptar_Click()
Dim intValida As Integer
Dim intContador As Integer
Dim i As Integer

Screen.MousePointer = vbHourglass

For intContador = 1 To lvwModelos.ListItems.Count
    Set lvwModelos.SelectedItem = lvwModelos.ListItems(intContador)
    If lvwModelos.ListItems(intContador).Checked = True Then
        intValida = Retorna_Valor_General("Select count(id_servicio) from Tllr_Servicio_Modelo Where Id_Marca='" & Me.dtcMarca.BoundText & "' And Id_Modelo='" & Me.lvwModelos.SelectedItem & "' And Id_Servicio='" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "'", gcdynamic)
        If intValida = 0 Then      'valida que no exista el servicio
            GrabaServicioModelo
        End If
        If Me.chkCopiaActividades.Value = 1 Then
            For i = 1 To frmTempServiciosMarMod.lvwActividades.ListItems.Count
                intValida = Retorna_Valor_General("Select count(id_Actividad) from Tllr_actividad_Servicio_Modelo Where Id_Marca='" & Me.dtcMarca.BoundText & "' And Id_Modelo='" & Me.lvwModelos.SelectedItem & "' And Id_Servicio='" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' And Id_Actividad='" & frmTempServiciosMarMod.lvwActividades.ListItems(i) & "'", gcdynamic)
                If intValida = 0 Then
                    GrabaActividadesServicio i
                    If Me.chkCopiarRepuestos.Value = 1 Then
                        GrabaRepuestosActividades i
                    End If
                End If
            Next i
        End If
    End If
Next
Screen.MousePointer = vbDefault
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    FillMarcas
End Sub
Sub FillMarcas()
    dtcMarca.Enabled = True
    mstrSql = "Select Id_marca as CODIGO, Descripcion as Nombre from Glbl_Marca where VIGENCIA = 'S' order by Descripcion"
    If Conexion.SendHost(mstrSql, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datMarcas
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcMarca.ListField = "Nombre"
                dtcMarca.BoundColumn = "Codigo"
                dtcMarca.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub

Sub FillModelos(strMarca As String)
        
Me.lvwModelos.ListItems.Clear

mstrSql = "Select Id_modelo as CODIGO, Descripcion as Nombre from Glbl_Modelo where VIGENCIA = 'S' and Id_marca = '" & strMarca & "'  order by Descripcion"
If Conexion.SendHost(mstrSql, AdoModelo, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoModelo
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set lsiItem = lvwModelos.ListItems.Add(, , !Codigo)
                lsiItem.SubItems(1) = !nombre
                .MoveNext
            Wend
        End If
    End With
End If
Conexion.CloseHost AdoModelo

End Sub

Private Sub dtcMarca_Change()
If dtcMarca.BoundText <> "" Then
    FillModelos dtcMarca.BoundText
End If
End Sub

Private Sub tlbOpciones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "SelectAll" '////Todos
        SelectingItem lvwModelos, gcSelectAll
    Case "UnSelectAll" '////Ninguno
        SelectingItem lvwModelos, gcUnSelectAll
    End Select
End Sub
Private Sub GrabaServicioModelo()
    mstrSql = "INSERT INTO TLLR_SERVICIO_MODELO ( Id_Marca, Id_Modelo, Id_Servicio, Valor, Horas ) "
    mstrSql = mstrSql & " VALUES( '" & Me.dtcMarca.BoundText & "' , "
    mstrSql = mstrSql & " '" & Me.lvwModelos.SelectedItem & "' , "
    mstrSql = mstrSql & " '" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' , " & CCur(Format(frmTempServiciosMarMod.lvwServicios.SelectedItem.SubItems(3), "####0")) & "," & CCur(frmTempServiciosMarMod.lvwServicios.SelectedItem.SubItems(2)) & ") "
    Conexion.SendHost mstrSql, , , , gcTiempoEspera
End Sub
Private Sub GrabaActividadesServicio(intIndice As Integer)
    mstrSql = "INSERT INTO Tllr_Actividad_Servicio_Modelo ( Id_Marca, Id_Modelo, Id_Servicio, ID_ACTIVIDAD, Horas, Valor ) "
    mstrSql = mstrSql & " VALUES( '" & Me.dtcMarca.BoundText & "' , "
    mstrSql = mstrSql & " '" & Me.lvwModelos.SelectedItem & "' , '" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' , "
    mstrSql = mstrSql & " '" & frmTempServiciosMarMod.lvwActividades.ListItems(intIndice) & "' ," & CCur(frmTempServiciosMarMod.lvwActividades.ListItems(intIndice).SubItems(2)) & ", " & CCur(Format(frmTempServiciosMarMod.lvwActividades.ListItems(intIndice).SubItems(3), "####0")) & ") "
    Conexion.SendHost mstrSql, , , , gcTiempoEspera
End Sub
Private Sub GrabaRepuestosActividades(intIndice As Integer)
Dim AdoTemp As New ADODB.Recordset

    'consulto los repuestos de la actividad
    mstrSql = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
    mstrSql = mstrSql & " Stck_Item.Descripcion AS NOMBRE, "
    mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
    mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Valor AS VLR, "
    mstrSql = mstrSql & " Stck_Item.Precio_Venta as Precio,"
    mstrSql = mstrSql & " Stck_Item.Id_Familia AS IDFAM, "
    mstrSql = mstrSql & " Glbl_Familia.Descripcion AS FAMILIA "
    mstrSql = mstrSql & " FROM Glbl_Familia RIGHT OUTER JOIN Stck_Item ON  Glbl_Familia.Id_Familia = Stck_Item.Id_Familia RIGHT OUTER JOIN Tllr_Actividad_Repuesto ON Stck_Item.Id_Item = Tllr_Actividad_Repuesto.Id_Item"
    mstrSql = mstrSql & " WHERE Tllr_Actividad_Repuesto.Id_Marca = '" & frmTempServiciosMarMod.dtcMarca.BoundText & "' AND Tllr_Actividad_Repuesto.Id_Modelo = '" & frmTempServiciosMarMod.dtcModelo.BoundText & "' AND Tllr_Actividad_Repuesto.Id_Servicio = '" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' AND Tllr_Actividad_Repuesto.Id_Actividad = '" & frmTempServiciosMarMod.lvwActividades.ListItems(intIndice) & "' "
        
    If Conexion.SendHost(mstrSql, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    mstrSql = "INSERT INTO Tllr_Actividad_Repuesto ( Id_Marca, Id_Modelo, Id_Servicio, Id_Actividad, Id_Item, Cantidad , Valor) "
                    mstrSql = mstrSql & " VALUES( '" & Me.dtcMarca.BoundText & "' , "
                    mstrSql = mstrSql & " '" & Me.lvwModelos.SelectedItem & "' , "
                    mstrSql = mstrSql & " '" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' , "
                    mstrSql = mstrSql & " '" & frmTempServiciosMarMod.lvwActividades.ListItems(intIndice) & "' , "
                    mstrSql = mstrSql & " '" & !Codigo & "' , "
                    mstrSql = mstrSql & " " & !CANTY & " , "
                    mstrSql = mstrSql & " " & CCur(Format(!VLR, "####0")) & ") "
                    
                    Conexion.SendHost mstrSql, , , , gcTiempoEspera
                    .MoveNext
                Wend
            End If
        End With
    End If
End Sub
