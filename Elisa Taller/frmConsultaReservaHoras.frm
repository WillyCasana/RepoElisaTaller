VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConsultaReservaHoras 
   Caption         =   "Consulta de Horas Reservadas"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   Icon            =   "frmConsultaReservaHoras.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumCitas 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen de Horas"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   11415
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Horas Disponibles:"
         Height          =   195
         Left            =   7800
         TabIndex        =   14
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lblHorasDisponibles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   315
         Left            =   9480
         TabIndex        =   13
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Horas Reservadas:"
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblHorasReservadas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   315
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblTotalHoras 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   870
      End
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthColumns    =   4
      ScrollRate      =   1
      StartOfWeek     =   95027201
      CurrentDate     =   41590
      MinDate         =   -16799
   End
   Begin MSComctlLib.ListView lvwConsultaHoras 
      Height          =   3450
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6085
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Horas"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Teléfono"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Actividad"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Estado"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Marca"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Modelo"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   4800
      Top             =   5880
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
            Picture         =   "frmConsultaReservaHoras.frx":038A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":049C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":08F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":0D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":11A4
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":12B6
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":13C8
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":14DA
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":15EC
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":16FE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1810
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1922
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1A34
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1B46
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1C58
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1D6A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1E7C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":1F8E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":20A0
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":21B2
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":2604
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaReservaHoras.frx":2A56
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCliente 
      Height          =   330
      Left            =   5280
      TabIndex        =   1
      Top             =   6120
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Seleccionar"
            Object.ToolTipText     =   "Seleccionar"
            ImageKey        =   "Seleccion"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcRecepcionista 
      Bindings        =   "frmConsultaReservaHoras.frx":2B68
      Height          =   315
      Left            =   8760
      TabIndex        =   4
      Top             =   2760
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc datRecepcionista 
      Height          =   330
      Left            =   8760
      Top             =   2760
      Visible         =   0   'False
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
   Begin MSDataListLib.DataCombo dtcSucursal 
      Bindings        =   "frmConsultaReservaHoras.frx":2B87
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "NOMBRE"
      BoundColumn     =   "CODIGO"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc datSucursal 
      Height          =   330
      Left            =   5400
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "adodc1"
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
   Begin VB.Label Label2 
      Caption         =   "Sucursal"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Recepcionista"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblFechaLarga 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   3615
   End
End
Attribute VB_Name = "frmConsultaReservaHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmItem As ListItem
Dim SW As Boolean
Dim adoTemp As New ADODB.Recordset
Dim mstrSQL As String
Dim i As Integer
Private Sub dtcRecepcionista_Change()
    ActualizaHoraRecepcionista Me.MonthView1.Value
End Sub
Private Sub dtcSucursal_Change()
    CargaRecepcionista dtcRecepcionista, datRecepcionista, dtcSucursal.BoundText
    dtcRecepcionista.BoundText = frmReservadeHoras.txtRecepcionista
End Sub

Private Sub Form_Activate()
    Me.MonthView1.Value = frmReservadeHoras.pckFechaEntrega
    CargaSucursal
    dtcSucursal.BoundText = gstrIdSucursal
    'FillRecepcionista dtcRecepcionista, datRecepcionista
    'dtcRecepcionista.BoundText = frmReservadeHoras.txtRecepcionista
    LlenaListaConHoras gintHoraInicio, gintHoratermino, gintIntervaloMinutos
    LlenaHoras frmReservadeHoras.pckFechaEntrega
    'kjcv 03.07.14
    Me.dtcRecepcionista.BoundText = gstrIdMecanico
    '//lreyes...
    ActualizaTotales frmReservadeHoras.pckFechaEntrega.Value
End Sub

Private Sub Form_Load()
SW = True
Me.Label2.Caption = gstrNombreSucursal
End Sub
Sub LlenaListaConHoras(intHraIni As Integer, intHraFin As Integer, Intervalo As Integer)
Dim intHra As Integer, intMin As Integer
For intHra = intHraIni To intHraFin
    For intMin = 0 To 59 Step Intervalo
        Set itmItem = Me.lvwConsultaHoras.ListItems.Add(, , Format$(intHra, "00") & ":" & Format$(intMin, "00"))
        itmItem.SubItems(1) = ""
    Next
Next
End Sub

Sub LlenaHoras(dateFecha As Date)
    'LimpiaActividades
    lblFechaLarga = Format(dateFecha, "long date")
    ActualizaHoraRecepcionista dateFecha
End Sub

Private Sub lvwConsultaHoras_DblClick()
If Me.dtcRecepcionista <> "" Then
    frmReservadeHoras.pckFechaEntrega.Value = MonthView1.Value
'    frmReservadeHoras.cboHora.Text = lvwConsultaHoras.SelectedItem
'kjcv 13.11.13
    frmReservadeHoras.txtHora.Text = lvwConsultaHoras.SelectedItem
    frmReservadeHoras.txtRecepcionista = dtcRecepcionista.BoundText
    frmReservadeHoras.txtSucursal = dtcSucursal.BoundText
    Unload Me
Else
    MsgBox "El Recepcionista debe contener un Valor", vbExclamation, "Advertencia"
End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim NumCitas As Integer
    LlenaHoras DateClicked
    
    '//ActualizaTotales
    ActualizaTotales DateClicked
    'kjcv 12.08.14 Reporta Numero de Citas para Taller
    NumReservaHoras DateClicked
    If Val(txtNumCitas.Text) > 16 Then
        MsgBox "El numero de Reserva de Horas llego a su limite!!!", vbCritical, "Elisa Taller"
    End If
End Sub
Sub LimpiaActividades()

    For i = 1 To Me.lvwConsultaHoras.ListItems.Count
        lvwConsultaHoras.ListItems(i).SubItems(1) = ""
        lvwConsultaHoras.ListItems(i).SubItems(2) = ""
        lvwConsultaHoras.ListItems(i).SubItems(3) = ""
        lvwConsultaHoras.ListItems(i).SubItems(4) = ""
        lvwConsultaHoras.ListItems(i).SubItems(5) = ""
        lvwConsultaHoras.ListItems(i).SubItems(6) = ""
    Next
End Sub
Sub ActualizaHoraRecepcionista(Fecha As Date)

    LimpiaActividades
    
'    mstrSql = "SELECT Tllr_ReservaHora.Patente,"
'    mstrSql = mstrSql & "Tllr_ReservaHora.Fecha_Reserva, "
'    mstrSql = mstrSql & "Tllr_ReservaHora.Hora_Reserva, "
'    mstrSql = mstrSql & "Tllr_ReservaHora.Reparacion, "
'    mstrSql = mstrSql & "Glbl_Marca.Descripcion AS marca, "
'    mstrSql = mstrSql & "Glbl_Modelo.Descripcion AS modelo, "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor, "
'    mstrSql = mstrSql & "Glbl_Cliente_Proveedor.Razon_Social, "
'    mstrSql = mstrSql & "Glbl_Cliente_Proveedor.Telefono, "
'    mstrSql = mstrSql & "Tllr_ReservaHora.Estado "
'    mstrSql = mstrSql & "FROM Tllr_ReservaHora INNER JOIN "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente ON "
'    mstrSql = mstrSql & "Tllr_ReservaHora.Patente = Tllr_Vehiculo_Cliente.Patente INNER "
'    mstrSql = mstrSql & "Join "
'    mstrSql = mstrSql & "Glbl_Modelo INNER JOIN "
'    mstrSql = mstrSql & "Glbl_Marca ON "
'    mstrSql = mstrSql & "Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Marca.Id_Marca AND "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo INNER "
'    mstrSql = mstrSql & "Join "
'    mstrSql = mstrSql & "Glbl_Cliente_Proveedor ON "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
    
    mstrSQL = "SELECT Tllr_ReservaHora.Patente,Tllr_ReservaHora.Fecha_Reserva, "
    mstrSQL = mstrSQL & "Tllr_ReservaHora.Hora_Reserva, Tllr_ReservaHora.Reparacion, "
    mstrSQL = mstrSQL & "Glbl_Marca.Descripcion AS marca,Glbl_Modelo.Descripcion AS modelo, "
    mstrSQL = mstrSQL & "Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor,Glbl_Cliente_Proveedor.Razon_Social, "
    mstrSQL = mstrSQL & "Glbl_Cliente_Proveedor.Telefono, Tllr_ReservaHora.Estado, "
    mstrSQL = mstrSQL & "Tllr_ReservaHora.SinPatente, Tllr_ReservaHora.Nombre, "
    mstrSQL = mstrSQL & "Tllr_ReservaHora.Vehiculo,Tllr_ReservaHora.Telefono AS Fono "
    mstrSQL = mstrSQL & "FROM Glbl_Modelo INNER JOIN "
    mstrSQL = mstrSQL & "Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca INNER JOIN "
    mstrSQL = mstrSQL & "Tllr_Vehiculo_Cliente ON Glbl_Modelo.Id_Modelo = Tllr_Vehiculo_Cliente.Id_Modelo AND "
    mstrSQL = mstrSQL & "Glbl_Modelo.Id_Marca = Tllr_Vehiculo_Cliente.Id_Marca INNER JOIN "
    mstrSQL = mstrSQL & "Glbl_Cliente_Proveedor ON Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
    mstrSQL = mstrSQL & "RIGHT OUTER JOIN Tllr_ReservaHora ON Tllr_Vehiculo_Cliente.Patente = Tllr_ReservaHora.Patente "
    mstrSQL = mstrSQL & "WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & dtcSucursal.BoundText & "' And "
    mstrSQL = mstrSQL & "Recepcionista='" & dtcRecepcionista.BoundText & "' And "
    mstrSQL = mstrSQL & "Fecha_reserva Between '" & Fecha & "' AND '" & Fecha & "' And Tllr_ReservaHora.Estado <> 'R' And Tllr_ReservaHora.Estado <> 'E' "
    mstrSQL = mstrSQL & "ORDER BY Hora_reserva"
    If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        With adoTemp
           If Not .BOF And Not .EOF Then
              While Not .EOF
                For i = 1 To Me.lvwConsultaHoras.ListItems.Count
                    If lvwConsultaHoras.ListItems(i) = !Hora_Reserva Then
                        lvwConsultaHoras.ListItems(i).SubItems(1) = ValorNulo(IIf(!SinPatente = "S", !Nombre, !Razon_Social))
                        lvwConsultaHoras.ListItems(i).SubItems(2) = ValorNulo(IIf(!SinPatente = "S", !FONO, !Telefono))
                        lvwConsultaHoras.ListItems(i).SubItems(3) = ValorNulo(!Reparacion)
                        lvwConsultaHoras.ListItems(i).SubItems(4) = IIf(!estado = "V", "VIGENTE", IIf(!estado = "C", "CONFIRMADA", IIf(!estado = "N", "NULA", IIf(!estado = "E", "CANCELADA", ""))))
                        lvwConsultaHoras.ListItems(i).SubItems(5) = ValorNulo(IIf(!SinPatente = "S", !Vehiculo, !Marca))
                        lvwConsultaHoras.ListItems(i).SubItems(6) = ValorNulo(!Modelo)
                        Exit For
                    End If
                Next
              adoTemp.MoveNext
              Wend
           End If
        End With
    End If
    
End Sub
Sub CargaSucursal()
mstrSQL = "Select Id_Sucursal as Codigo, Descripcion as Nombre From Glbl_Sucursal "
mstrSQL = mstrSQL & "Where Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSQL, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datSucursal
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcSucursal.ListField = "Nombre"
            dtcSucursal.BoundColumn = "Codigo"
        End If
    End With
End If
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal

End Sub

Sub CargaRecepcionista(dtcObjeto As DataCombo, datObjeto As Adodc, pstrIdSucursal As String)
    gstrSql = "SELECT Id_Mecanico AS CODIGO, Nombre FROM Tllr_Mecanicos WHERE Es_Recepcionista = 'S' AND ID_EMPRESA='" & gstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "' and Vigencia='S' ORDER BY Nombre"
    dtcObjeto.Enabled = True
    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datObjeto
            Set .Recordset = gadoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcObjeto.ListField = "Nombre"
                dtcObjeto.BoundColumn = "Codigo"
            End If
        End With
    End If ' por el otro
    Set gadoPrincipal = New ADODB.Recordset
    Conexion.CloseHost gadoPrincipal
End Sub
Private Sub ActualizaTotales(dtfecha As Date)
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    strSql = "exec Tllr_Total_Agenda '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', '" & Format(dtfecha, "dd/mm/yyyy") & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            Me.lblTotalHoras = Format(adoTemp!Total_Horas, "#,##0.#0")
            Me.lblHorasReservadas = Format(adoTemp!Horas_Reserva, "#,##0.#0")
            Me.lblHorasDisponibles = Format(Round(adoTemp!Total_Horas - adoTemp!Horas_Reserva, 1), "#,##0.0")
        End If
    End If
    Conexion.CloseHost adoTemp
    Me.MousePointer = vbDefault
End Sub

Private Sub NumReservaHoras(dtfecha As Date)

Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT COUNT(id_reserva) AS NumCitas FROM Tllr_ReservaHora WHERE Id_Empresa='" & gstrIdEmpresa & "' "
strSql = strSql & " And Id_Sucursal='" & gstrIdSucursal & "' "
strSql = strSql & " and Estado ='V' "
strSql = strSql & " and Fecha_Reserva='" & Format(dtfecha, "dd/mm/yyyy") & "'"

If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        Me.txtNumCitas = recAux!NumCitas
    End If
End If
End Sub
