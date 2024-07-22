VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmConsultadeHoras 
   Caption         =   "Consulta de Tareas Asignadas"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   Icon            =   "FrmConsultadeHoras.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   11025
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   10815
      Begin VB.CommandButton cmdLimpiar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5160
         Picture         =   "FrmConsultadeHoras.frx":179A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   315
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fechas"
         Height          =   1095
         Left            =   6720
         TabIndex        =   5
         Top             =   120
         Width           =   3615
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   345
            Left            =   1680
            TabIndex        =   6
            Top             =   645
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   83427329
            CurrentDate     =   37382
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   345
            Left            =   1680
            TabIndex        =   7
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   83427329
            CurrentDate     =   37382
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   720
            Width           =   615
         End
      End
      Begin MSAdodcLib.Adodc AdoMecanico 
         Height          =   330
         Left            =   2715
         Top             =   360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
      Begin MSDataListLib.DataCombo cmbMecanico 
         Bindings        =   "FrmConsultadeHoras.frx":189C
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "Id_mecanico"
         Text            =   "DataCombo2"
      End
      Begin VB.Label lblsucursal 
         AutoSize        =   -1  'True
         Caption         =   "Mecanico"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   705
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro (Ctrl+G)"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Registro (Ctrl+D)"
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
            Object.Visible         =   0   'False
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro (Ctrl+U)"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Cargar"
            Object.ToolTipText     =   "Cargar Archivo"
         EndProperty
      EndProperty
      Begin Crystal.CrystalReport crInforme 
         Left            =   6120
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin MSComctlLib.ListView lsvdetalle 
      Height          =   5580
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   9843
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro. O/T"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Mecanico"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Seccion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cod. Tarea"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tarea Descripcion"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Horas Asignadas"
         Object.Width           =   2540
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
            Picture         =   "FrmConsultadeHoras.frx":18B6
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":19C8
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":1ADA
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":1BEC
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":1CFE
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":1E10
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":1F22
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":2034
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":2146
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":2258
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":236A
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":247C
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":258E
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":26A0
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":27B2
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":28C4
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":29D6
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":2E28
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":327A
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":338C
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":34E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":3644
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":37A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":38FC
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":43C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":481C
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":4980
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":4DDC
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":4F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":6244
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":67E0
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":693C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":6A98
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":6DEC
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":7140
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":7494
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":77E8
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":7B3C
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":8080
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":8192
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":84E6
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":883A
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":894E
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":8CA2
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":8DB4
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultadeHoras.frx":9108
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmConsultadeHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoRecordMecanico As New ADODB.Recordset
Private Sub ImprimirConsulta()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String



    'Devuelve la ruta del directorio Windows
    Dim rc As Long
    Dim WinPath As String
    WinPath = Space$(300)
    rc = GetWindowsDirectory(WinPath, 300)
    GcamBaseTem = Trim$(WinPath)
    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
    '---------------------------------------
    
    If lsvdetalle.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE ( ot text, seccion text, mecanico text, cod_tarea text, descripcion text, horas text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lsvdetalle.ListItems.Count
        Tabla.AddNew
        Set lsvdetalle.SelectedItem = lsvdetalle.ListItems(i)
        Tabla!OT = IIf(lsvdetalle.ListItems(i) = "", " ", lsvdetalle.ListItems(i))
        Tabla!Mecanico = IIf(lsvdetalle.SelectedItem.SubItems(1) = "", " ", lsvdetalle.SelectedItem.SubItems(1))
        Tabla!Seccion = IIf(lsvdetalle.SelectedItem.SubItems(2) = "", " ", lsvdetalle.SelectedItem.SubItems(2))
        Tabla!cod_tarea = IIf(lsvdetalle.SelectedItem.SubItems(3) = "", " ", lsvdetalle.SelectedItem.SubItems(3))
        Tabla!Descripcion = IIf(lsvdetalle.SelectedItem.SubItems(4) = "", " ", lsvdetalle.SelectedItem.SubItems(4))
        Tabla!Horas = IIf(lsvdetalle.SelectedItem.SubItems(5) = "", " ", lsvdetalle.SelectedItem.SubItems(5))
      
        Tabla.Update
   Next i
   Tabla.Close
   
   With crInforme
        .ReportFileName = gstrPathReporte & "\HorasAsignadas.rpt"
        .WindowTitle = Me.Caption
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & USRID & "'"
        .Formulas(1) = "TITULO='CONSULTA DE HORAS ASIGNADAS'"
        .Formulas(2) = "Razonsocial='" & Retorna_Valor_General("Select Razon_social From Glbl_empresa Where Id_empresa ='" & gstrIdEmpresa & "'") & "'"
        .Formulas(3) = "Ruc='" & gstrIdEmpresa & "'"
        .Formulas(4) = "Direccion='" & Retorna_Valor_General("Select Direccion From Glbl_empresa Where Id_empresa ='" & gstrIdEmpresa & "'") & "'"
        .Formulas(5) = "Marcamodulo='ElisaTaller'"

        .Destination = crptToWindow
        .Action = True
   End With
   
   'D0bsnueva.Close
   Screen.MousePointer = 1
End Sub



Private Sub LLenalista()
    Dim PSQL As String
    Dim AdoTemp As New ADODB.Recordset
    Dim mstrWhere As String
    Dim mstrSql As String
    
    If Me.cmbMecanico.BoundText <> "" Then
        mstrWhere = mstrWhere & " and dbo.Tllr_Mecanica_OT.Mecanico_Designado = '" & Me.cmbMecanico.BoundText & "'"
    End If
    
'    mstrSql = "SELECT dbo.Tllr_Mecanica_OT.Id_OT, dbo.Tllr_Mecanica_OT.Mecanico_Designado, dbo.Tllr_Mecanica_OT.Id_Tarea, dbo.Tllr_Mecanica_OT.Horas, " _
'              & "dbo.Tllr_Mecanica_OT.Id_Servicio, dbo.Tllr_Servicio.Descripcion as servicio, dbo.Tllr_OT.Fecha_Emision, dbo.Tllr_OT.Estado, dbo.Tllr_Mecanica_OT.Seccion_OT, " _
'              & "dbo.Tllr_Mecanicos.Nombre FROM dbo.Tllr_Mecanica_OT INNER JOIN dbo.Tllr_OT ON dbo.Tllr_Mecanica_OT.Id_Empresa = dbo.Tllr_OT.Id_Empresa AND dbo.Tllr_Mecanica_OT.Id_Sucursal = dbo.Tllr_OT.Id_Sucursal AND " _
'              & "dbo.Tllr_Mecanica_OT.Id_OT = dbo.Tllr_OT.Id_OT AND dbo.Tllr_Mecanica_OT.Seccion_OT = dbo.Tllr_OT.Seccion_OT INNER JOIN dbo.Tllr_Servicio ON dbo.Tllr_Mecanica_OT.Id_Servicio = dbo.Tllr_Servicio.Id_Servicio INNER JOIN " _
'              & "dbo.Tllr_Mecanicos ON dbo.Tllr_Mecanica_OT.Id_Empresa = dbo.Tllr_Mecanicos.Id_Empresa AND dbo.Tllr_Mecanica_OT.Id_Sucursal = dbo.Tllr_Mecanicos.Id_Sucursal AND " _
'              & "dbo.Tllr_Mecanica_OT.Mecanico_Designado = dbo.Tllr_Mecanicos.Id_Mecanico " _
'//LREYES (vale...)

'              & "WHERE dbo.Tllr_OT.Estado = 'V' and dbo.Tllr_Mecanica_OT.id_empresa='" & gstrIdEmpresa & "' and dbo.Tllr_Mecanica_OT.id_sucursal='" & gstrIdSucursal & "' and dbo.Tllr_OT.Fecha_Emision between '" & Me.dtpFechaDesde & "' and '" & Me.dtpFechaHasta & "'" & mstrWhere & "order by Tllr_Mecanicos.Nombre"
    mstrSql = "exec Tllr_Tareas_Asignadas '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', '" & Format(Me.dtpFechaDesde.Value, "dd/mm/yyyy") & "', '" & Format(Me.dtpFechaHasta.Value, "dd/mm/yyyy") & "', '" & Me.cmbMecanico.BoundText & "'"
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) <> apOk Then
        Exit Sub
    End If
    
    Me.lsvdetalle.ListItems.Clear
    With AdoTemp
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            Dim Item As ListItem
            AdoTemp.MoveFirst
            While Not .EOF
                Set Item = Me.lsvdetalle.ListItems.Add(, , ValorNulo(AdoTemp!Id_OT))
                Item.SubItems(1) = ValorNulo(AdoTemp!Nombre)
                If ValorNulo(AdoTemp!Seccion_OT) = "M" Then
                   Item.SubItems(2) = "Mecanica"
                Else
                   Item.SubItems(2) = "Carroceria"
                End If
                Item.SubItems(3) = ValorNulo(AdoTemp!Id_tarea)
                Item.SubItems(4) = ValorNulo(AdoTemp!servicio)
                Item.SubItems(5) = ValorNulo(AdoTemp!Horas)
                
                .MoveNext
            Wend
        End If
    End With
     
    Conexion.CloseHost AdoTemp
   
End Sub

Private Sub LimpiaTodo()
    Me.lsvdetalle.ListItems.Clear
    Me.cmbMecanico.BoundText = ""
End Sub



Private Sub cmdLimpiar_Click()
    Me.cmbMecanico.BoundText = ""
End Sub

Private Sub Form_Activate()
    If Not Atributos("Glbl", "Tllr_30_0150", False, False, False, False) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
   'Llena el combo Mecanico
     mstrSql = "Select Id_mecanico, nombre From tllr_mecanicos Where Id_Empresa ='" + gstrIdEmpresa + "' Order By nombre"
     If Conexion.SendHost(mstrSql, AdoRecordMecanico, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        Set Me.AdoMecanico.Recordset = AdoRecordMecanico
     End If
     
     Me.dtpFechaDesde = BOM(Date)
     Me.dtpFechaHasta = EOM(Date)
 
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            LimpiaTodo
        Case "Buscar"
            LLenalista
        Case "Imprimir"
            ImprimirConsulta
        Case "Cerrar"
            Unload Me
        
    End Select
    Screen.MousePointer = vbDefault
End Sub


