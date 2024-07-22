VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmTareasasignadasMecanico 
   Caption         =   "Tareas Asignadas a Mecanicos"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "frmTareasAsignadasMecanico.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9360
      TabIndex        =   13
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   11295
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   5040
         TabIndex        =   19
         Top             =   120
         Width           =   6135
         Begin VB.CommandButton cmdLimpiar 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3840
            Picture         =   "frmTareasAsignadasMecanico.frx":179A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox txtTarea 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            TabIndex        =   20
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin MSAdodcLib.Adodc AdoMecanico 
            Height          =   330
            Left            =   1800
            Top             =   600
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
            Bindings        =   "frmTareasAsignadasMecanico.frx":189C
            Height          =   315
            Left            =   960
            TabIndex        =   22
            Top             =   600
            Width           =   2805
            _ExtentX        =   4948
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
            Caption         =   "Mecánico:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Tarea:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sección"
         Height          =   1335
         Left            =   9240
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton opcSeccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Ambas"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton opcSeccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Mecánica"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton opcSeccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Carrocería"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Estados 
         Caption         =   "Estado de Tareas"
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton opcEstadoTarea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opcEstadoTarea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Términado"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton opcEstadoTarea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Suspendido"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton opcEstadoTarea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Iniciado"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fechas"
         Height          =   1335
         Left            =   2160
         TabIndex        =   2
         Top             =   120
         Width           =   2775
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   345
            Left            =   840
            TabIndex        =   3
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   87162881
            CurrentDate     =   37382
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   345
            Left            =   840
            TabIndex        =   4
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   87162881
            CurrentDate     =   37382
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Termino:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   885
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   225
            Width           =   420
         End
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
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
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
            ImageKey        =   "Cancelar"
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
      Height          =   4860
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   8573
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
      NumItems        =   11
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
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tarea "
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Fecha Inicio "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Fecha Termino"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Servicio"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Hora Inicio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Hora Termino"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Total Horas"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Estado "
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
            Picture         =   "frmTareasAsignadasMecanico.frx":18B6
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":19C8
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":1ADA
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":1BEC
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":1CFE
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":1E10
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":1F22
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":2034
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":2146
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":2258
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":236A
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":247C
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":258E
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":26A0
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":27B2
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":28C4
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":29D6
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":2E28
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":327A
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":338C
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":34E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":3644
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":37A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":38FC
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":43C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":481C
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":4980
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":4DDC
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":4F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":6244
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":67E0
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":693C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":6A98
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":6DEC
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":7140
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":7494
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":77E8
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":7B3C
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":8080
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":8192
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":84E6
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":883A
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":894E
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":8CA2
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":8DB4
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareasAsignadasMecanico.frx":9108
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Total Horas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   14
      Top             =   7320
      Width           =   1095
   End
End
Attribute VB_Name = "frmTareasasignadasMecanico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
'    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
'    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE ( ot text, mecanico text, seccion text, cod_tarea text, fecha_inicio text, fecha_termino text, servicio text, hora_inicio text, hora_termino text,total_horas text, estado text )"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lsvdetalle.ListItems.Count
        Tabla.AddNew
        Set lsvdetalle.SelectedItem = lsvdetalle.ListItems(i)
        Tabla!OT = IIf(lsvdetalle.ListItems(i) = "", " ", lsvdetalle.ListItems(i))
        Tabla!Mecanico = IIf(lsvdetalle.SelectedItem.SubItems(1) = "", " ", lsvdetalle.SelectedItem.SubItems(1))
        Tabla!Seccion = IIf(lsvdetalle.SelectedItem.SubItems(2) = "", " ", lsvdetalle.SelectedItem.SubItems(2))
        Tabla!cod_tarea = IIf(lsvdetalle.SelectedItem.SubItems(3) = "", " ", lsvdetalle.SelectedItem.SubItems(3))
        Tabla!Fecha_Inicio = IIf(lsvdetalle.SelectedItem.SubItems(4) = "", " ", lsvdetalle.SelectedItem.SubItems(4))
        Tabla!Fecha_Termino = IIf(lsvdetalle.SelectedItem.SubItems(5) = "", " ", lsvdetalle.SelectedItem.SubItems(5))
        Tabla!servicio = IIf(lsvdetalle.SelectedItem.SubItems(6) = "", " ", lsvdetalle.SelectedItem.SubItems(6))
        Tabla!Hora_Inicio = IIf(lsvdetalle.SelectedItem.SubItems(7) = "", " ", lsvdetalle.SelectedItem.SubItems(7))
        Tabla!Hora_Termino = IIf(lsvdetalle.SelectedItem.SubItems(8) = "", " ", lsvdetalle.SelectedItem.SubItems(8))
        Tabla!Total_Horas = IIf(lsvdetalle.SelectedItem.SubItems(9) = "", " ", lsvdetalle.SelectedItem.SubItems(9))
        Tabla!estado = IIf(lsvdetalle.SelectedItem.SubItems(10) = "", " ", lsvdetalle.SelectedItem.SubItems(10))
        Tabla.Update
   Next i
   Tabla.Close
   Dbsnueva.Close
   
   With crInforme
        .ReportFileName = gstrPathReporte & "\Horasmecanicos.rpt"
        .WindowTitle = Me.Caption
'        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrIdUsuario & "'"
        .Formulas(1) = "TITULO='HORAS ASIGNADAS A MECANICOS'"
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
    '//LREYES...
    Dim adoTemp As New ADODB.Recordset
    Dim mstrSql As String
    Dim Total As Double
    
  
    Dim strSeccion As String
    Dim strEstadoTarea As String
    
    If Me.opcSeccion(0).Value Then
        strSeccion = "C"
    ElseIf Me.opcSeccion(1).Value Then
        strSeccion = "M"
    Else
        strSeccion = "T"
    End If

    If Me.opcEstadoTarea(0).Value Then
        strEstadoTarea = "I"
    ElseIf Me.opcEstadoTarea(1).Value Then
        strEstadoTarea = "S"
    ElseIf Me.opcEstadoTarea(2).Value Then
        strEstadoTarea = "T"
    Else
        strEstadoTarea = ""
    End If
    If Me.txtTarea = "" Then
        Me.txtTarea = "0"
    End If
    mstrSql = "exec Tllr_tareas_asignadas_proceso '" & gstrIdEmpresa & "', '" & Me.cmbMecanico.BoundText & "', '" & strSeccion & "', " & CDbl(Me.txtTarea) & ", '" & strEstadoTarea & "', '" & Format(Me.dtpFechaDesde.Value, "dd/mm/yyyy") & "', '" & Format(Me.dtpFechaHasta.Value, "dd/mm/yyyy") & "'"
    If Conexion.SendHost(mstrSql, adoTemp, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
    End If
    
    Total = 0
    Me.lsvdetalle.ListItems.Clear
    With adoTemp
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            Dim item As ListItem
            adoTemp.MoveFirst
            While Not .EOF
                Set item = Me.lsvdetalle.ListItems.Add(, , ValorNulo(adoTemp!Id_OT))
                item.SubItems(1) = ValorNulo(adoTemp!Nombre)
                If ValorNulo(adoTemp!Seccion_OT) = "M" Then
                   item.SubItems(2) = "Mecanica"
                Else
                   item.SubItems(2) = "Carroceria"
                End If
                item.SubItems(3) = ValorNulo(adoTemp!Id_tarea)
                item.SubItems(4) = ValorNulo(adoTemp!fech_inicio)
                item.SubItems(5) = ValorNulo(adoTemp!fech_termino)
                item.SubItems(6) = ValorNulo(adoTemp!servicio)
                item.SubItems(7) = Format(ValorNulo(adoTemp!Hora_Inicio), "HH:MM")
                item.SubItems(8) = Format(ValorNulo(adoTemp!Hora_Termino), "HH:MM")
                item.SubItems(9) = ValorNulo(adoTemp!Total_Horas)
                
                If ValorNulo(adoTemp!estado_tarea) = "T" Then
                    item.SubItems(10) = "TERMINADO"
                End If
                If ValorNulo(adoTemp!estado_tarea) = "I" Then
                    item.SubItems(10) = "INICIADO"
                End If
                If ValorNulo(adoTemp!estado_tarea) = "S" Then
                    item.SubItems(10) = "SUSPENDIDO"
                End If
                
                
                Total = Val(Total) + Val(ValorNulo(adoTemp!Total_Horas))
                Me.txtTotal = Val(Total)
                .MoveNext
            Wend
        End If
    End With
     
    Conexion.CloseHost adoTemp
   
End Sub



Private Sub cmdLimpiar_Click()
    Me.cmbMecanico.BoundText = ""
End Sub

Private Sub Form_Activate()
    If Not Atributos("Glbl", "Tllr_30_0140", False, False, False, False) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    'Llena el combo Mecanico
     Dim mstrSql As String
     mstrSql = "Select Id_mecanico, nombre From tllr_mecanicos Where Id_Empresa ='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "' Order By nombre"
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
Private Sub LimpiaTodo()
    Me.lsvdetalle.ListItems.Clear
    Me.cmbMecanico.BoundText = ""
End Sub

Private Sub txtTarea_GotFocus()
    MarcaTexto txtTarea
End Sub

Private Sub txtTarea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Exit Sub
    End If
    KeyAscii = 0
End Sub
