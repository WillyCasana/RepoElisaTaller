VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmMailing 
   Caption         =   "Mailing"
   ClientHeight    =   7575
   ClientLeft      =   -345
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMailing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Clientes"
      TabPicture(0)   =   "frmMailing.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFiltro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFechaIncorporacion(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFechaNacimiento"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Vehículo / Servicio"
      TabPicture(1)   =   "frmMailing.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Servicio"
         Height          =   4575
         Left            =   -68880
         TabIndex        =   80
         Top             =   360
         Width           =   5655
         Begin VB.CommandButton cmdBuscaServicio1 
            Height          =   330
            Left            =   5040
            Picture         =   "frmMailing.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   360
            Width           =   330
         End
         Begin VB.CommandButton cmdBuscaServicio2 
            Height          =   330
            Left            =   5040
            Picture         =   "frmMailing.frx":057C
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   960
            Width           =   330
         End
         Begin MSComctlLib.ListView lvwRevTecnica 
            Height          =   2655
            Left            =   360
            TabIndex        =   83
            Top             =   1800
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   4683
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
               Text            =   "Digito"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Mes"
               Object.Width           =   6879
            EndProperty
         End
         Begin VB.Label lblRevision2 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   320
            Left            =   1800
            TabIndex        =   100
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label lblRevision1 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   320
            Left            =   1800
            TabIndex        =   99
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label16 
            Caption         =   "Revisión Técnica"
            Height          =   255
            Left            =   360
            TabIndex        =   86
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "y no Ha Venido a :"
            Height          =   255
            Left            =   360
            TabIndex        =   85
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Asistio a  :"
            Height          =   255
            Left            =   360
            TabIndex        =   84
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Vehículo"
         Height          =   4575
         Left            =   -74880
         TabIndex        =   71
         Top             =   360
         Width           =   5775
         Begin MSComCtl2.DTPicker dtpCompradoHasta 
            Height          =   315
            Left            =   2760
            TabIndex        =   92
            Top             =   3810
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Format          =   82968577
            CurrentDate     =   37211
         End
         Begin MSComCtl2.DTPicker dtpCompradoDesde 
            Height          =   315
            Left            =   270
            TabIndex        =   91
            Top             =   3840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Format          =   82968577
            CurrentDate     =   37211
         End
         Begin MSComCtl2.DTPicker dtpVendidoHasta 
            Height          =   345
            Left            =   2700
            TabIndex        =   90
            Top             =   2970
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
            _Version        =   393216
            Format          =   82968577
            CurrentDate     =   37211
         End
         Begin MSComCtl2.DTPicker dtpVendidoDesde 
            Height          =   315
            Left            =   240
            TabIndex        =   89
            Top             =   3000
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            Format          =   82968577
            CurrentDate     =   37211
         End
         Begin MSDataListLib.DataCombo dtcCondicionVehiculo 
            Bindings        =   "frmMailing.frx":067E
            Height          =   315
            Left            =   3180
            TabIndex        =   87
            Top             =   1290
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nombre"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin VB.TextBox txtColorExterior 
            Height          =   315
            Left            =   3180
            TabIndex        =   81
            Top             =   2070
            Width           =   2445
         End
         Begin MSDataListLib.DataCombo dtcTipoVehiculo 
            Bindings        =   "frmMailing.frx":06A1
            Height          =   315
            Left            =   180
            TabIndex        =   74
            Top             =   1290
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nombre"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin VB.TextBox txtKilometrajeMayor 
            Height          =   315
            Left            =   1740
            TabIndex        =   73
            Top             =   2070
            Width           =   1215
         End
         Begin VB.TextBox txtañoFabricacion 
            Height          =   315
            Left            =   210
            TabIndex        =   72
            Top             =   2070
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo dtcMarca 
            Bindings        =   "frmMailing.frx":06BF
            Height          =   315
            Left            =   120
            TabIndex        =   97
            Top             =   480
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datMarcas 
            Height          =   330
            Left            =   120
            Top             =   495
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
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
         Begin MSDataListLib.DataCombo dtcModelo 
            Bindings        =   "frmMailing.frx":06D7
            Height          =   315
            Left            =   2700
            TabIndex        =   98
            Top             =   480
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datModelos 
            Height          =   330
            Left            =   2880
            Top             =   480
            Visible         =   0   'False
            Width           =   2010
            _ExtentX        =   3545
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
            Caption         =   "Adodc2"
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
         Begin MSAdodcLib.Adodc datTipoVehiculo 
            Height          =   330
            Left            =   720
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
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
            Caption         =   "adotipovehiculo"
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
         Begin MSAdodcLib.Adodc datCondicionVehiculo 
            Height          =   330
            Left            =   3480
            Top             =   1320
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
         Begin VB.Label Label19 
            Caption         =   "Comprado entre"
            Height          =   255
            Left            =   300
            TabIndex        =   94
            Top             =   3570
            Width           =   1395
         End
         Begin VB.Label Label18 
            Caption         =   "Vendido entre "
            Height          =   285
            Left            =   240
            TabIndex        =   93
            Top             =   2670
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Condición Vehiculo"
            Height          =   195
            Left            =   3210
            TabIndex        =   88
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label13 
            Caption         =   "Color"
            Height          =   255
            Left            =   3210
            TabIndex        =   82
            Top             =   1860
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo Vehículo"
            Height          =   255
            Left            =   180
            TabIndex        =   79
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Kilometraje > a"
            Height          =   255
            Left            =   1770
            TabIndex        =   78
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Año Fabricación"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   1830
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Modelo"
            Height          =   255
            Left            =   2700
            TabIndex        =   76
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Marca"
            Height          =   255
            Left            =   180
            TabIndex        =   75
            Top             =   270
            Width           =   555
         End
      End
      Begin VB.Frame fraFechaNacimiento 
         Caption         =   "Filtro por fecha de nacimiento"
         Height          =   915
         Left            =   6000
         TabIndex        =   64
         Top             =   4080
         Width           =   5790
         Begin MSComCtl2.DTPicker dtpFechaNacimientoDesde 
            Height          =   315
            Left            =   2880
            TabIndex        =   65
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   82968577
            CurrentDate     =   36715
         End
         Begin MSComCtl2.DTPicker dtpFechaNacimientoHasta 
            Height          =   315
            Left            =   4320
            TabIndex        =   66
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   82968577
            CurrentDate     =   36715
         End
         Begin VB.OptionButton optActivarfechanacimiento 
            Caption         =   "Activar"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   885
         End
         Begin VB.OptionButton optActivarfechanacimiento 
            Caption         =   "Desactivar"
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   69
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblHasta 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   0
            Left            =   4440
            TabIndex        =   68
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lblDesde 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   2880
            TabIndex        =   67
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame fraFechaIncorporacion 
         Caption         =   "Filtro por fecha de incorporación"
         Height          =   915
         Index           =   0
         Left            =   120
         TabIndex        =   57
         Top             =   4080
         Width           =   5850
         Begin VB.OptionButton optActivarFechaincorporacion 
            Caption         =   "Activar"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optActivarFechaincorporacion 
            Caption         =   "Desactivar"
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   58
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpFechaIncorporacionDesde 
            Height          =   315
            Left            =   2760
            TabIndex        =   60
            Top             =   480
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   82968577
            CurrentDate     =   36715
         End
         Begin MSComCtl2.DTPicker dtpFechaIncorporacionHasta 
            Height          =   315
            Left            =   4440
            TabIndex        =   61
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   82968577
            CurrentDate     =   36715
         End
         Begin VB.Label lblHasta 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   1
            Left            =   4440
            TabIndex        =   63
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lblDesde 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   62
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame fraFiltro 
         Height          =   3735
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   11655
         Begin VB.TextBox Text4 
            Height          =   300
            Left            =   4440
            TabIndex        =   95
            Top             =   3330
            Width           =   3345
         End
         Begin VB.ComboBox cboClasificacion 
            Height          =   315
            ItemData        =   "frmMailing.frx":06F0
            Left            =   7605
            List            =   "frmMailing.frx":06FD
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   945
            Width           =   2175
         End
         Begin VB.TextBox txtRut 
            Height          =   315
            Left            =   165
            TabIndex        =   23
            Top             =   360
            Width           =   2325
         End
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   2610
            TabIndex        =   22
            Top             =   360
            Width           =   4500
         End
         Begin VB.TextBox txtDireccion 
            Height          =   315
            Left            =   7200
            TabIndex        =   21
            Top             =   360
            Width           =   4305
         End
         Begin VB.TextBox txtTelefono 
            Height          =   315
            Left            =   2145
            TabIndex        =   20
            Top             =   1545
            Width           =   2325
         End
         Begin VB.TextBox txtNombreContacto 
            Height          =   315
            Left            =   4575
            TabIndex        =   19
            Top             =   1545
            Width           =   4305
         End
         Begin VB.TextBox txtGiro 
            Height          =   315
            Left            =   8880
            TabIndex        =   18
            Top             =   1545
            Width           =   2670
         End
         Begin VB.ComboBox cboEstadoCivil 
            Height          =   315
            ItemData        =   "frmMailing.frx":071C
            Left            =   3195
            List            =   "frmMailing.frx":072F
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2160
            Width           =   2460
         End
         Begin VB.TextBox txtHijos 
            Height          =   300
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   16
            Top             =   2805
            Width           =   2445
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Si"
            Height          =   195
            Index           =   0
            Left            =   2925
            TabIndex        =   15
            Top             =   3435
            Width           =   855
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   3555
            TabIndex        =   14
            Top             =   3435
            Width           =   855
         End
         Begin VB.ComboBox cboSexo 
            Height          =   315
            ItemData        =   "frmMailing.frx":0766
            Left            =   180
            List            =   "frmMailing.frx":0773
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1545
            Width           =   1890
         End
         Begin VB.TextBox txtPuestoNegocio 
            Height          =   300
            Left            =   8880
            MaxLength       =   50
            TabIndex        =   12
            Top             =   2805
            Width           =   2715
         End
         Begin VB.TextBox txtCompañia 
            Height          =   300
            Left            =   5205
            MaxLength       =   200
            TabIndex        =   11
            Top             =   2805
            Width           =   3630
         End
         Begin VB.TextBox txtTelefonoNegocio 
            Height          =   300
            Left            =   240
            MaxLength       =   50
            TabIndex        =   10
            Top             =   3345
            Width           =   2445
         End
         Begin VB.TextBox txtMaximo 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   10710
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   3225
            Width           =   480
         End
         Begin MSAdodcLib.Adodc datTipoCliente 
            Height          =   330
            Left            =   10200
            Top             =   945
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
         Begin MSAdodcLib.Adodc datPais 
            Height          =   330
            Left            =   600
            Top             =   975
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
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
         Begin MSAdodcLib.Adodc datCiudad 
            Height          =   330
            Left            =   3165
            Top             =   930
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
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
         Begin MSAdodcLib.Adodc datComuna 
            Height          =   330
            Left            =   5220
            Top             =   930
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
         Begin MSDataListLib.DataCombo dbcboComuna 
            Bindings        =   "frmMailing.frx":079D
            DataSource      =   "datComuna"
            Height          =   315
            Left            =   5130
            TabIndex        =   25
            Top             =   945
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Comuna"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcboPais 
            Bindings        =   "frmMailing.frx":07B5
            DataSource      =   "datPais"
            Height          =   315
            Left            =   180
            TabIndex        =   26
            Top             =   945
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Pais"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcboCiudad 
            Bindings        =   "frmMailing.frx":07CB
            DataSource      =   "datCiudad"
            Height          =   315
            Left            =   2655
            TabIndex        =   27
            Top             =   945
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Ciudad"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcboTipoCliente 
            Bindings        =   "frmMailing.frx":07E3
            DataSource      =   "datTipoCliente"
            Height          =   315
            Left            =   9840
            TabIndex        =   28
            Top             =   945
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Tipo_Cliente"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datMotivoVisita 
            Height          =   330
            Left            =   240
            Top             =   2160
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
         Begin MSDataListLib.DataCombo dbcboMotivoVisita 
            Bindings        =   "frmMailing.frx":0800
            Height          =   315
            Left            =   180
            TabIndex        =   29
            Top             =   2160
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Motivo_Visita"
            Text            =   "DataCombo1"
         End
         Begin MSAdodcLib.Adodc datDeporte 
            Height          =   330
            Left            =   5835
            Top             =   2160
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
         Begin MSDataListLib.DataCombo dbcboDeporte 
            Bindings        =   "frmMailing.frx":081E
            Height          =   315
            Left            =   5670
            TabIndex        =   30
            Top             =   2160
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Deporte"
            Text            =   "DataCombo1"
         End
         Begin MSAdodcLib.Adodc datSeguro 
            Height          =   330
            Left            =   8880
            Top             =   2160
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
         Begin MSDataListLib.DataCombo dbcboSeguro 
            Bindings        =   "frmMailing.frx":0837
            Height          =   315
            Left            =   8760
            TabIndex        =   31
            Top             =   2160
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Compañia"
            Text            =   "DataCombo1"
         End
         Begin MSAdodcLib.Adodc datEquipo 
            Height          =   330
            Left            =   240
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
         Begin MSDataListLib.DataCombo dbcboEquipo 
            Bindings        =   "frmMailing.frx":084F
            Height          =   315
            Left            =   180
            TabIndex        =   32
            Top             =   2805
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Equipo"
            Text            =   "DataCombo1"
         End
         Begin MSComCtl2.UpDown updMaximo 
            Height          =   315
            Left            =   11265
            TabIndex        =   33
            Top             =   3225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Value           =   10
            AutoBuddy       =   -1  'True
            BuddyControl    =   "fraFiltro"
            BuddyDispid     =   196635
            OrigLeft        =   11355
            OrigTop         =   3465
            OrigRight       =   11595
            OrigBottom      =   3780
            Max             =   100
            Min             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label20 
            Caption         =   "Motivo Visita"
            Height          =   165
            Left            =   4470
            TabIndex        =   96
            Top             =   3150
            Width           =   1095
         End
         Begin VB.Label lblClasificacion 
            Caption         =   "Clasificación"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7635
            TabIndex        =   56
            Top             =   705
            Width           =   1695
         End
         Begin VB.Label lblPais 
            AutoSize        =   -1  'True
            Caption         =   "País"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   210
            TabIndex        =   55
            Top             =   705
            Width           =   330
         End
         Begin VB.Label lblCiudad 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2685
            TabIndex        =   54
            Top             =   705
            Width           =   495
         End
         Begin VB.Label lblComuna 
            AutoSize        =   -1  'True
            Caption         =   "Comuna"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5130
            TabIndex        =   53
            Top             =   705
            Width           =   585
         End
         Begin VB.Label lblTipoCliente 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   9900
            TabIndex        =   52
            Top             =   705
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rut"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   210
            TabIndex        =   51
            Top             =   150
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2685
            TabIndex        =   50
            Top             =   150
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7230
            TabIndex        =   49
            Top             =   150
            Width           =   675
         End
         Begin VB.Label lblTelefono 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2175
            TabIndex        =   48
            Top             =   1335
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Contacto"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4620
            TabIndex        =   47
            Top             =   1335
            Width           =   1245
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Giro"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9090
            TabIndex        =   46
            Top             =   1335
            Width           =   285
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Visita por"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   45
            Top             =   1935
            Width           =   645
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil"
            Height          =   195
            Left            =   3240
            TabIndex        =   44
            Top             =   1935
            Width           =   825
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Deporte"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5700
            TabIndex        =   43
            Top             =   1935
            Width           =   570
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "N° Hijos"
            Height          =   195
            Left            =   2745
            TabIndex        =   42
            Top             =   2580
            Width           =   570
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Seguro"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8820
            TabIndex        =   41
            Top             =   1935
            Width           =   510
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Equipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   40
            Top             =   2580
            Width           =   495
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Señora Maneja:"
            Height          =   195
            Left            =   2940
            TabIndex        =   39
            Top             =   3180
            Width           =   1125
         End
         Begin VB.Label Label28 
            Caption         =   "Sexo"
            Height          =   285
            Left            =   195
            TabIndex        =   38
            Top             =   1335
            Width           =   600
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Puesto"
            Height          =   195
            Left            =   8895
            TabIndex        =   37
            Top             =   2580
            Width           =   495
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Compañía"
            Height          =   195
            Left            =   5205
            TabIndex        =   36
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono Compañía"
            Height          =   195
            Left            =   225
            TabIndex        =   35
            Top             =   3165
            Width           =   1410
         End
         Begin VB.Label Label7 
            Caption         =   "Máximo de Registros por Búsqueda"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   8100
            TabIndex        =   34
            Top             =   3240
            Width           =   2505
         End
      End
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Seleccionar Todos"
      Height          =   270
      Left            =   3600
      TabIndex        =   6
      Top             =   7320
      Width           =   2850
   End
   Begin MSComctlLib.Toolbar BarraHerramientas 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Nueva búsqueda"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar "
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir "
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mailing"
            Object.ToolTipText     =   "Mailing"
            ImageKey        =   "Mailing"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar "
            ImageKey        =   "Cerrar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwListaFiltro 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   3254
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   27
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rut"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Dirección"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "País"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ciudad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Comuna"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Clasificación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Sexo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Teléfono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Nombre Contacto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Giro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Visita por:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Estado Civil"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Deporte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Seguro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Equipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Hijos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Compañía"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Puesto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Fono Compañía"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Señora Maneja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Fecha Incorporación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Fecha Nacimiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Fax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "E-Mail"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Codigo Postal"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3090
      TabIndex        =   2
      Top             =   1410
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4215
      TabIndex        =   3
      Top             =   1410
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   0
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailing.frx":0867
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailing.frx":0979
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailing.frx":0A8B
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailing.frx":0B9D
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailing.frx":0CAF
            Key             =   "Mailing"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport rptListaClientes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   7320
      Width           =   1650
   End
   Begin VB.Label lblEncontrados 
      Height          =   240
      Left            =   1920
      TabIndex        =   4
      Top             =   7320
      Width           =   1110
   End
End
Attribute VB_Name = "frmMailing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset

Private Sub BarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            'Limpia
        Case "Buscar"
            Buscar
        'Case "Imprimir"
        
        Case "Mailing"
            'Generar
            
        Case "Cerrar"
            Unload Me
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdBuscaServicio1_Click()
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(Select Id_Servicio,Descripcion from Tllr_Servicio ) as MyTabla", "Id_Servicio", "Descripcion", Me.Caption)
    Me.lblRevision1 = Retorna_Valor_General("Select descripcion from Tllr_servicio where id_servicio='" & gstrBusca & "'", gcdynamic)
    Me.lblRevision1.Tag = gstrBusca
    
End Sub

Private Sub cmdBuscaServicio2_Click()
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(Select Id_Servicio,Descripcion from Tllr_Servicio ) as MyTabla", "Id_Servicio", "Descripcion", Me.Caption)
    Me.lblRevision2 = Retorna_Valor_General("Select descripcion from Tllr_servicio where id_servicio='" & gstrBusca & "'", gcdynamic)
    Me.lblRevision2.Tag = gstrBusca
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dbcboCiudad_Change()
    Me.dbcboComuna.Text = ""
    mstrSql = "SELECT * FROM Glbl_Comuna WHERE id_pais = '" & Me.dbcboPais.BoundText & "' And id_ciudad = '" & Me.dbcboCiudad.BoundText & "' ORDER BY Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datComuna.Recordset = adoPrincipal
    End If
End Sub

Private Sub dbcboPais_Change()
    Me.dbcboCiudad.Text = ""
    mstrSql = "SELECT * FROM Glbl_Ciudad WHERE id_pais = '" & Me.dbcboPais.BoundText & "' ORDER BY Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datCiudad.Recordset = adoPrincipal
    End If
End Sub

Private Sub dtcMarca_Change()
    Me.dtcModelo.Text = ""
    mstrSql = "Select Id_modelo as CODIGO, Descripcion as Nombre from Glbl_Modelo where VIGENCIA = 'S' and Id_marca = '" & Me.dtcMarca.BoundText & "'  order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datModelos.Recordset = adoPrincipal
    End If ' por el otro
End Sub

Private Sub Form_Activate()
    Dim blnBoolean As Boolean
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    Dim tbRegistros As New ADODB.Recordset
    Dim lstrQuery As String
    Dim itmAux As ListItem
    mblnSW = True
    '// Pais
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "SELECT * FROM glbl_Pais WHERE Vigencia = 'S' ORDER BY Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set datPais.Recordset = tbRegistros
    End If
        
    '// TipoCliente
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "SELECT * FROM glbl_Tipo_Cliente WHERE Vigencia = 'S' ORDER BY Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set datTipoCliente.Recordset = tbRegistros
    End If
    
    '// Motivo Visita
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "SELECT * FROM Glbl_Motivo_Visita_Cliente WHERE Vigencia = 'S' ORDER BY Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datMotivoVisita.Recordset = tbRegistros
    End If
    
    '// Deportes
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "SELECT * FROM Glbl_Deportes WHERE Vigencia = 'S' ORDER BY Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datDeporte.Recordset = tbRegistros
    End If
    
    '// Compania Seguro
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "SELECT * FROM Glbl_Compania_Seguro WHERE Vigencia = 'S' ORDER BY Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datSeguro.Recordset = tbRegistros
    End If
    
    '// Equipo Futbol
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "SELECT * FROM Glbl_Equipos_Futbol WHERE Vigencia = 'S' ORDER BY Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datEquipo.Recordset = tbRegistros
    End If
    
    Me.dtpFechaIncorporacionDesde.Value = Format(Date, "dd/mm/yyyy")
    Me.dtpFechaIncorporacionHasta.Value = Format(Date, "dd/mm/yyyy")
    Me.dtpFechaNacimientoDesde.Value = Format(Date, "dd/mm/yyyy")
    Me.dtpFechaNacimientoHasta.Value = Format(Date, "dd/mm/yyyy")
    txtMaximo.Text = updMaximo.Value
    
    'marca
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "Select Id_marca as CODIGO, Descripcion as Nombre from Glbl_Marca where VIGENCIA = 'S' order by Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datMarcas.Recordset = tbRegistros
    End If
    
    'tipo vehiculo
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "Select Id_TipoVehiculo as CODIGO, Descripcion as Nombre from Glbl_Tipo_Vehiculo order by Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datTipoVehiculo.Recordset = tbRegistros
    End If
    
    'tipo vehiculo
    Set tbRegistros = New ADODB.Recordset
    lstrQuery = "Select Id_Condicion_Vehiculo as CODIGO, Descripcion as Nombre from Glbl_Condicion_Vehiculo order by Descripcion"
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set Me.datCondicionVehiculo.Recordset = tbRegistros
    End If
    
    'Revision Tecnica
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "1")
    itmAux.SubItems(1) = "ABRIL"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "2")
    itmAux.SubItems(1) = "MAYO"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "3")
    itmAux.SubItems(1) = "JUNIO"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "4")
    itmAux.SubItems(1) = "JULIO"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "5")
    itmAux.SubItems(1) = "AGOSTO"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "6")
    itmAux.SubItems(1) = "SEPTIEMBRE"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "7")
    itmAux.SubItems(1) = "OCTUBRE"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "8")
    itmAux.SubItems(1) = "NOVIEMBRE"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "9")
    itmAux.SubItems(1) = "DICIEMBRE"
    Set itmAux = lvwRevTecnica.ListItems.Add(, , "0")
    itmAux.SubItems(1) = "ENERO"
End Sub

Private Sub optActivarFechaincorporacion_Click(Index As Integer)
    Select Case Index
    Case 0
        dtpFechaIncorporacionDesde.Enabled = True
        dtpFechaIncorporacionHasta.Enabled = True
    Case 1
        dtpFechaIncorporacionDesde.Enabled = False
        dtpFechaIncorporacionHasta.Enabled = False
    End Select
End Sub

Private Sub optActivarfechanacimiento_Click(Index As Integer)
    Select Case Index
    Case 0
        dtpFechaNacimientoDesde.Enabled = True
        dtpFechaNacimientoHasta.Enabled = True
    Case 1
        dtpFechaNacimientoDesde.Enabled = False
        dtpFechaNacimientoHasta.Enabled = False
    End Select
    
End Sub
Private Sub Buscar()
Dim mstrWhere As String

mstrWhere = "Where Id_Empresa='" & gstrIdEmpresa & "'"

If txtRut <> "" Then
    mstrWhere = " And Id_Cliente_Proveedor='" & txtRut & "'"
End If

If Me.txtNombre <> "" Then
    mstrWhere = " And Razon_Social='" & Me.txtNombre & "'"
End If

If Me.txtDireccion <> "" Then
    mstrWhere = " And Direccion='" & Me.txtDireccion & "'"
End If

If Me.dbcboPais.Text <> "" Then
    mstrWhere = " And Id_Pais='" & Me.dbcboPais.BoundText & "'"
End If

If Me.dbcboCiudad.Text <> "" Then
    mstrWhere = " And Id_Ciudad='" & Me.dbcboCiudad.BoundText & "'"
End If

If Me.dbcboComuna.Text <> "" Then
    mstrWhere = " And Id_Comuna='" & Me.dbcboComuna.BoundText & "'"
End If

If Me.cboClasificacion.Text <> "" Then
    mstrWhere = " And Id_Tipo_Cliente='" & Mid(Me.cboClasificacion.Text, 1, 1) & "'"
End If

If Me.dbcboTipoCliente.Text <> "" Then
    mstrWhere = " And Cliente_Proveedor='" & Me.dbcboTipoCliente.BoundText & "'"
End If

If Me.cboSexo <> "" Then
    mstrWhere = " And Sexo='" & Mid(Me.cboSexo.Text, 1, 1) & "'"
End If

If Me.txtTelefono <> "" Then
    mstrWhere = " And Telefono='" & Me.txtTelefono & "'"
End If

If Me.txtNombreContacto <> "" Then
    mstrWhere = " And NombreContacto='" & Me.txtNombreContacto & "'"
End If

If Me.txtGiro <> "" Then
    mstrWhere = " And Giro_Comercial='" & Me.txtGiro & "'"
End If

If Me.dbcboMotivoVisita.Text <> "" Then
    mstrWhere = " And Id_Motivo_Visita='" & Me.dbcboMotivoVisita.BoundText & "'"
End If

If Me.cboEstadoCivil <> "" Then
    mstrWhere = " And Estado_Civil='" & Mid(Me.cboEstadoCivil.Text, 1, 2) & "'"
End If

If Me.dbcboDeporte.Text <> "" Then
    mstrWhere = " And Id_Deporte='" & Me.dbcboDeporte.BoundText & "'"
End If

If Me.dbcboSeguro.Text <> "" Then
    mstrWhere = " And Id_Compania='" & Me.dbcboSeguro.BoundText & "'"
End If

If Me.dbcboEquipo.Text <> "" Then
    mstrWhere = " And Id_Equipo='" & Me.dbcboEquipo.BoundText & "'"
End If

If Me.txtHijos <> "" Then
    mstrWhere = " And Numeros_Hijos='" & Me.txtHijos & "'"
End If

If Me.txtCompañia <> "" Then
    mstrWhere = " And NombreTrabajo='" & Me.txtCompañia & "'"
End If




'sql para las revisiones
mstrSql = "SELECT Tllr_OT.Patente, Tllr_OT.Id_OT, Tllr_OT.Entrega_Estimada, Tllr_Servicio.Descripcion"
mstrSql = mstrSql & "FROM Tllr_OT INNER JOIN"
mstrSql = mstrSql & "Tllr_Mecanica_OT ON Tllr_OT.Id_Empresa = Tllr_Mecanica_OT.Id_Empresa AND"
mstrSql = mstrSql & "Tllr_OT.Id_Sucursal = Tllr_Mecanica_OT.Id_Sucursal AND Tllr_OT.Seccion_OT = Tllr_Mecanica_OT.Seccion_OT AND"
mstrSql = mstrSql & "Tllr_OT.Id_OT = Tllr_Mecanica_OT.Id_OT INNER JOIN"
mstrSql = mstrSql & "Tllr_Servicio ON Tllr_Mecanica_OT.Id_Servicio = Tllr_Servicio.Id_Servicio"
mstrSql = mstrSql & "WHERE (Tllr_Mecanica_OT.Id_Servicio = 'RV10001') AND (Tllr_OT.Estado <> 'P')"
mstrSql = mstrSql & "AND (Tllr_OT.Estado <> 'R')"
mstrSql = mstrSql & "AND Tllr_OT.patente not in(SELECT Tllr_OT.Patente FROM Tllr_OT INNER JOIN"
mstrSql = mstrSql & "Tllr_Mecanica_OT ON Tllr_OT.Id_Empresa = Tllr_Mecanica_OT.Id_Empresa AND"
mstrSql = mstrSql & "Tllr_OT.Id_Sucursal = Tllr_Mecanica_OT.Id_Sucursal AND Tllr_OT.Seccion_OT = Tllr_Mecanica_OT.Seccion_OT AND"
mstrSql = mstrSql & "Tllr_OT.Id_OT = Tllr_Mecanica_OT.Id_OT INNER JOIN"
mstrSql = mstrSql & "Tllr_Servicio ON Tllr_Mecanica_OT.Id_Servicio = Tllr_Servicio.Id_Servicio"

mstrSql = mstrSql & "WHERE (Tllr_Mecanica_OT.Id_Servicio = 'RV10002') AND (Tllr_OT.Estado <> 'P')"
mstrSql = mstrSql & "AND (Tllr_OT.Estado <> 'R'))"
mstrSql = mstrSql & "order by entrega_estimada"

'sql glbl_cliente
'SELECT Glbl_Cliente_Proveedor.Id_Cliente_Proveedor, Glbl_Comuna.Descripcion AS Comuna, Glbl_Ciudad.Descripcion AS Ciudad,
'                      Glbl_Pais.Descripcion AS Pais, Glbl_Cliente_Proveedor.Razon_Social, Glbl_Cliente_Proveedor.Direccion,
'                      Glbl_Cliente_Proveedor.Cliente_Proveedor, Glbl_Tipo_Cliente.Descripcion AS Clasificacion, Glbl_Cliente_Proveedor.Sexo,
'                      Glbl_Cliente_Proveedor.Telefono, Glbl_Cliente_Proveedor.NombreContacto, Glbl_Cliente_Proveedor.Giro_Comercial,
'                      GLBL_MOTIVO_VISITA_CLIENTE.Descripcion AS MotivoVisita, Glbl_Cliente_Proveedor.Estado_Civil,
'                      GLBL_DEPORTES.Descripcion AS Deporte, GLBL_COMPANIA_SEGURO.Descripcion AS CiaSeguro,
'                      GLBL_EQUIPOS_FUTBOL.Descripcion AS EquipoFutbol, Glbl_Cliente_Proveedor.Numero_Hijos, Glbl_Cliente_Proveedor.NombreTrabajo,
'                      Glbl_Cliente_Proveedor.PuestoTrabajo, Glbl_Cliente_Proveedor.TelefonoTrabajo, Glbl_Cliente_Proveedor.Señora_Maneja,
'                      Glbl_Cliente_Proveedor.Fecha_Incorporacion, Glbl_Cliente_Proveedor.Fecha_Nacimiento, Glbl_Cliente_Proveedor.Fax,
'                      Glbl_Cliente_Proveedor.E_Mail, Glbl_Cliente_Proveedor.CodigoPostal
'FROM         Glbl_Modelo INNER JOIN
'                      Tllr_Vehiculo_Cliente ON Glbl_Modelo.Id_Modelo = Tllr_Vehiculo_Cliente.Id_Modelo AND
'                      Glbl_Modelo.Id_Marca = Tllr_Vehiculo_Cliente.Id_Marca INNER JOIN
'                      Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca RIGHT OUTER JOIN
'                      Glbl_Cliente_Proveedor ON Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor LEFT OUTER JOIN
'                      GLBL_EQUIPOS_FUTBOL ON Glbl_Cliente_Proveedor.id_Equipo = GLBL_EQUIPOS_FUTBOL.id_Equipo LEFT OUTER JOIN
'                      GLBL_COMPANIA_SEGURO ON Glbl_Cliente_Proveedor.id_Compania = GLBL_COMPANIA_SEGURO.id_Compania LEFT OUTER JOIN
'                      GLBL_DEPORTES ON Glbl_Cliente_Proveedor.id_Deporte = GLBL_DEPORTES.id_Deporte LEFT OUTER JOIN
'                      Glbl_Tipo_Cliente ON Glbl_Cliente_Proveedor.Id_Tipo_Cliente = Glbl_Tipo_Cliente.Id_Tipo_Cliente AND
'                      Glbl_Cliente_Proveedor.Id_Tipo_Cliente = Glbl_Tipo_Cliente.Id_Tipo_Cliente LEFT OUTER JOIN
'                      GLBL_MOTIVO_VISITA_CLIENTE ON
'                      Glbl_Cliente_Proveedor.id_Motivo_Visita = GLBL_MOTIVO_VISITA_CLIENTE.id_Motivo_Visita LEFT OUTER JOIN
'                      Glbl_Pais ON Glbl_Cliente_Proveedor.Id_Pais = Glbl_Pais.Id_Pais LEFT OUTER JOIN
'                      Glbl_Ciudad ON Glbl_Cliente_Proveedor.Id_Ciudad = Glbl_Ciudad.Id_Ciudad AND
'                      Glbl_Cliente_Proveedor.Id_Pais = Glbl_Ciudad.Id_Pais LEFT OUTER JOIN
'                      Glbl_Comuna ON Glbl_Cliente_Proveedor.Id_Comuna = Glbl_Comuna.Id_Comuna AND
'                      Glbl_Cliente_Proveedor.Id_Comuna = Glbl_Comuna.Id_Comuna AND Glbl_Cliente_Proveedor.Id_Ciudad = Glbl_Comuna.Id_Ciudad AND
'                      Glbl_Cliente_Proveedor.Id_Pais = Glbl_Comuna.Id_Pais
'WHERE     (Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = '122622355')

End Sub
