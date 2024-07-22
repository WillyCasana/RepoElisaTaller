VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmPresupuesto 
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   12135
   Icon            =   "frmPresupuesto.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   12135
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptOT 
      Left            =   105
      Top             =   7095
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
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   121
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear OT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar OT"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir OT"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Activar"
            Object.ToolTipText     =   "Activar OT"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anular"
            Object.ToolTipText     =   "Anular OT"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Liquidar"
            Object.ToolTipText     =   "Liquidar OT"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Renovar"
            Object.ToolTipText     =   "Refrescar Registros"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar"
            ImageKey        =   "Salir"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Confirmar"
            Object.ToolTipText     =   "Confirmar Reserva "
            ImageKey        =   "Confirmar"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Vaciar"
            Object.ToolTipText     =   "Eliminar Reserva"
            ImageKey        =   "Vaciar"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LiquidarPres"
            Object.ToolTipText     =   "Liquidar Presupuesto"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AnularPres"
            Object.ToolTipText     =   "Anular Presupuesto"
            ImageIndex      =   26
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame8 
      Caption         =   "Sección"
      Height          =   540
      Left            =   11040
      TabIndex        =   58
      Top             =   8520
      Visible         =   0   'False
      Width           =   2790
      Begin VB.OptionButton optRecepcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Carrocería"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1620
         TabIndex        =   60
         Tag             =   "Carrocería"
         Top             =   195
         Width           =   1110
      End
      Begin VB.OptionButton optRecepcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Mecánica"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   255
         TabIndex        =   59
         Tag             =   "Mecánica"
         Top             =   195
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Width           =   11655
      Begin VB.TextBox lblNroRecepcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   200
         Width           =   2100
      End
      Begin MSComCtl2.DTPicker pckFechaAtencion 
         Height          =   315
         Left            =   4950
         TabIndex        =   67
         Top             =   200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   143589377
         CurrentDate     =   36776
      End
      Begin VB.Label lblEstadoOTValor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7560
         TabIndex        =   72
         Top             =   200
         Width           =   1815
      End
      Begin VB.Label lblEstadoOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado OT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6540
         TabIndex        =   69
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Atención"
         Height          =   195
         Index           =   9
         Left            =   3660
         TabIndex        =   27
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label lblCorrelativo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Presupuesto Nº :"
         Height          =   195
         Left            =   105
         TabIndex        =   26
         Top             =   240
         Width           =   1200
      End
   End
   Begin TabDlg.SSTab stbServicios 
      Height          =   6135
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmPresupuesto.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeCia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fmePat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Inventario Recepción - Comentario"
      TabPicture(1)   =   "frmPresupuesto.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeInv"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fmeCom"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Servicios Mecánica"
      TabPicture(2)   =   "frmPresupuesto.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeMec"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "stbTotalMec"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Servicios Carroceria"
      TabPicture(3)   =   "frmPresupuesto.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "stbTotalPintura"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "stbTotalCarroceria"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "stbTotalArmeyDesarme"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "stbTotalDesabolladura"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fmeCar"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Otros Servicios"
      TabPicture(4)   =   "frmPresupuesto.frx":03FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "stbTotalOtros"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fmeOtr"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Servicios de Terceros"
      TabPicture(5)   =   "frmPresupuesto.frx":0416
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "stbTotalTerceros"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "fmeTer"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Repuestos"
      TabPicture(6)   =   "frmPresupuesto.frx":0432
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "stbInsumos"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "stbTotalMateriales"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "stbTotalRepuestos"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "fmeRep"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "StbLubricantes"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).ControlCount=   5
      Begin MSComctlLib.StatusBar StbLubricantes 
         Height          =   405
         Left            =   -73335
         TabIndex        =   127
         Top             =   5670
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Lubricantes"
               TextSave        =   "Total Lubricantes"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fmePat 
         Height          =   4275
         Left            =   45
         TabIndex        =   28
         Top             =   350
         Width           =   11700
         Begin VB.OptionButton optMantencion 
            Caption         =   "Mantención"
            Height          =   240
            Left            =   8160
            TabIndex        =   136
            Top             =   960
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.OptionButton optReparacion 
            Caption         =   "Reparación"
            Height          =   240
            Left            =   6600
            TabIndex        =   135
            Top             =   960
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtFolioGarantia 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5040
            MaxLength       =   30
            TabIndex        =   122
            Top             =   285
            Width           =   2595
         End
         Begin VB.TextBox txtFonos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   10365
            MaxLength       =   3
            TabIndex        =   113
            Top             =   2955
            Visible         =   0   'False
            Width           =   345
         End
         Begin MSComCtl2.DTPicker pckFecVta 
            Height          =   315
            Left            =   8070
            TabIndex        =   3
            Top             =   2505
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   143392769
            CurrentDate     =   36796
         End
         Begin VB.TextBox txtRut 
            Height          =   315
            Left            =   8955
            MaxLength       =   50
            TabIndex        =   111
            Top             =   4305
            Width           =   2085
         End
         Begin VB.TextBox txtComuna 
            Height          =   315
            Left            =   4710
            MaxLength       =   50
            TabIndex        =   110
            Top             =   4290
            Width           =   4185
         End
         Begin VB.TextBox txtDir 
            Height          =   315
            Left            =   435
            MaxLength       =   50
            TabIndex        =   109
            Top             =   4275
            Width           =   4185
         End
         Begin VB.TextBox txtConcesionario 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3615
            TabIndex        =   2
            Top             =   2520
            Width           =   3210
         End
         Begin VB.TextBox txtSolicita 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3870
            Width           =   4185
         End
         Begin VB.TextBox txtPatente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   915
            MaxLength       =   7
            TabIndex        =   0
            Top             =   990
            Width           =   1200
         End
         Begin VB.TextBox txtNroCono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   3
            TabIndex        =   5
            Top             =   3435
            Width           =   930
         End
         Begin VB.TextBox txtAño 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7635
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   29
            Top             =   1695
            Width           =   600
         End
         Begin VB.TextBox txtKilAct 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   930
            MaxLength       =   6
            TabIndex        =   1
            Top             =   2535
            Width           =   1380
         End
         Begin VB.ComboBox cboHora 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4245
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   3870
            Width           =   1170
         End
         Begin MSComCtl2.DTPicker pckFechaEntrega 
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Top             =   3885
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   143392769
            CurrentDate     =   36733
         End
         Begin MSDataListLib.DataCombo dtcTipoCono 
            Bindings        =   "frmPresupuesto.frx":044E
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   3405
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datTipoCono 
            Height          =   330
            Left            =   1740
            Top             =   3405
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
         Begin MSComctlLib.Toolbar tlbPatente 
            Height          =   330
            Left            =   2235
            TabIndex        =   30
            Top             =   960
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nuevo"
                  Object.ToolTipText     =   "Nuevo Patente"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Buscar"
                  Object.ToolTipText     =   "Buscar Patente"
                  ImageIndex      =   9
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcRecepcionista 
            Bindings        =   "frmPresupuesto.frx":0468
            Height          =   315
            Left            =   6945
            TabIndex        =   6
            Top             =   3420
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datRecepcionista 
            Height          =   330
            Left            =   9930
            Top             =   3420
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
         Begin MSComctlLib.ImageList ImgBarraHerramienta 
            Left            =   9915
            Top             =   810
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
                  Picture         =   "frmPresupuesto.frx":0487
                  Key             =   "Crear"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":0599
                  Key             =   "Menos"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":09F1
                  Key             =   "Mas"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":0E49
                  Key             =   "Persona"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":12A1
                  Key             =   "Editar"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":13B3
                  Key             =   "Grabar"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":14C5
                  Key             =   "Cancelar"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":15D7
                  Key             =   "Borrar"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":16E9
                  Key             =   "Buscar"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":17FB
                  Key             =   "Imprimir"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":190D
                  Key             =   "Cerrar"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":1A1F
                  Key             =   "Ayuda"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":1B31
                  Key             =   "Primero"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":1C43
                  Key             =   "Anterior"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":1D55
                  Key             =   "Siguiente"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":1E67
                  Key             =   "Ultimo"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":1F79
                  Key             =   "Renovar"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":208B
                  Key             =   "SortAsc"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":219D
                  Key             =   "SortDesc"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":22AF
                  Key             =   "Seleccion"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":2701
                  Key             =   "Seleccion1"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":2B53
                  Key             =   "Copiar"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":2C65
                  Key             =   "Vaciar"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":30B9
                  Key             =   "Confirmar"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":33D5
                  Key             =   "LiquidarPres"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":382D
                  Key             =   "AnularPres"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPresupuesto.frx":3C81
                  Key             =   "Salir"
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcGarantia 
            Bindings        =   "frmPresupuesto.frx":3FD3
            Height          =   315
            Left            =   900
            TabIndex        =   123
            Top             =   285
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datGarantia 
            Height          =   330
            Left            =   1950
            Top             =   -255
            Visible         =   0   'False
            Width           =   1980
            _ExtentX        =   3493
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
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero OT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   8280
            TabIndex        =   139
            Top             =   285
            Width           =   975
         End
         Begin VB.Label lblPresupuesto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9360
            TabIndex        =   138
            Top             =   285
            Width           =   1935
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kms. Act."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   33
            Left            =   105
            TabIndex        =   126
            Top             =   2610
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo OT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   125
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folio Gtía."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   30
            Left            =   4080
            TabIndex        =   124
            Top             =   285
            Width           =   915
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   195
            X2              =   11040
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   4
            X1              =   135
            X2              =   10980
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   135
            X2              =   10980
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Venta"
            Height          =   195
            Index           =   6
            Left            =   7050
            TabIndex        =   112
            Top             =   2595
            Width           =   915
         End
         Begin VB.Label lblMotor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4695
            TabIndex        =   108
            Top             =   2100
            Width           =   2790
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   210
            X2              =   11055
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fonos"
            Height          =   195
            Index           =   32
            Left            =   7545
            TabIndex        =   66
            Top             =   2985
            Width           =   435
         End
         Begin VB.Label lblFono 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8070
            TabIndex        =   65
            Top             =   2955
            Width           =   2250
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Solicita "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   31
            Left            =   6135
            TabIndex        =   64
            Top             =   3855
            Width           =   705
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VIN"
            Height          =   195
            Index           =   29
            Left            =   7650
            TabIndex        =   63
            Top             =   2115
            Width           =   270
         End
         Begin VB.Label lblVin 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8070
            TabIndex        =   62
            Top             =   2085
            Width           =   3510
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chasis"
            Height          =   195
            Index           =   22
            Left            =   135
            TabIndex        =   51
            Top             =   2130
            Width           =   465
         End
         Begin VB.Label lblChasis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   930
            TabIndex        =   50
            Top             =   2085
            Width           =   2850
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recepcionista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   16
            Left            =   5595
            TabIndex        =   47
            Top             =   3465
            Width           =   1230
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Cono"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   3420
            TabIndex        =   46
            Top             =   3435
            Width           =   720
         End
         Begin VB.Label lblCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   945
            TabIndex        =   45
            Top             =   2955
            Width           =   5880
         End
         Begin VB.Label lblColorE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8700
            TabIndex        =   44
            Top             =   1695
            Width           =   2880
         End
         Begin VB.Label lblModelo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3675
            TabIndex        =   43
            Top             =   1695
            Width           =   3540
         End
         Begin VB.Label lblMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   930
            TabIndex        =   42
            Top             =   1695
            Width           =   1980
         End
         Begin VB.Label lblPat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   1020
            Width           =   675
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   1695
            Width           =   450
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo"
            Height          =   195
            Index           =   2
            Left            =   3045
            TabIndex        =   39
            Top             =   1695
            Width           =   525
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   3
            Left            =   7290
            TabIndex        =   38
            Top             =   1695
            Width           =   285
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color Exterior"
            Height          =   195
            Index           =   5
            Left            =   8295
            TabIndex        =   37
            Top             =   1725
            Width           =   390
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   36
            Top             =   2955
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concesionario"
            Height          =   195
            Index           =   10
            Left            =   2475
            TabIndex        =   35
            Top             =   2535
            Width           =   1005
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cono"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   34
            Top             =   3420
            Width           =   450
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ent."
            Height          =   195
            Index           =   13
            Left            =   150
            TabIndex        =   33
            Top             =   3930
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Entrega"
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   14
            Left            =   3195
            TabIndex        =   32
            Top             =   3930
            Width           =   945
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Motor"
            Height          =   195
            Index           =   21
            Left            =   3900
            TabIndex        =   31
            Top             =   2100
            Width           =   705
         End
         Begin VB.Label lblIdMarca 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   930
            TabIndex        =   49
            Top             =   1695
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblIdModelo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5670
            TabIndex        =   48
            Top             =   1695
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblIdCliente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5820
            TabIndex        =   61
            Top             =   2970
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   5
            X1              =   135
            X2              =   10980
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   3
            X1              =   135
            X2              =   10980
            Y1              =   750
            Y2              =   750
         End
      End
      Begin VB.Frame fmeCar 
         Height          =   4905
         Left            =   -74950
         TabIndex        =   84
         Top             =   350
         Width           =   11700
         Begin VB.TextBox txtSeccion 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1995
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   92
            Top             =   405
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox txtValorDefCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4965
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   86
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtValorFinCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7455
            MaxLength       =   8
            TabIndex        =   89
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtPorcDesCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5955
            TabIndex        =   87
            Text            =   "00.0"
            Top             =   405
            Visible         =   0   'False
            Width           =   500
         End
         Begin VB.TextBox txtMtoDesCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6450
            MaxLength       =   8
            TabIndex        =   88
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtHorasCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   85
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   525
         End
         Begin MSDataListLib.DataCombo dtcCargoCar 
            Bindings        =   "frmPresupuesto.frx":3FED
            Height          =   315
            Left            =   8460
            TabIndex        =   90
            Top             =   405
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcMecanicoCar 
            Bindings        =   "frmPresupuesto.frx":4007
            Height          =   315
            Left            =   9645
            TabIndex        =   91
            Top             =   405
            Visible         =   0   'False
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvwServiciosCarroceria 
            Height          =   4140
            Left            =   45
            TabIndex        =   93
            Top             =   165
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   7303
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   19
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "CONCEPTO"
               Text            =   "Concepto"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "IDCONCEPTO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "DESCRIPCION"
               Text            =   "Descripción"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Key             =   "D_P"
               Text            =   "D/P/A"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "IDPARTEPIEZA"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   "HORAS"
               Text            =   "Cantidad"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   "VALORDEF"
               Text            =   "Precio Unitario"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Key             =   "PORCREC"
               Text            =   "% Recargo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Key             =   "MTOREC"
               Text            =   "(S/.) Recargo"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Key             =   "VALORFIN"
               Text            =   "Precio Venta"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Key             =   "PORCDESC"
               Text            =   "% Desc"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Key             =   "MTODESC"
               Text            =   "(S/.) Desc."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Key             =   "TIPOCARGO"
               Text            =   "Cargo"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Key             =   "IDCARGO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Key             =   "MECANICO"
               Text            =   "Proveedor"
               Object.Width           =   3492
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Key             =   "IDMEC"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   16
               Key             =   "TOTAL"
               Text            =   "Subtotal"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   17
               Key             =   "FACTURADO"
               Text            =   "Facturado"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   18
               Key             =   "ID_CODIGO"
               Text            =   "Codigo Servicio"
               Object.Width           =   0
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcPartePieza 
            Bindings        =   "frmPresupuesto.frx":4021
            Height          =   315
            Left            =   2370
            TabIndex        =   94
            Top             =   405
            Visible         =   0   'False
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcConceptos 
            Bindings        =   "frmPresupuesto.frx":403F
            Height          =   315
            Left            =   60
            TabIndex        =   95
            Top             =   405
            Visible         =   0   'False
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datPartesPiezas 
            Height          =   330
            Left            =   2400
            Top             =   435
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
         Begin MSAdodcLib.Adodc datConceptos 
            Height          =   330
            Left            =   150
            Top             =   420
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
         Begin MSAdodcLib.Adodc datCargoCar 
            Height          =   330
            Left            =   8160
            Top             =   390
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
         Begin MSAdodcLib.Adodc datMecanico 
            Height          =   330
            Left            =   10410
            Top             =   390
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
         Begin MSComctlLib.Toolbar tlbAddServicioCar 
            Height          =   330
            Left            =   135
            TabIndex        =   117
            Top             =   4455
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto"
            Height          =   195
            Index           =   24
            Left            =   720
            TabIndex        =   105
            Top             =   210
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   25
            Left            =   1995
            TabIndex        =   104
            Top             =   210
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parte / Pieza"
            Height          =   195
            Index           =   26
            Left            =   2925
            TabIndex        =   103
            Top             =   225
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$ Def."
            Height          =   195
            Index           =   27
            Left            =   5205
            TabIndex        =   102
            Top             =   210
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$ a Utilizar"
            Height          =   195
            Index           =   28
            Left            =   7530
            TabIndex        =   101
            Top             =   210
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$ Desc."
            Height          =   195
            Index           =   48
            Left            =   6675
            TabIndex        =   100
            Top             =   195
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% Desc."
            Height          =   195
            Index           =   49
            Left            =   5880
            TabIndex        =   99
            Top             =   195
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mecánico Asigado"
            Height          =   195
            Index           =   59
            Left            =   9930
            TabIndex        =   98
            Top             =   210
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Cargo"
            Height          =   195
            Index           =   60
            Left            =   8580
            TabIndex        =   97
            Top             =   195
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horas"
            Height          =   195
            Index           =   62
            Left            =   4485
            TabIndex        =   96
            Top             =   225
            Visible         =   0   'False
            Width           =   420
         End
      End
      Begin VB.Frame fmeOtr 
         Height          =   5250
         Left            =   -74950
         TabIndex        =   81
         Top             =   350
         Width           =   11700
         Begin MSComctlLib.ListView lvwOtrosServicios 
            Height          =   4400
            Left            =   50
            TabIndex        =   82
            Top             =   250
            Width           =   11600
            _ExtentX        =   20479
            _ExtentY        =   7752
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "CODIGO"
               Text            =   "Codigo Servicio"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "DESCRIPCION"
               Text            =   "Descripción"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   "NROHORAS"
               Text            =   "Nº de Horas"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   "PRECIOUNITARIO"
               Text            =   "(S/.) Unitario"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   "PORCDESC"
               Text            =   "% Descuento"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   "MTODESC"
               Text            =   "(S/.) Descuento"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   "IDCARGO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Key             =   "CARGO"
               Text            =   "Tipo Cargo"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Key             =   "IDMECANICO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Key             =   "MECANICO"
               Text            =   "Mecánico"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Key             =   "SUBTOTAL"
               Text            =   "Sub - Total"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Key             =   "FACTURADO"
               Text            =   "Facturado"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbAddServicioOtr 
            Height          =   330
            Left            =   105
            TabIndex        =   118
            Top             =   4770
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
      End
      Begin VB.Frame fmeTer 
         Height          =   5310
         Left            =   -74950
         TabIndex        =   77
         Top             =   350
         Width           =   11700
         Begin MSComctlLib.ListView lvwServiciosTerceros 
            Height          =   4400
            Left            =   50
            TabIndex        =   78
            Top             =   250
            Width           =   11600
            _ExtentX        =   20479
            _ExtentY        =   7752
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   16
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "ID_SERVICIO"
               Text            =   "Código Servicio"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Proveedor"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "IdProv"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "DESCRIPCION"
               Text            =   "Descripción"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   "NROFACTURA"
               Text            =   "Nº Factura"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   "VALOR"
               Text            =   "Precio Unitario"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   "CANTIDAD"
               Text            =   "Cantidad"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Key             =   "PORC_RECARGO"
               Text            =   "% Recargo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Key             =   "MTO_RECARGO"
               Text            =   "(S/.) Recargo"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Key             =   "PRECIO_FINAL"
               Text            =   "Precio Final"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Key             =   "PorcDescuento"
               Text            =   "% Dscto"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Key             =   "MONTODSCTO"
               Text            =   "(S/.) Dscto."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Key             =   "SUBTOTAL"
               Text            =   "Sub - Total"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Key             =   "TIPOCARGO"
               Text            =   "Tipo de Cargo"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Key             =   "ID_TIPOCARGO"
               Text            =   "IdCargo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Key             =   "FACTURADO"
               Text            =   "Facturado"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   2
            Left            =   0
            TabIndex        =   79
            Top             =   330
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbAddServicioTer 
            Height          =   330
            Left            =   90
            TabIndex        =   119
            Top             =   4755
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
      End
      Begin VB.Frame fmeRep 
         Height          =   4800
         Left            =   -74950
         TabIndex        =   73
         Top             =   350
         Width           =   11700
         Begin MSComctlLib.ListView lvwRepuestos 
            Height          =   4065
            Left            =   60
            TabIndex        =   74
            Top             =   240
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   7170
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "CODIGOITEM"
               Text            =   "Código Item"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "DESCRIPCION"
               Text            =   "Descripción"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   "CANTIDAD"
               Text            =   "Cantidad"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   "VALORUNITARIO"
               Text            =   "(S/.) Unitario"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   "PORCDESC"
               Text            =   "% Descuento"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   "MTODESC"
               Text            =   "(S/.) Descuento"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Key             =   "TIPOCARGO"
               Text            =   "Tipo Cargo"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Key             =   "IDCARGO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Key             =   "SUBTOTAL"
               Text            =   "Sub  - Total"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Key             =   "IDFAM"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Key             =   "FACTURADO"
               Text            =   "Facturado"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Key             =   "CONSUMO"
               Text            =   "Consumo"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   3
            Left            =   11805
            TabIndex        =   75
            Top             =   330
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
            ButtonWidth     =   1693
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbAddRep 
            Height          =   330
            Left            =   120
            TabIndex        =   120
            Top             =   4365
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
      End
      Begin MSComctlLib.StatusBar stbTotalDesabolladura 
         Height          =   405
         Left            =   -72375
         TabIndex        =   71
         Top             =   5280
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Desabolladura"
               TextSave        =   "Total Desabolladura"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalMec 
         Height          =   405
         Left            =   -68300
         TabIndex        =   68
         Top             =   5650
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Mecánica"
               TextSave        =   "Total Mecánica"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fmeMec 
         Height          =   5295
         Left            =   -74950
         TabIndex        =   56
         Top             =   350
         Width           =   11700
         Begin VB.CommandButton cmdAnularReserva 
            Appearance      =   0  'Flat
            Caption         =   "&Anular Reserva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   137
            Top             =   4800
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.CommandButton cmdReserva 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Reservar Repuestos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9450
            TabIndex        =   134
            Top             =   4785
            Visible         =   0   'False
            Width           =   2130
         End
         Begin MSComctlLib.Toolbar tlbAgregarRepuestos 
            Height          =   330
            Left            =   90
            TabIndex        =   131
            Top             =   4860
            Visible         =   0   'False
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin MSComctlLib.ListView lvwServiciosMecanica 
            Height          =   1740
            Left            =   45
            TabIndex        =   57
            Top             =   225
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   3069
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "CODIGO"
               Text            =   "Codigo Servicio"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "DESCRIPCION"
               Text            =   "Descripción"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   "NROHORAS"
               Text            =   "Nº de Horas"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   "PRECIOUNITARIO"
               Text            =   "(S/.) Unitario"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   "PORCDESC"
               Text            =   "% Descuento"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   "MTODESC"
               Text            =   "(S/.) Descuento"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   "IDCARGO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Key             =   "CARGO"
               Text            =   "Tipo Cargo"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Key             =   "IDMEC"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Key             =   "MECANICO"
               Text            =   "Mecánico"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Key             =   "SUBTOTAL"
               Text            =   "Sub - Total"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Key             =   "FACTURADO"
               Text            =   "Facturado"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Key             =   "IDNUEVO"
               Text            =   "NUEVO"
               Object.Width           =   882
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbAddServicioMec 
            Height          =   330
            Left            =   105
            TabIndex        =   116
            Top             =   2025
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin MSComctlLib.ListView lvwRepuestosMantencion 
            Height          =   1920
            Left            =   30
            TabIndex        =   133
            Top             =   2820
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   3387
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   10019
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Cantidad"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Valor"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Familia"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "IDFAM"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Tipo"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Repuestos Mantención"
            Height          =   285
            Left            =   105
            TabIndex        =   132
            Top             =   2565
            Width           =   1980
         End
      End
      Begin VB.Frame fmeCom 
         Caption         =   "Comentario"
         Height          =   5655
         Left            =   -70350
         TabIndex        =   54
         Top             =   350
         Width           =   6645
         Begin VB.TextBox txtComentario 
            Appearance      =   0  'Flat
            Height          =   5300
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   6330
         End
      End
      Begin VB.Frame fmeInv 
         Caption         =   "Inventario Recepciòn"
         Height          =   5655
         Left            =   -74950
         TabIndex        =   52
         Top             =   350
         Width           =   4425
         Begin MSComctlLib.ListView lvwInventario 
            Height          =   5300
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   9340
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   5433
            EndProperty
         End
      End
      Begin VB.Frame fmeCia 
         Height          =   1545
         Left            =   60
         TabIndex        =   11
         Top             =   4545
         Width           =   11700
         Begin VB.Frame Frame3 
            Caption         =   "Deducible"
            Height          =   675
            Left            =   150
            TabIndex        =   16
            Top             =   765
            Width           =   5400
            Begin VB.TextBox txtDeduciblePesos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3330
               MaxLength       =   8
               TabIndex        =   18
               Top             =   225
               Width           =   1920
            End
            Begin VB.TextBox txtDeducibleUF 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   720
               MaxLength       =   4
               TabIndex        =   17
               Top             =   240
               Width           =   1920
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pesos"
               Height          =   195
               Index           =   19
               Left            =   2730
               TabIndex        =   21
               Top             =   270
               Width           =   435
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dólares"
               Height          =   195
               Index           =   20
               Left            =   105
               TabIndex        =   19
               Top             =   270
               Width           =   540
            End
         End
         Begin VB.TextBox txtLiquidador 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   24
            Top             =   1125
            Width           =   4020
         End
         Begin VB.TextBox txtNroPoliza 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6960
            MaxLength       =   30
            TabIndex        =   22
            Top             =   750
            Width           =   2940
         End
         Begin VB.TextBox txtNroSiniestro 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6960
            MaxLength       =   30
            TabIndex        =   20
            Top             =   330
            Width           =   2925
         End
         Begin MSComctlLib.Toolbar tlbCiaSeg 
            Height          =   330
            Left            =   5085
            TabIndex        =   107
            Top             =   405
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nueva"
                  Object.ToolTipText     =   "Nueva Cia. Seguro"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Buscar"
                  Object.ToolTipText     =   "Buscar Cia. Seguro"
                  ImageIndex      =   9
               EndProperty
            EndProperty
         End
         Begin VB.Label lblCompañia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   150
            TabIndex        =   23
            Top             =   420
            Width           =   4890
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Siniestro"
            Height          =   195
            Index           =   18
            Left            =   5940
            TabIndex        =   15
            Top             =   405
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Poliza"
            Height          =   195
            Index           =   17
            Left            =   5925
            TabIndex        =   14
            Top             =   825
            Width           =   645
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liquidador"
            Height          =   195
            Index           =   15
            Left            =   5940
            TabIndex        =   13
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compañia de Seguro"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   12
            Top             =   225
            Width           =   1485
         End
      End
      Begin MSComctlLib.StatusBar stbTotalRepuestos 
         Height          =   405
         Left            =   -68310
         TabIndex        =   76
         Top             =   5655
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Repuestos"
               TextSave        =   "Total Repuestos"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalTerceros 
         Height          =   405
         Left            =   -68295
         TabIndex        =   80
         Top             =   5650
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total  Terceros"
               TextSave        =   "Total  Terceros"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalOtros 
         Height          =   405
         Left            =   -68300
         TabIndex        =   83
         Top             =   5655
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Otros Servicios"
               TextSave        =   "Total Otros Servicios"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalMateriales 
         Height          =   405
         Left            =   -68310
         TabIndex        =   106
         Top             =   5235
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Materiales"
               TextSave        =   "Total Materiales"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbInsumos 
         Height          =   405
         Left            =   -73320
         TabIndex        =   115
         Top             =   5235
         Visible         =   0   'False
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Insumos"
               TextSave        =   "Total Insumos"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalArmeyDesarme 
         Height          =   405
         Left            =   -68295
         TabIndex        =   128
         Top             =   5280
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Arme y Desarme"
               TextSave        =   "Total Arme y Desarme"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalCarroceria 
         Height          =   405
         Left            =   -68300
         TabIndex        =   129
         Top             =   5670
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Carrocería"
               TextSave        =   "Total Carrocería"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalPintura 
         Height          =   405
         Left            =   -72375
         TabIndex        =   130
         Top             =   5670
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Total Pintura"
               TextSave        =   "Total Pintura"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbTotalOT 
      Height          =   315
      Left            =   6960
      TabIndex        =   70
      Top             =   7320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Total OT"
            TextSave        =   "Total OT"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoPrincipal As New ADODB.Recordset
Dim mstrSQL As String
Dim mstrWhere As String
Dim mstrOrderBy As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mblnSW As Boolean
Dim itmAux As ListItem
Dim lsiItem As ListItem
Dim intIndice As Integer
Dim curValor As Currency
Dim mstrTipoCargo As String
Dim mstrIdOT As String
Dim mstrCargo As String
Dim mdblTotalInicial As Double
Dim mstrIdPresupuestoOrigen As String
Dim mblnBloqueo As Boolean
Dim dblTotalInicial As Double
Dim KilometrajeEntrada As Double 'Variable de ILeiva 07/02/2001 para conservar el kilometraje de entrada asi lo comparo con el que va a ingresar en la recepción debe ser mayor
Dim gstrEstadoMantencion As String
Dim gstrEstadoReparacion As String
Dim gstrEstadoDisponible As String
Dim gstrBuscaReserva As String
Dim NroRegularizacion As String
Dim gstrKmsAutoNuevo As String
Dim mstrEstadoPresupuesto As String
Dim mstrAgregaPresupuesto As Boolean
Dim mstrLiquidaPresupuesto As Boolean
Sub TipoOt(pstrTipoOt As String)
'lblPat.Caption
Select Case pstrTipoOt
Case "GFB"
    With Me
        .lblPat.Caption = gstrNombrePatente
        If Me.fmePat.Enabled = False Then
            fmePat.Enabled = True
        End If
        .txtFolioGarantia.Enabled = True
        .txtFolioGarantia.SetFocus
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "CS"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INA"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INR"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INS"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INU"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False ' True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INW"
    With Me
        .lblPat.Caption = "V.I.N."
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "NGN"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INC"
    With Me
        .lblPat.Caption = "V.I.N."
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "PEX"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "REN"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = True
        .optReparacion.Visible = True
        .tlbAddRep.Visible = False
        If InStr(gstrEmpresa, "SERINFO") Then
            .cmdAnularReserva.Visible = False 'True
            .cmdReserva.Visible = False 'True
        End If
        .tlbAgregarRepuestos.Visible = True
    End With
Case "PRE"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .tlbAddRep.Visible = True
        .cmdAnularReserva.Visible = False 'False
        .cmdReserva.Visible = False
        .tlbAgregarRepuestos.Visible = False
        mstrEstadoPresupuesto = "ON"
        mstrLiquidaPresupuesto = False
    End With
End Select
End Sub


Function ExistePatente(pstrPatente As String) As Boolean

mstrSQL = "Select top 1 * From Tllr_Vehiculo_Cliente"
mstrSQL = mstrSQL & " WHERE Tllr_Vehiculo_Cliente.Patente = '" & pstrPatente & "'"
If Conexion.SendHost(mstrSQL, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            ExistePatente = True
        Else
            ExistePatente = False
        End If
    End With
End If

End Function

Sub Bloqueo(pstrEstado As String)
If pstrEstado = "V" Or pstrEstado = "F" Or pstrEstado = "B" Or pstrEstado = "R" Or pstrEstado = "P" Then
    fmePat.Enabled = True
    fmeCia.Enabled = True
    fmeInv.Enabled = True
    fmeCom.Enabled = True
    mblnBloqueo = False
Else
    fmePat.Enabled = False
    fmeCia.Enabled = False
    fmeInv.Enabled = False
    fmeCom.Enabled = False
    mblnBloqueo = True
End If
End Sub

Function TotalSeccionCargo(pstrIdEmpresa As String, _
                            pstrIdSucursal As String, _
                            pstrIdOT As String, _
                            pstrIdTipoCargo As String, _
                            pstrTipoOt As String, _
                            Seccion As SumSec) As Currency
If pstrIdTipoCargo = "" Then
    If Seccion = ssMec Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_MECANICA_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssOtr Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM Tllr_OTRO_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssCar Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN  FROM TLLR_CARROCERIA_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssTer Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_TERCEROS_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssRep Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_REPUESTOS_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    End If
Else
    If Seccion = ssMec Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_MECANICA_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssOtr Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM Tllr_OTRO_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssCar Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN  FROM TLLR_CARROCERIA_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssTer Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_TERCEROS_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssRep Then
        gstrSql = "SELECT SUM(SUBTOTAL)  AS RESUMEN FROM TLLR_REPUESTOS_PRESUPUESTO"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    End If
End If
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            If Not IsNull(!Resumen) Then
                TotalSeccionCargo = !Resumen
            Else
                TotalSeccionCargo = 0
            End If
        End If
        .Close
    End With
End If

End Function
Function VerificaLubricantesTipoCargo(pstrIdEmpresa As String, _
                            pstrIdSucursal As String, _
                            pstrIdOT As String, _
                            pstrIdTipoCargo As String, _
                            pstrTipoOt As String, _
                            Seccion As SumSec) As Currency
                            
Dim SumaLubricantes As Currency
Dim SumaMateriales As Currency
                            
gstrSql = "SELECT Tllr_Repuestos_Presupuesto.SUBTOTAL,"
gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_PRESUPUESTO"
gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_PRESUPUESTO.ID_ITEM"
gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = " & gstrCodigoLubricantes ' 90"
        
SumaLubricantes = 0

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        While Not .EOF
            SumaLubricantes = SumaLubricantes + !SubTotal
            .MoveNext
        Wend
        .Close
    End With
End If
VerificaLubricantesTipoCargo = SumaLubricantes

'///// MATERIALES
gstrSql = "SELECT TLLR_REPUESTOS_PRESUPUESTO.SUBTOTAL,"
gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_PRESUPUESTO"
gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_PRESUPUESTO.ID_ITEM"
gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "'"
gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = " & gstrCodigoMateriales ' 85"

SumaLubricantes = 0

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        While Not .EOF
            SumaMateriales = SumaMateriales + !SubTotal
            .MoveNext
        Wend
        .Close
    End With
End If
gcurMateriales = SumaMateriales

End Function

Function AccesoEliminar(itmSeleccionado As ListItem) As Boolean
'If itmSeleccionado.SubItems(5) = "85" Then
    AccesoEliminar = True
'Else
'    AccesoEliminar = False
'End If
End Function

Sub PrintOT()
Dim mstrIdCargo As String
Dim mcurTNeto As Currency
Dim mcurTMec As Currency
Dim mcurTOtr As Currency
Dim mcurTCar As Currency
Dim mcurTTer As Currency
Dim mcurTRep As Currency
Dim mcurTMat As Currency
Dim mcurTIns As Currency
Dim mcurTLub As Currency
Dim mcurDeducible As Currency

On Error GoTo Solucion
    
mcurDeducible = CCur(Val(txtDeduciblePesos))
'gstrSql = "SELECT ID_TIPO_CARGO FROM TLLR_TIPO_CARGO"
'If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
'    With gadoPrincipal
'        If Not .BOF And Not .EOF Then
'            .MoveLast
'            While Not .BOF
'                mstrIdCargo = !Id_Tipo_Cargo
                mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssMec)
                mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssOtr)
                mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssCar)
                mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssTer)
                'MODIFICADO POR FDO DIAZ EL 04/01/2001
                mcurTLub = VerificaLubricantesTipoCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep)
                mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssRep) '- IIf(mstrIdCargo = "01", gcurMateriales, 0)
                mcurTIns = CalculoInsumos(8)
                'mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = "01", gcurMateriales, 0) + IIf(mstrIdCargo = "01", gcurInsumo, 0)
                mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep ' + IIf(mstrIdCargo = "01", gcurInsumo, 0)
                
                If mcurTNeto > 0 Then
                    With rptOT
                        
                        .ReportFileName = gstrPathReporte & "\PresupuestoConsulta.rpt"
                        '.ReportFileName = "C:\TallerSql\Reportes_C_Lub" & "\PresupuestoConsulta.rpt"
                        '.Destination = crptToPrinter
                        .Destination = crptToWindow
                        .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
                        .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
                        .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
                        .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
                        .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
                        .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
                        .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
                        
                        .Formulas(7) = "TMecanica=" & mcurTMec & ""
                        .Formulas(8) = "TOtros=" & mcurTOtr & ""
                        .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
                        .Formulas(10) = "TRepuesto=" & mcurTRep - mcurTLub - gcurMateriales - mcurTIns & ""
                        .Formulas(11) = "TDyP=" & mcurTCar & ""
                        .Formulas(12) = "TTerceros=" & mcurTTer & ""
                        
                        .Formulas(13) = "TMateriales=" & gcurMateriales + mcurTIns   '& IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
                        .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0) & ""
                        .Formulas(20) = "TLubricantes=" & mcurTLub & ""
                        
                        If mstrIdCargo = gstrCargoDeducibleMenos Then
                            If mcurDeducible <= mcurTNeto Then
                                .Formulas(15) = "Anexo= 'Deducible ( - )'"
                                .Formulas(16) = "TAnexo=" & mcurDeducible & ""
                                'mcurTNeto = mcurTNeto - mcurDeducible
                            End If
                        ElseIf mstrIdCargo = gstrCargoDeducibleMas Then
                                .Formulas(15) = "Anexo= 'Deducible ( + )'"
                                .Formulas(16) = "TAnexo=" & mcurDeducible & ""
                                'mcurTNeto = mcurTNeto + mcurDeducible
                        End If
                        .Formulas(17) = "TNetoOT=" & mcurTNeto & ""
                        .Formulas(18) = "IGV=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
                        .Formulas(19) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
                        .Formulas(21) = "NombreRuc='" & gstrNombreRut & "'"
                        .Formulas(22) = "NombrePlaca='" & gstrNombrePatente & "'"
                        .Formulas(23) = "NombreIgv='" & gstrNombreIva & "'"
                        .Connect = Conexion.ConnectionString
                        .Action = True
                    End With
                    '.MovePrevious
                Else
                    '.MovePrevious
                End If
                mcurTMec = 0
                mcurTOtr = 0
                mcurTCar = 0
                mcurTTer = 0
                mcurTRep = 0
                mcurTLub = 0
                mcurTNeto = 0
'            Wend
'        End If
'    End With
'End If
Solucion:
    If Err.Number <> 0 Then
        MsgBox "Impresión Cancelada por el usuario", vbExclamation, "Imprimir"
        Screen.MousePointer = 1
        Exit Sub
    End If

End Sub
Sub ImprimeCompletaSinDeducible()
Dim mstrIdCargo As String
Dim mcurTNeto As Currency
Dim mcurTMec As Currency
Dim mcurTOtr As Currency
Dim mcurTCar As Currency
Dim mcurTTer As Currency
Dim mcurTRep As Currency
Dim mcurTMat As Currency
Dim mcurTIns As Currency
Dim mcurDeducible As Currency
    mstrIdCargo = ""
    mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssMec)
    mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssOtr)
    mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssCar)
    mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssTer)
    mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep) - IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurMateriales, 0)
    mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurMateriales, 0) + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0)

    With rptOT
        .ReportFileName = gstrPathReporte & "\OTSTD" & ".rpt"
        .Destination = crptToPrinter
        .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
        .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
        .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
        .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
        .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
        .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
        
        .Formulas(7) = "TMecanica=" & mcurTMec & ""
        .Formulas(8) = "TOtros=" & mcurTOtr & ""
        .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
        .Formulas(10) = "TRepuesto=" & mcurTRep & ""
        .Formulas(11) = "TDyP=" & mcurTCar & ""
        .Formulas(12) = "TTerceros=" & mcurTTer & ""
        
        .Formulas(13) = "TMateriales=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurMateriales, 0) & ""
        .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0) & ""
        .Formulas(15) = "TNetoOT=" & mcurTNeto & ""
        .Formulas(16) = "IGV=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
        .Formulas(17) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
        .Action = True
    End With
End Sub


Sub EstadosOT(ModeAction As gAccionEstadoOT)
If ModeAction = gOTActivar Then
    '//////////////////////////////////////VERIFICAR
    If VeriLiq() = True Then
        gstrSql = "UPDATE Tllr_Presupuesto SET ESTADO = 'V' ,"
        gstrSql = gstrSql & "Fecha_Activacion = '" & CDate(pckFechaAtencion.Value) & "' , "
        gstrSql = gstrSql & "Quien_Activa = '" & gstrIdUsuario & "' "
        gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_Presupuesto.Id_OT = '" & lblNroRecepcion & "' AND Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' "
        If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
            lblEstadoOTValor = "VIGENTE"
            tlbBarraHerramientas.Buttons.Item(2).Enabled = True     'guardar
            tlbBarraHerramientas.Buttons.Item(13).Enabled = False   'ACTIVAR
            tlbBarraHerramientas.Buttons.Item(14).Enabled = True    'ANULAR
            tlbBarraHerramientas.Buttons.Item(15).Enabled = True    'LIQUIDAR
        End If
        EliminaRegistros gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, gstrSeccion
        MsgBox "La OT Nº " & lblNroRecepcion & " Fue Activada"
        Bloqueo "V"
    Else
        MsgBox "Lo siento, La Contraseña Ingresada no es la Correcta"
    End If
ElseIf ModeAction = gOTAnular Then
    If VeriLiq() = True Then
        gstrSql = "UPDATE Tllr_Presupuesto SET ESTADO = 'N' ,"
        gstrSql = gstrSql & "Fecha_Anulacion = '" & CDate(pckFechaAtencion.Value) & "' , "
        gstrSql = gstrSql & "Quien_Anula = '" & gstrIdUsuario & "' "
        gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_Presupuesto.Id_OT = '" & lblNroRecepcion & "' AND Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' "
        If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
            lblEstadoOTValor = "NULA"
            tlbBarraHerramientas.Buttons.Item(2).Enabled = False 'guardar
            tlbBarraHerramientas.Buttons.Item(13).Enabled = True    'ACTIVAR
            tlbBarraHerramientas.Buttons.Item(14).Enabled = False 'ANULAR
            tlbBarraHerramientas.Buttons.Item(15).Enabled = False  'LIQUIDAR
        End If
        MsgBox "La OT Nº " & lblNroRecepcion & " Fue Anulada"
    Else
        MsgBox "Lo siento, La Contraseña Ingresada no es la Correcta"
    End If
ElseIf ModeAction = gOTLiquidar Then
    
    frmLiquidacion.Show 1
    GrabarRegistro
    If VeriLiq() = True Then
        EliminaRegistros gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, gstrSeccion
        gstrSql = "UPDATE Tllr_Presupuesto SET ESTADO = 'L' ,"
        gstrSql = gstrSql & "Fecha_Liquidacion = '" & CDate(Format(Now, "dd/mm/yyyy")) & "' , "
        gstrSql = gstrSql & "Quien_Liquida = '" & gstrIdUsuario & "' ,"
        gstrSql = gstrSql & "Total_Insumos=" & gcurInsumo & " ,"
        gstrSql = gstrSql & "Total_Materiales=" & gcurMateriales & " ,"
        gstrSql = gstrSql & "Total_Iva=" & Round(gcurTotalIVA, gintDecimalesMoneda) & " ,"
        gstrSql = gstrSql & "Total_OT_IVA=" & Round(gcurTotalNetoMasIVA, gintDecimalesMoneda) & " ,"
        gstrSql = gstrSql & "Total_OT=" & Round(gcurTotalNeto, gintDecimalesMoneda) & " "
        gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_Presupuesto.Id_OT = '" & lblNroRecepcion & "' AND Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' "
        If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
            lblEstadoOTValor = "LIQUIDADA"
            tlbBarraHerramientas.Buttons.Item(2).Enabled = False 'guardar
            tlbBarraHerramientas.Buttons.Item(13).Enabled = True 'ACTIVAR
            tlbBarraHerramientas.Buttons.Item(14).Enabled = False 'ANULAR
            tlbBarraHerramientas.Buttons.Item(15).Enabled = False 'LIQUIDAR
            gstrImpresion = "O"
            Dim FechaLiquidacion As Date
            FechaLiquidacion = CDate(Format(Now, "dd/mm/yyyy"))
            GeneraRegistroFactura gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, gstrSeccion, txtPatente, lblMarca, lblModelo, lblCliente, gcurInsumo, gcurMateriales, gcurSeguroTaller, lblIdCliente, FechaLiquidacion
            
            'actualizar datos de rent a car
            If Me.dtcGarantia.BoundText = "REN" And Me.optMantencion.Value = True Then
                gstrEstadoDisponible = Retorna_Valor_General("Select EstadoDisponible from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
                gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoDisponible & "', "
                gstrSql = gstrSql & " KilometrajeActual=" & Me.txtKilAct
                gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
                If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
                End If
                
                'actualiza valores en auto stock
                gstrSql = "UPDATE Rent_Anexo_Auto_Stock SET Fecha_Ultima_Mantencion='" & Me.pckFechaEntrega & "',"
                gstrSql = gstrSql & " Kilometraje_Ultima_Mantencion='" & Me.txtKilAct & "'"
                gstrSql = gstrSql & " Where Id_Cajon_Pedido='" & Me.lblVin & "'"
                If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
                End If
                
            End If
            If Me.dtcGarantia.BoundText = "REN" And Me.optReparacion.Value = True Then
                gstrEstadoDisponible = Retorna_Valor_General("Select EstadoDisponible from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
                gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoDisponible & "'"
                gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
                If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
                End If
            End If

            PrintOT
        End If
        MsgBox "La OT Nº " & lblNroRecepcion & " Fue Liquidada"
        Bloqueo "L"
    Else
        MsgBox "Lo siento, La Contraseña Ingresada no es la Correcta"
    End If
Else
    DoEvents
End If
End Sub


Function CalculoMateriales(IndiceSubItem As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwRepuestos
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoMateriales Then
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
        End If
    Next
End With
CalculoMateriales = dblPreSuma
End Function
Function CalculoInsumos(IndiceSubItem As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwRepuestos
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoInsumos Then
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
        End If
    Next
End With
CalculoInsumos = dblPreSuma
End Function
Function CalculoLubricantes(IndiceSubItem As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwRepuestos
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoLubricantes Then
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
        End If
    Next
End With
CalculoLubricantes = dblPreSuma
End Function

Sub LimpiaLinea()
With Me
    .dtcConceptos.BoundText = ""
    .txtSeccion = ""
    .dtcPartePieza.BoundText = ""
    .txtHorasCar = ""
    .txtValorDefCar = ""
    .txtPorcDesCar = ""
    .txtMtoDesCar = ""
    .txtValorFinCar = ""
    '.dtcCargoCar.BoundColumn = ""
    .dtcMecanicoCar.BoundText = ""
    .dtcConceptos.SetFocus
End With
End Sub


Sub ServicioCarroceria(Accion As mAccionItem)
If Accion = mAddItem Then
    Set itmAux = lvwServiciosCarroceria.ListItems.Add(, , dtcConceptos.Text)
    Set lvwServiciosCarroceria.SelectedItem = itmAux
    itmAux.SubItems(1) = dtcConceptos.BoundText
    itmAux.SubItems(2) = txtSeccion.Text
    itmAux.SubItems(3) = dtcPartePieza.Text
    itmAux.SubItems(4) = dtcPartePieza.BoundText
    itmAux.SubItems(5) = FormatoValor(IIf(txtHorasCar <> "", txtHorasCar, 0), "", 1)
    itmAux.SubItems(6) = FormatoValor(IIf(txtValorDefCar <> "", txtValorDefCar, 0), "", gintDecimalesMoneda)
    itmAux.SubItems(7) = FormatoValor(IIf(txtPorcDesCar <> "", txtPorcDesCar, 0), "", 2)
    itmAux.SubItems(8) = FormatoValor(IIf(txtMtoDesCar <> "", txtMtoDesCar, 0), "", gintDecimalesMoneda)
    itmAux.SubItems(9) = FormatoValor(IIf(txtValorFinCar <> "", txtValorFinCar, 0), "", gintDecimalesMoneda)
    itmAux.SubItems(10) = IIf(dtcCargoCar = "", TraeCargoDes(gstrIdCargo), dtcCargoCar.Text)
    itmAux.SubItems(11) = IIf(dtcCargoCar = "", gstrIdCargo, dtcCargoCar.BoundText)
    itmAux.SubItems(12) = dtcMecanicoCar.Text  'TraeNombreMecanico(gstrMecanicoDefectoSecCar)
    itmAux.SubItems(13) = dtcMecanicoCar.BoundText  'gstrMecanicoDefectoSecCar
    itmAux.SubItems(14) = FormatoValor(CalculoSubTotal(mcFichaCarroceria), "", gintDecimalesMoneda)
    itmAux.SubItems(15) = "N"
End If
If Accion = mDelItem Then
    If lvwServiciosCarroceria.ListItems.Count > 0 Then
        lvwServiciosCarroceria.ListItems.Remove lvwServiciosCarroceria.SelectedItem.Index
    End If
End If
End Sub
Sub AsignaTotal(Seccion As mcFicha, Objeto As statusBar)
Dim Resta As Double

If Seccion = mcFichaMecanica Then '///////////total mecanica
    With Objeto
        .Panels(2).Text = FormatoValor(TotalSeccion(lvwServiciosMecanica, 10), "", gintDecimalesMoneda)
    End With
ElseIf Seccion = mcFichaCarroceria Then '///////////total carroceria
    With Objeto
        .Panels(2).Text = FormatoValor(TotalSeccion(lvwServiciosCarroceria, 16), "", gintDecimalesMoneda)
        stbTotalDesabolladura.Panels(2).Text = FormatoValor(SubTotalDesabolladura, "", gintDecimalesMoneda)
        stbTotalPintura.Panels(2).Text = FormatoValor(SubTotalPintura, "", gintDecimalesMoneda)
        stbTotalArmeyDesarme.Panels(2).Text = FormatoValor(SubTotalArmeDesarme, "", gintDecimalesMoneda)
    End With
ElseIf Seccion = mcFichaTerceros Then '///////////total terceros
    With Objeto
        .Panels(2).Text = FormatoValor(TotalSeccion(lvwServiciosTerceros, 12), "", gintDecimalesMoneda)
    End With
ElseIf Seccion = mcFichaRepuestos Then '///////////total repuestos
    With Objeto
        gcurMateriales = CalculoMateriales(8)
        gcurLubricantes = CalculoLubricantes(8)
        'gcurMateriales = gcurMateriales + gcurLubricantes
        Resta = CalculoInsumos(8) + gcurMateriales + gcurLubricantes
        '.Panels(2).Text = FormatoValor(TotalSeccion(lvwRepuestos, 8) - Resta + CalculoInsumos(8), "", 0)   // suma los insumos a los repuestos
        .Panels(2).Text = FormatoValor(TotalSeccion(lvwRepuestos, 8) - Resta, "", gintDecimalesMoneda)
        stbTotalMateriales.Panels(2).Text = FormatoValor(gcurMateriales + CalculoInsumos(8), "", gintDecimalesMoneda) '// sumo insumos a materiales
        StbLubricantes.Panels(2).Text = FormatoValor(gcurLubricantes, "", gintDecimalesMoneda)
    End With
ElseIf Seccion = mcFichaOtros Then '///////////total otros
    With Objeto
        .Panels(2).Text = FormatoValor(TotalSeccion(lvwOtrosServicios, 10), "", gintDecimalesMoneda)
    End With
End If
End Sub
Sub LimpiaTotales()
With Me
    .stbTotalMec.Panels(2).Text = "0"
    .stbTotalCarroceria.Panels(2).Text = "0"
    .stbTotalDesabolladura.Panels(2).Text = "0"
    .stbTotalPintura.Panels(2).Text = "0"
    .stbTotalArmeyDesarme.Panels(2) = "0"
    .stbTotalOtros.Panels(2).Text = "0"
    .stbTotalTerceros.Panels(2).Text = "0"
    .stbTotalRepuestos.Panels(2).Text = "0"
    .stbTotalMateriales.Panels(2).Text = "0"
    .StbLubricantes.Panels(2).Text = "0"
    .stbTotalOT.Panels(2).Text = "0"
End With
End Sub

Function SubTotalDesabolladura() As Double
Dim intS As Integer
Dim dblPreSuma As Double

dblPreSuma = 0
With lvwServiciosCarroceria
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        If .SelectedItem.SubItems(3) = "D" Then
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(.SelectedItem.SubItems(16), ""))
        End If
    Next
End With
SubTotalDesabolladura = dblPreSuma
End Function

Function SubTotalPintura() As Double
Dim intS As Integer
Dim dblPreSuma As Double

dblPreSuma = 0
With lvwServiciosCarroceria
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        If .SelectedItem.SubItems(3) = "P" Then
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(.SelectedItem.SubItems(16), ""))
        End If
    Next
End With
SubTotalPintura = dblPreSuma

End Function
Function SubTotalArmeDesarme() As Double
Dim intS As Integer
Dim dblPreSuma As Double

dblPreSuma = 0
With lvwServiciosCarroceria
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        If .SelectedItem.SubItems(3) = "A" Then
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(.SelectedItem.SubItems(16), ""))
        End If
    Next
End With
SubTotalArmeDesarme = dblPreSuma

End Function

Sub TotalFinal()
    stbTotalOT.Panels(2).Text = FormatoValor(TotalOT, "", gintDecimalesMoneda)
End Sub

Function TotalOT() As Double
Dim dblSemiTotal As Double
With Me
    dblSemiTotal = Val(SacarFormatoValor(.stbTotalMec.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalCarroceria.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalOtros.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalTerceros.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalRepuestos.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + IIf(Not IsNull(gcurInsumo), gcurInsumo, 0)
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalMateriales.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.StbLubricantes.Panels(2).Text, ""))
End With
TotalOT = dblSemiTotal
End Function
Function CalculoSubTotal(Ficha As mcFicha) As Double
Dim Total As Double

Total = 0
If Ficha = mcFichaMecanica Then
    With lvwServiciosMecanica
        If .ListItems.Count > 0 Then
        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(2), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))
        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
        End If
    End With
ElseIf Ficha = mcFichaCarroceria Then
    With lvwServiciosCarroceria
        If .ListItems.Count > 0 Then
            Total = Val(SacarFormatoValor(.SelectedItem.SubItems(5), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(9), ""))
            Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(8), ""))
        End If
    End With
ElseIf Ficha = mcFichaTerceros Then
    With lvwServiciosTerceros
        If .ListItems.Count > 0 Then
        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(4), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(7), ""))
        End If
    End With
ElseIf Ficha = mcFichaRepuestos Then
    With lvwRepuestos
        If .ListItems.Count > 0 Then
        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(2), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))
        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
        End If
    End With
ElseIf Ficha = mcFichaOtros Then
    With lvwOtrosServicios
        If .ListItems.Count > 0 Then
        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(2), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))
        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
        End If
    End With
End If
    CalculoSubTotal = Total
End Function

Function TotalSeccion(lvwObjeto As ListView, IndiceSubItem As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwObjeto
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
    Next
End With
TotalSeccion = dblPreSuma
End Function
Function DatosCliente(strIdCliente As String) As Boolean
If strIdCliente <> "" Then
    mstrSQL = "SELECT Glbl_Cliente_Proveedor.Razon_Social as NOMBRE, Glbl_Cliente_Proveedor.Direccion AS DIREC, Glbl_Comuna.Descripcion AS COMUNA, Glbl_Cliente_Proveedor.Rut AS RUT ,Glbl_Cliente_Proveedor.Telefono AS FONO FROM Glbl_Cliente_Proveedor INNER JOIN Glbl_Comuna ON Glbl_Cliente_Proveedor.Id_Comuna = Glbl_Comuna.Id_Comuna "
    mstrSQL = mstrSQL & " Where Glbl_Cliente_Proveedor.Id_Cliente_Proveedor='" & strIdCliente & "'"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                lblCliente = IIf(Not IsNull(!Nombre), !Nombre, "")
                txtDir = IIf(Not IsNull(!DirEC), !DirEC, "")
                txtComuna = IIf(Not IsNull(!Comuna), !Comuna, "")
                txtRut = IIf(Not IsNull(!rut), !rut, "")
                lblFono = ValorNulo(!FONO)
            End If
        End With
    End If
    Conexion.CloseHost AdoPrincipal
End If
End Function

Function ExisteRegistro(IdCiaSeguro As String, IdConcepto As String, IdPtePza As String) As Boolean
Dim adoTemp As New ADODB.Recordset
ExisteRegistro = False
mstrSQL = "SELECT top 1 * From Tllr_CiaSeguro_Concepto_Parte_Pieza"
mstrSQL = mstrSQL & " WHERE Id_Compañia_Seguro = '" & IdCiaSeguro & "'  AND Id_Concepto = '" & IdConcepto & "' AND Id_Parte_Pieza = '" & IdPtePza & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
        ExisteRegistro = True
    Else
        mstrSQL = "Insert into Tllr_CiaSeguro_Concepto_Parte_Pieza (Id_Compañia_Seguro, Id_Concepto, Id_Parte_Pieza, Valor, Horas) Values ('" & IdCiaSeguro & "' ,'" & IdConcepto & "' ,'" & IdPtePza & "',0,0)"
        If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            ExisteRegistro = True
        Else
            ExisteRegistro = False
        End If
    End If
End If
End Function

Sub FillInventarioOT(strIdEmpresa As String, strIdSucursal As String, strIdRecepcion As String, strSeccion As String)
SetCheckOff lvwInventario
mstrSQL = "SELECT Id_Estado_Recepcion as Codigo From Tllr_Inventario_Presupuesto"
mstrSQL = mstrSQL & " WHERE Id_Empresa = '" & strIdEmpresa & "' AND Id_Sucursal = '" & strIdSucursal & "' AND Id_OT = '" & strIdRecepcion & "' AND Seccion_OT = '" & strSeccion & "'"
mstrSQL = mstrSQL & " Order by Id_Estado_Recepcion"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        While Not .EOF
            Set lvwInventario.SelectedItem = lvwInventario.FindItem(CStr(!Codigo), , , 1)
            lvwInventario.SelectedItem.Checked = True
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub

Sub FillMecanicaOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
lvwServiciosMecanica.ListItems.Clear
mstrSQL = " SELECT Tllr_Mecanica_Presupuesto.Id_Servicio AS ID, Tllr_Servicio.Descripcion,"
mstrSQL = mstrSQL & " Tllr_Mecanica_Presupuesto.Horas, Tllr_Mecanica_Presupuesto.Valor,"
mstrSQL = mstrSQL & " Tllr_Mecanica_Presupuesto.Porcentaje_Descuento AS PORC,"
mstrSQL = mstrSQL & " Tllr_Mecanica_Presupuesto.Monto_Descuento AS MONTO,"
mstrSQL = mstrSQL & " Tllr_Mecanica_Presupuesto.Id_Tipo_Cargo AS IDCARGO,"
mstrSQL = mstrSQL & " Tllr_Tipo_Cargo.Descripcion AS CARGO,"
mstrSQL = mstrSQL & " Tllr_Mecanica_Presupuesto.Mecanico_Designado AS IDMEC,"
mstrSQL = mstrSQL & " Tllr_Mecanicos.Nombre AS MEC,"
mstrSQL = mstrSQL & " Tllr_Mecanica_Presupuesto.SubTotal AS TOTAL, Tllr_Mecanica_Presupuesto.Facturado"
mstrSQL = mstrSQL & " FROM Tllr_Mecanicos RIGHT OUTER JOIN Tllr_Mecanica_Presupuesto ON Tllr_Mecanicos.Id_Mecanico = Tllr_Mecanica_Presupuesto.Mecanico_Designado"
mstrSQL = mstrSQL & " LEFT OUTER JOIN Tllr_Tipo_Cargo ON Tllr_Mecanica_Presupuesto.Id_Tipo_Cargo = Tllr_Tipo_Cargo.Id_Tipo_Cargo and Tllr_Mecanica_Presupuesto.Id_Empresa = Tllr_Tipo_Cargo.Id_Empresa "
mstrSQL = mstrSQL & " LEFT OUTER JOIN Tllr_Servicio RIGHT OUTER JOIN Tllr_Servicio_Modelo ON  Tllr_Servicio.Id_Servicio = Tllr_Servicio_Modelo.Id_Servicio ON"
mstrSQL = mstrSQL & " Tllr_Mecanica_Presupuesto.Id_Marca = Tllr_Servicio_Modelo.Id_Marca AND Tllr_Mecanica_Presupuesto.Id_Modelo = Tllr_Servicio_Modelo.Id_Modelo"
mstrSQL = mstrSQL & " AND Tllr_Mecanica_Presupuesto.Id_Servicio = Tllr_Servicio_Modelo.Id_Servicio"
mstrSQL = mstrSQL & " WHERE (Tllr_Mecanica_Presupuesto.Id_Empresa = '" & strIdEmpresa & "') AND"
mstrSQL = mstrSQL & " (Tllr_Mecanica_Presupuesto.Id_Sucursal = '" & strIdSucursal & "') AND"
mstrSQL = mstrSQL & " (Tllr_Mecanica_Presupuesto.Id_Ot = '" & strIdDocumento & "') AND"
mstrSQL = mstrSQL & " (Tllr_Mecanica_Presupuesto.Seccion_OT = '" & strSeccion & "')"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwServiciosMecanica.ListItems.Add(, , ValorNulo(!ID))
            Set lvwServiciosMecanica.SelectedItem = itmAux
            itmAux.SubItems(1) = ValorNulo(!Descripcion)
            itmAux.SubItems(2) = FormatoValor(!Horas, "", 1)
            itmAux.SubItems(3) = FormatoValor(!Valor, "", gintDecimalesMoneda)
            itmAux.SubItems(4) = FormatoValor(!PORC, "", 2)
            itmAux.SubItems(5) = FormatoValor(!MONTO, "", gintDecimalesMoneda)
            itmAux.SubItems(6) = ValorNulo(!IDCARGO)
            itmAux.SubItems(7) = IIf(ValorNulo(!CARGO) = "", "(Ninguno)", !CARGO)
            itmAux.SubItems(8) = ValorNulo(!idmec)
            itmAux.SubItems(9) = IIf(ValorNulo(!mec) = "", "(Ninguno)", !mec)
            itmAux.SubItems(10) = FormatoValor(!Total, "", gintDecimalesMoneda)
            itmAux.SubItems(11) = ValorNulo(!Facturado)
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub
Sub FillRepuestosReservados(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
lvwRepuestosMantencion.ListItems.Clear

mstrSQL = "SELECT Tllr_Repuestos_Reservados.Id_Item, "
mstrSQL = mstrSQL & "Stck_Item.Descripcion, Tllr_Repuestos_Reservados.Solicitado, "
mstrSQL = mstrSQL & "Tllr_Repuestos_Reservados.Precio_Unitario, Tllr_Repuestos_Reservados.Estado, "
mstrSQL = mstrSQL & "Glbl_Familia.Descripcion AS Familia, "
mstrSQL = mstrSQL & "Tllr_Repuestos_Reservados.Id_OT "
mstrSQL = mstrSQL & "FROM Tllr_Repuestos_Reservados INNER JOIN "
mstrSQL = mstrSQL & "Stck_Item ON "
mstrSQL = mstrSQL & "Tllr_Repuestos_Reservados.Id_Item = Stck_Item.Id_Item INNER "
mstrSQL = mstrSQL & "Join "
mstrSQL = mstrSQL & "Glbl_Familia ON "
mstrSQL = mstrSQL & "Stck_Item.Id_Familia = Glbl_Familia.Id_Familia "
mstrSQL = mstrSQL & " WHERE (Tllr_Repuestos_Reservados.Id_Empresa = '" & strIdEmpresa & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Reservados.Id_Sucursal = '" & strIdSucursal & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Reservados.Id_OT = '" & strIdDocumento & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Reservados.Seccion_OT = '" & strSeccion & "')"

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwRepuestosMantencion.ListItems.Add(, , ValorNulo(!Id_Item))
            Set lvwRepuestosMantencion.SelectedItem = itmAux
            itmAux.SubItems(1) = ValorNulo(!Descripcion)
            itmAux.SubItems(2) = FormatoValor(!Solicitado, "", 1)
            itmAux.SubItems(3) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
            itmAux.SubItems(4) = ValorNulo(!Familia)
            'itmAux.SubItems(5) = lvwServiciosMecanica.SelectedItem.SubItems(6)
            
            If !estado = "S" Then
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HFF0000
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HFF0000
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HFF0000
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HFF0000
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HFF0000
                'lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HFF0000
            End If
            If !estado = "P" Then
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HC0&
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
                lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
            End If
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub
Sub FillRepuestosFaltantes(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)

mstrSQL = "SELECT Tllr_Repuestos_Faltantes.Id_Item, "
mstrSQL = mstrSQL & "Stck_Item.Descripcion, Tllr_Repuestos_Faltantes.Solicitado, "
mstrSQL = mstrSQL & "Tllr_Repuestos_Faltantes.Precio_Unitario, "
mstrSQL = mstrSQL & "Glbl_Familia.Descripcion AS Familia, "
mstrSQL = mstrSQL & "Tllr_Repuestos_Faltantes.Id_OT "
mstrSQL = mstrSQL & "FROM Tllr_Repuestos_Faltantes INNER JOIN "
mstrSQL = mstrSQL & "Stck_Item ON "
mstrSQL = mstrSQL & "Tllr_Repuestos_Faltantes.Id_Item = Stck_Item.Id_Item INNER "
mstrSQL = mstrSQL & "Join "
mstrSQL = mstrSQL & "Glbl_Familia ON "
mstrSQL = mstrSQL & "Stck_Item.Id_Familia = Glbl_Familia.Id_Familia "
mstrSQL = mstrSQL & " WHERE (Tllr_Repuestos_Faltantes.Id_Empresa = '" & strIdEmpresa & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Faltantes.Id_Sucursal = '" & strIdSucursal & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Faltantes.Id_OT = '" & strIdDocumento & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Faltantes.Seccion_OT = '" & strSeccion & "')"

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwRepuestosMantencion.ListItems.Add(, , ValorNulo(!Id_Item))
            Set lvwRepuestosMantencion.SelectedItem = itmAux
            itmAux.SubItems(1) = ValorNulo(!Descripcion)
            itmAux.SubItems(2) = FormatoValor(!Solicitado, "", 1)
            itmAux.SubItems(3) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
            itmAux.SubItems(4) = ValorNulo(!Familia)
            itmAux.SubItems(5) = lvwServiciosMecanica.SelectedItem.SubItems(6)
            
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
                
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub


Sub FillCarroceriaOT(strIdEmpresa As String, strIdSucursal As String, strIdRecepcion As String, strSeccion As String, strIdCiaSeguro As String)
lvwServiciosCarroceria.ListItems.Clear

'/// mod por fdo diaz 29/01/2001 elimine la cia de seguros
mstrSQL = "SELECT Tllr_Carroceria_Presupuesto.Id_Concepto AS IDCONCEP,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.D_P,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Descripcion AS DescCarr,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Id_Parte_Pieza AS IDPARTE,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Horas, Tllr_Carroceria_Presupuesto.Valor,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Valor_definido AS DEFINIDO,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Porcentaje_Descuento AS PORC,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Monto_Descuento AS MONTO,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Porcentaje_Recargo AS PORCREC,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Monto_Recargo AS MONTOREC,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Id_Tipo_Cargo AS IDCARGO,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Id_Servicio_Carroceria AS codigo,"
mstrSQL = mstrSQL & " Tllr_Tipo_Cargo.Descripcion AS CARGO,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Id_Proveedor AS IDPROV,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Facturado,"
mstrSQL = mstrSQL & " Glbl_Cliente_Proveedor.Razon_Social AS Provee,"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.SubTotal"
mstrSQL = mstrSQL & " FROM Glbl_Cliente_Proveedor RIGHT OUTER JOIN"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto ON"
mstrSQL = mstrSQL & " Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Carroceria_Presupuesto.Id_Proveedor"
mstrSQL = mstrSQL & " LEFT OUTER JOIN"
mstrSQL = mstrSQL & " Tllr_Tipo_Cargo ON"
mstrSQL = mstrSQL & " Tllr_Carroceria_Presupuesto.Id_Tipo_Cargo = Tllr_Tipo_Cargo.Id_Tipo_Cargo and Tllr_Carroceria_Presupuesto.Id_Empresa = Tllr_Tipo_Cargo.Id_Empresa "
mstrSQL = mstrSQL & " WHERE (Tllr_Carroceria_Presupuesto.Id_Empresa = '" & strIdEmpresa & "') AND"
mstrSQL = mstrSQL & " (Tllr_Carroceria_Presupuesto.Id_Sucursal = '" & strIdSucursal & "') AND"
mstrSQL = mstrSQL & " (Tllr_Carroceria_Presupuesto.Id_OT = '" & strIdRecepcion & "') AND"
mstrSQL = mstrSQL & " (Tllr_Carroceria_Presupuesto.Seccion_OT = '" & strSeccion & "')"
mstrSQL = mstrSQL & " Order by Tllr_Carroceria_Presupuesto.Id_Servicio_Carroceria"

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwServiciosCarroceria.ListItems.Add(, , "")          '///des concepto
            itmAux.SubItems(1) = IIf(IsNull(!IDCONCEP), "", !IDCONCEP)                                            '///id concepto
            itmAux.SubItems(2) = IIf(IsNull(!DescCarr), "", !DescCarr)                   '///d_p
            itmAux.SubItems(3) = IIf(IsNull(!D_P), "", !D_P)                                               '/// des parte
            itmAux.SubItems(4) = IIf(IsNull(!IDPARTE), "", !IDPARTE)                                             '///idparte
            itmAux.SubItems(5) = FormatoValor(!Horas, "", 1)                              '///valor definido Format(ValorNulo(!HORAS), "#0.0")
            itmAux.SubItems(6) = FormatoValor(!Valor, "", gintDecimalesMoneda)
            itmAux.SubItems(7) = FormatoValor(!PORCREC, "", 2)
            itmAux.SubItems(8) = FormatoValor(!MONTOREC, "", gintDecimalesMoneda)
            itmAux.SubItems(9) = FormatoValor(!DEFINIDO, "", gintDecimalesMoneda)
            itmAux.SubItems(10) = FormatoValor(!PORC, "", 2)
            itmAux.SubItems(11) = FormatoValor(!MONTO, "", gintDecimalesMoneda)
            itmAux.SubItems(12) = !CARGO
            itmAux.SubItems(13) = !IDCARGO
            itmAux.SubItems(14) = IIf(ValorNulo(!Provee) = "", "(Ninguno)", !Provee)
            itmAux.SubItems(15) = ValorNulo(!IDPROV)
            itmAux.SubItems(16) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
            itmAux.SubItems(17) = ValorNulo(!Facturado)
            itmAux.SubItems(18) = IIf(IsNull(!Codigo), 1, !Codigo)
          
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub
Sub FillOtrosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
lvwOtrosServicios.ListItems.Clear
mstrSQL = "SELECT Tllr_Otro_Presupuesto.Id_Otro_Servicio as ID,"
mstrSQL = mstrSQL & " Tllr_Otro_Presupuesto.Descripcion_Otro AS DES , Tllr_Otro_Presupuesto.Horas AS TIEMPO,"
mstrSQL = mstrSQL & " Tllr_Otro_Presupuesto.Valor AS UNITARIO, Tllr_Otro_Presupuesto.Porcentaje_Descuento AS PORCDESC,"
mstrSQL = mstrSQL & " Tllr_Otro_Presupuesto.Monto_Descuento AS MTODESC, Tllr_Otro_Presupuesto.Id_Tipo_Cargo AS IDCARGO,"
'mstrSql = mstrSql & " Tllr_Tipo_Cargo.Descripcion AS CARGO,"
mstrSQL = mstrSQL & " Tllr_Otro_Presupuesto.Mecanico_Asignado AS IDMEC," ' Tllr_Mecanicos.Nombre AS MECANICO,"
mstrSQL = mstrSQL & " Tllr_Otro_Presupuesto.SubTotal, Tllr_Otro_Presupuesto.Facturado "
mstrSQL = mstrSQL & " FROM Tllr_Otro_Presupuesto " 'LEFT OUTER JOIN Tllr_Mecanicos ON Tllr_Presupuestoro_OT.Mecanico_Asignado = Tllr_Mecanicos.Id_Mecanico  LEFT OUTER JOIN Tllr_Tipo_Cargo ON Tllr_Presupuestoro_OT.Id_Tipo_Cargo = Tllr_Tipo_Cargo.Id_Tipo_Cargo"
mstrSQL = mstrSQL & " WHERE (Tllr_Otro_Presupuesto.Id_Empresa = '" & strIdEmpresa & "') AND"
mstrSQL = mstrSQL & " (Tllr_Otro_Presupuesto.Id_Sucursal = '" & strIdSucursal & "') AND"
mstrSQL = mstrSQL & " (Tllr_Otro_Presupuesto.Id_OT = '" & strIdDocumento & "') AND"
mstrSQL = mstrSQL & " (Tllr_Otro_Presupuesto.Seccion_OT = '" & strSeccion & "')"

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwOtrosServicios.ListItems.Add(, , !ID)            '///des concepto
            itmAux.SubItems(1) = !Des                                              '///id concepto
            itmAux.SubItems(2) = !TIEMPO                                                   '///d_p
            itmAux.SubItems(3) = FormatoValor(!UNITARIO, "", gintDecimalesMoneda)                                               '/// des parte)
            itmAux.SubItems(4) = FormatoValor(!PORCDESC, "", 2)                                 '///valor definido Format(ValorNulo(!HORAS), "#0.0")
            itmAux.SubItems(5) = FormatoValor(!MTODESC, "", gintDecimalesMoneda)
            itmAux.SubItems(6) = !IDCARGO
            itmAux.SubItems(7) = TraeCargoDes(!IDCARGO)
            itmAux.SubItems(8) = !idmec
            itmAux.SubItems(9) = MecanicoD(!idmec)
            itmAux.SubItems(10) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
            itmAux.SubItems(11) = ValorNulo(!Facturado)
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal

End Sub


Sub FillTercerosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
lvwServiciosTerceros.ListItems.Clear
mstrSQL = " SELECT Tllr_Terceros_Presupuesto.Id_Servicio_Tercero AS IDSERVICIO, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Descripcion AS SERVICIO, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Id_Proveedor AS IDPROV, "
mstrSQL = mstrSQL & " Glbl_Cliente_Proveedor.Razon_Social AS PROVEEDOR, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.NroFarctura AS NROFACT, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Valor AS PREUNI, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Cantidad AS CANTY, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Porcentaje_Recargo AS PRECARGO, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Monto_Recargo AS MRECARGO, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Porcentaje_Dscto AS PDSCTO, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Monto_Dscto AS MDSCTO, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Precio_Final as PREFIN, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.SubTotal AS STOTAL, "
mstrSQL = mstrSQL & " Tllr_Terceros_Presupuesto.Id_Tipo_Cargo AS IDCARGO, Tllr_Terceros_Presupuesto.Facturado,"
mstrSQL = mstrSQL & " Tllr_Tipo_Cargo.Descripcion AS CARGO "
'mstrSql = mstrSql & " FROM Tllr_Terceros_Presupuesto LEFT OUTER JOIN Tllr_Tipo_Cargo ON Tllr_Terceros_Presupuesto.Id_Tipo_Cargo = Tllr_Tipo_Cargo.Id_Tipo_Cargo AND Tllr_Terceros_Presupuesto.Id_Tipo_Cargo = Tllr_Tipo_Cargo.Id_Tipo_Cargo LEFT OUTER JOIN Tllr_Proveedor_Servicio ON Tllr_Terceros_Presupuesto.Id_Proveedor = Tllr_Proveedor_Servicio.Id_Proveedor "
mstrSQL = mstrSQL & " FROM Tllr_Terceros_Presupuesto LEFT OUTER JOIN Glbl_Cliente_Proveedor ON Tllr_Terceros_Presupuesto.Id_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor LEFT OUTER JOIN Tllr_Tipo_Cargo ON Tllr_Terceros_Presupuesto.Id_Tipo_Cargo = Tllr_Tipo_Cargo.Id_Tipo_Cargo AND Tllr_Terceros_Presupuesto.Id_Empresa = Tllr_Tipo_Cargo.Id_Empresa "
mstrSQL = mstrSQL & " WHERE (Tllr_Terceros_Presupuesto.Id_Empresa = '" & strIdEmpresa & "') AND "
mstrSQL = mstrSQL & " (Tllr_Terceros_Presupuesto.Id_Sucursal = '" & strIdSucursal & "') AND "
mstrSQL = mstrSQL & " (Tllr_Terceros_Presupuesto.Id_OT = '" & strIdDocumento & "') AND "
mstrSQL = mstrSQL & " (Tllr_Terceros_Presupuesto.Seccion_OT = '" & strSeccion & "') "


If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwServiciosTerceros.ListItems.Add(, , !idServicio)            '///des concepto
            itmAux.SubItems(1) = ValorNulo(!Proveedor)  '///id concepto
            itmAux.SubItems(2) = ValorNulo(!IDPROV)
            itmAux.SubItems(3) = ValorNulo(!servicio) '/// des parte
            itmAux.SubItems(4) = ValorNulo(!NROFACT)
            itmAux.SubItems(5) = FormatoValor(!PREUNI, "", gintDecimalesMoneda)
            itmAux.SubItems(6) = FormatoValor(!CANTY, "", 1)                                 '///valor definido Format(ValorNulo(!HORAS), "#0.0")
            itmAux.SubItems(7) = FormatoValor(!PRECARGO, "", 2)
            itmAux.SubItems(8) = FormatoValor(!MRECARGO, "", gintDecimalesMoneda)
            itmAux.SubItems(9) = FormatoValor(!PREFIN, "", gintDecimalesMoneda)
            itmAux.SubItems(10) = FormatoValor(IIf(IsNull(!PDSCTO), "0", !PDSCTO), "", 2)
            itmAux.SubItems(11) = FormatoValor(IIf(IsNull(!MDSCTO), "0", !MDSCTO), "", gintDecimalesMoneda)
            itmAux.SubItems(12) = FormatoValor(!STotal, "", gintDecimalesMoneda)
            itmAux.SubItems(13) = TraeCargoDes(!IDCARGO)
            itmAux.SubItems(14) = !IDCARGO
            itmAux.SubItems(15) = ValorNulo(!Facturado)
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
mstrSQL = ""
End Sub
Sub FillRepuestosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
lvwRepuestos.ListItems.Clear
mstrSQL = "SELECT Tllr_Repuestos_Presupuesto.Id_Item AS ID,"
mstrSQL = mstrSQL & " Stck_Item.Descripcion AS ITEM,"
mstrSQL = mstrSQL & " Tllr_Repuestos_Presupuesto.Cantidad AS CANTY,"
mstrSQL = mstrSQL & " Tllr_Repuestos_Presupuesto.Valor AS VALOR,"
mstrSQL = mstrSQL & " Tllr_Repuestos_Presupuesto.Porcentaje_Descuento AS PORCDES,"
mstrSQL = mstrSQL & " Tllr_Repuestos_Presupuesto.Monto_Descuento AS MTODES,"
mstrSQL = mstrSQL & " Tllr_Repuestos_Presupuesto.Id_Tipo_Cargo AS IDCARGO,"
mstrSQL = mstrSQL & " Tllr_Repuestos_Presupuesto.SubTotal AS SUBTOTAL, Tllr_Repuestos_Presupuesto.Facturado, Tllr_Repuestos_Presupuesto.Consumo"
mstrSQL = mstrSQL & " FROM Tllr_Repuestos_Presupuesto LEFT OUTER JOIN"
mstrSQL = mstrSQL & " Stck_Item ON Tllr_Repuestos_Presupuesto.Id_Item = Stck_Item.Id_Item"
mstrSQL = mstrSQL & " WHERE (Tllr_Repuestos_Presupuesto.Id_Empresa = '" & strIdEmpresa & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Presupuesto.Id_Sucursal = '" & strIdSucursal & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Presupuesto.Id_Ot = '" & strIdDocumento & "') AND"
mstrSQL = mstrSQL & " (Tllr_Repuestos_Presupuesto.Seccion_OT = '" & strSeccion & "')"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            If !CANTY > 0 Then  '///valores > 0
                Set itmAux = lvwRepuestos.ListItems.Add(, , !ID)            '///des concepto
                itmAux.SubItems(1) = ValorNulo(!Item)                                              '///id concepto
                itmAux.SubItems(2) = FormatoValor(!CANTY, "", 1)
                itmAux.SubItems(3) = FormatoValor(!Valor, "", gintDecimalesMoneda)
                itmAux.SubItems(4) = FormatoValor(!PORCDES, "", 2)
                itmAux.SubItems(5) = FormatoValor(!MTODES, "", gintDecimalesMoneda)
                itmAux.SubItems(6) = TraeCargoDes(!IDCARGO)
                itmAux.SubItems(7) = !IDCARGO
                itmAux.SubItems(8) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
                itmAux.SubItems(9) = FamiliaRep(!ID)
                itmAux.SubItems(10) = ValorNulo(!Facturado)
                itmAux.SubItems(11) = IIf(IsNull(!Consumo), "STOCK", IIf(!Consumo = "C", "STOCK", "PRESUPUESTO"))
            End If
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub
Sub DatosVehiculo(strPatente As String)
If strPatente <> "" Then
    mstrSQL = "SELECT Tllr_Vehiculo_Cliente.Patente,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Marca AS IDMARCA,"
    mstrSQL = mstrSQL & " Glbl_Marca.Descripcion AS MARCA,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Modelo AS IDMODELO,"
    mstrSQL = mstrSQL & " Glbl_Modelo.Descripcion AS MODELO,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Año,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Color_Exterior AS IDCOLOR,"
    mstrSQL = mstrSQL & " Glbl_Color_Exterior.Descripcion AS COLOR,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Kilometros_Actuales AS KILACT,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Nro_Motor AS MOTOR,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Nro_Chasis AS CHASIS,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.VIN AS VIN,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor AS IDCLI,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Fecha_Venta AS FECVTA,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Concesionario AS CONCES"
    mstrSQL = mstrSQL & " FROM Glbl_Cliente_Proveedor RIGHT OUTER JOIN Glbl_Color_Exterior RIGHT OUTER JOIN Tllr_Vehiculo_Cliente ON Glbl_Color_Exterior.Id_Color_Exterior = Tllr_Vehiculo_Cliente.Id_Color_Exterior LEFT OUTER JOIN Glbl_Modelo LEFT OUTER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca ON Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor"
    '///NEO
    'mstrSql = mstrSql & " WHERE Tllr_Vehiculo_Cliente.Patente='" & txtPatente & "'"
    mstrSQL = mstrSQL & " WHERE Tllr_Vehiculo_Cliente.Patente='" & strPatente & "'"
    '///
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            With AdoPrincipal
                lblMarca = ValorNulo(!Marca)
                lblIdMarca = ValorNulo(!IdMarca)
                lblModelo = ValorNulo(!Modelo)
                lblIdModelo = ValorNulo(!IdModelo)
                lblChasis = ValorNulo(!chasis)
                lblMotor = ValorNulo(!motor)
                lblVin = ValorNulo(!VIN)
                txtAño = ValorNulo(!Año)
                lblColorE = ValorNulo(!Color)
                'lblCliente = ValorNulo(!idCLI)
                txtConcesionario = ValorNulo(!CONCES)
                pckFecVta.Value = IIf(Not IsNull(!FECVTA), !FECVTA, Now)
                txtKilAct = IIf(Not IsNull(!kilact), !kilact, "0")
                lblIdCliente = ValorNulo(!idCLI)
                KilometrajeEntrada = txtKilAct 'Variable de ileiva 07/02/2001
            End With
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End If
End Sub
Sub FillConceptosInventario()
mstrSQL = "SELECT Id_Estado_Recepcion AS Codigo, Descripcion AS Nombre FROM Tllr_Estado_Recepcion WHERE Vigencia = 'S' Order By Id_Estado_Recepcion"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set itmAux = lvwInventario.ListItems.Add(, , !Codigo)
                itmAux.SubItems(1) = !Nombre
                .MoveNext
            Wend
        End If
    End With
End If
End Sub
Private Function GuardaCarroceria(strIdDocumento As String, strSeccion As String, strCiaSeguro As String, gParametro As gcParametro) As Boolean
Dim mstrNombreTabla As String

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Carroceria_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Carroceria_Presupuesto"
End If

GuardaCarroceria = True
mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' "
If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
    With lvwServiciosCarroceria
        If .ListItems.Count > 0 Then
            For intIndice = 1 To .ListItems.Count
                Set .SelectedItem = .ListItems(intIndice)
                '/////////////////////////////////////////////////VALIDAR SI EXISTE EN PARENT
                'If ExisteRegistro(strCiaSeguro, .SelectedItem.SubItems(1), .SelectedItem.SubItems(4)) = True Then
                    mstrSQL = "INSERT INTO " & mstrNombreTabla
                    mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
                    mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
                    mstrSQL = mstrSQL & " Id_Compañia_Seguro, "
                    mstrSQL = mstrSQL & " Id_Concepto, "
                    mstrSQL = mstrSQL & " D_P,"
                    mstrSQL = mstrSQL & " Id_Parte_Pieza, "
                    mstrSQL = mstrSQL & " Id_Tipo_Cargo, Mecanico_Designado,"
                    mstrSQL = mstrSQL & " Horas, Valor,Valor_Definido ,"
                    mstrSQL = mstrSQL & " Porcentaje_Descuento,Monto_Descuento,"
                    mstrSQL = mstrSQL & " SubTotal,Facturado,Porcentaje_Recargo,Monto_Recargo,Id_Proveedor,Descripcion,Id_Servicio_Carroceria)"
                    mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "       '///empresa, sucursal
                    mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & strSeccion & "',"                  '///nro ot, seccion
                    mstrSQL = mstrSQL & " '" & strCiaSeguro & "', "                                         '///cia seguro
                    mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(1)) & "', "                      '///concepto
                    mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(3) & "',"                                                   'Trim(.SelectedItem.SubItems(2)) ///d_p
                    mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(4)) & "', "                      '///parte y pieza
                    mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(13) & "','" & gstrMecanicoDefectoSecCar & "',"            '///mecanico designado
                    mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(5), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(6), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(9), "######.00"))) & " ,"
                    mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "######.00"))) & ","
                    mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(16), "######.00"))) & ",'" & .SelectedItem.SubItems(17) & "',"
                    mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "######.00"))) & ","
                    mstrSQL = mstrSQL & " " & IIf(.SelectedItem.SubItems(15) = "", "NULL" & ",", " '" & .SelectedItem.SubItems(15) & "',")
                    mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(2)) & "',"
                    mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(18)) & "')"
                    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                        GuardaCarroceria = False
                        Exit Function
                    End If
                'End If
            Next
        Else
            GuardaCarroceria = True
        End If
    End With
Else
    GuardaCarroceria = False
    Exit Function
End If
End Function
Private Function GuardaTerceros(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
Dim mstrNombreTabla As String

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Terceros_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Terceros_Presupuesto"
End If

GuardaTerceros = True
mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' "
If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
    With lvwServiciosTerceros
        If .ListItems.Count > 0 Then
            For intIndice = 1 To .ListItems.Count
                Set .SelectedItem = .ListItems(intIndice)
                mstrSQL = "INSERT INTO " & mstrNombreTabla
                mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
                mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
                mstrSQL = mstrSQL & " Id_Proveedor, "
                mstrSQL = mstrSQL & " Id_Servicio_Tercero,"
                mstrSQL = mstrSQL & " Id_Tipo_Cargo, "
                mstrSQL = mstrSQL & " Cantidad,Valor,"
                mstrSQL = mstrSQL & " Porcentaje_Recargo,Monto_Recargo,"
                mstrSQL = mstrSQL & " Precio_Final,"
                mstrSQL = mstrSQL & " Descripcion , NroFarctura, "
                mstrSQL = mstrSQL & " SubTotal, Facturado, "
                mstrSQL = mstrSQL & " Porcentaje_Dscto, Monto_Dscto)"
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(2) & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem) & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(14)) & "', "
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(6), "#####0.0"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "#####0.0"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(9), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(3) & "', "
                mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(4) & "', "
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(12), "#####0.0"))) & ",'" & .SelectedItem.SubItems(15) & "',"
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "#####0.0"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "#####0.0"))) & ")"
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaTerceros = False
                    Exit Function
                End If
            Next
        Else
            GuardaTerceros = True
        End If
    End With
Else
    GuardaTerceros = False
    Exit Function
End If
End Function

Private Function GuardaOtros(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
Dim mstrNombreTabla As String

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Presupuestoro_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Presupuestoro_Presupuesto"
End If

GuardaOtros = True
mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' "
If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
    With lvwOtrosServicios
        If .ListItems.Count > 0 Then
            For intIndice = 1 To .ListItems.Count
                Set .SelectedItem = .ListItems(intIndice)
                mstrSQL = "INSERT INTO " & mstrNombreTabla
                mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
                mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
                mstrSQL = mstrSQL & " Id_Otro_Servicio, "
                mstrSQL = mstrSQL & " Id_Tipo_Cargo,"
                mstrSQL = mstrSQL & " Mecanico_Asignado, "
                mstrSQL = mstrSQL & " Horas,Valor,"
                mstrSQL = mstrSQL & " Porcentaje_Descuento,Monto_Descuento,"
                mstrSQL = mstrSQL & " SubTotal,Descripcion_Otro,Facturado)"
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(6)) & "', "
                mstrSQL = mstrSQL & " '" & IIf(Trim(.SelectedItem.SubItems(8)) = "", "SIN", Trim(.SelectedItem.SubItems(8))) & "', "
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.0"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.0"))) & ",'" & UCase(Trim(.SelectedItem.SubItems(1))) & "','" & UCase(Trim(.SelectedItem.SubItems(11))) & "')"
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaOtros = False
                    Exit Function
                End If
            Next
        Else
            GuardaOtros = True
        End If
    End With
Else
    GuardaOtros = False
    Exit Function
End If
End Function
Private Function GuardaRepuestos(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
Dim mstrNombreTabla As String

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Repuestos_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Repuestos_Presupuesto"
End If

GuardaRepuestos = True

If Me.dtcGarantia.BoundText = "PRE" Then  'elimina solo si son presupuestos
    mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' "
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
End If

With lvwRepuestos
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            If VerificaRepuesto(.SelectedItem, lblNroRecepcion, strSeccion, mstrNombreTabla) = True Then
                mstrSQL = "UPDATE " & mstrNombreTabla
                mstrSQL = mstrSQL & " SET Id_Tipo_Cargo='" & Trim(.SelectedItem.SubItems(7)) & "',"
                mstrSQL = mstrSQL & " Cantidad = " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.0"))) & ", "
                mstrSQL = mstrSQL & " Valor = " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " Porcentaje_Descuento = " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " Monto_Descuento = " & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " SubTotal = " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " Facturado = " & UCase(Trim(IIf(.SelectedItem.SubItems(10) = "", "'N'", "'" & .SelectedItem.SubItems(10) & "'"))) & ","
                mstrSQL = mstrSQL & " Consumo = '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "'"
                mstrSQL = mstrSQL & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
                mstrSQL = mstrSQL & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
                mstrSQL = mstrSQL & " Id_OT = '" & strIdDocumento & "' AND  "
                mstrSQL = mstrSQL & " Seccion_OT = '" & strSeccion & "' AND "
                mstrSQL = mstrSQL & " Id_Item = '" & .SelectedItem & "' "
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaRepuestos = False
                    Exit Function
                End If
            Else
                '///////////////////////////////////VALIDAR SI EXISTE EN PARENT
                mstrSQL = "INSERT INTO " & mstrNombreTabla
                mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
                mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
                mstrSQL = mstrSQL & " Id_Item, "
                mstrSQL = mstrSQL & " Id_Tipo_Cargo, "
                mstrSQL = mstrSQL & " Cantidad, Valor,"
                mstrSQL = mstrSQL & " Porcentaje_Descuento,Monto_Descuento,"
                mstrSQL = mstrSQL & " SubTotal,Facturado,Consumo)"
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(7)) & "', "
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.0"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.0"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.0"))) & ","
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.0"))) & ",'" & .SelectedItem.SubItems(10) & "',"
                mstrSQL = mstrSQL & " '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "')"
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaRepuestos = False
                    Exit Function
                End If
            End If
        Next
    Else
        GuardaRepuestos = True
    End If
End With
End Function
Private Function GuardaInventario(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
Dim mstrNombreTabla As String

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Inventario_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Inventario_Presupuesto"
End If


GuardaInventario = True
mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND " & mstrNombreTabla & ".ID_OT='" & strIdDocumento & "' and " & mstrNombreTabla & ".Seccion_OT = '" & strSeccion & "'"
If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
    For intIndice = 1 To lvwInventario.ListItems.Count
        Set lvwInventario.SelectedItem = lvwInventario.ListItems(intIndice)
        If lvwInventario.SelectedItem.Checked = True Then
            mstrSQL = "Insert Into " & mstrNombreTabla
            mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,Id_Estado_Recepcion, Id_OT, Seccion_OT) "
            mstrSQL = mstrSQL & " values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "','" & lvwInventario.SelectedItem & "', '" & strIdDocumento & "', '" & strSeccion & "' )"
            If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                GuardaInventario = False
                Exit Function
            End If
        End If
    Next
    GuardaInventario = True
Else
    GuardaInventario = False
    Exit Function
End If
End Function
Private Function GuardaMecanica(strIdDocumento As String, gParametro As gcParametro) As Boolean
Dim mstrNombreTabla As String

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Mecanica_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Mecanica_Presupuesto"
End If

GuardaMecanica = True
mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And ID_OT='" & strIdDocumento & "' And Seccion_OT='" & gstrSeccion & "'"
If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
    With lvwServiciosMecanica
        If .ListItems.Count > 0 Then
            For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            mstrSQL = "Insert Into " & mstrNombreTabla
            mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
            mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
            mstrSQL = mstrSQL & " Id_Marca, Id_Modelo, "
            mstrSQL = mstrSQL & " Id_Servicio, "
            mstrSQL = mstrSQL & " Id_Tipo_Cargo,Mecanico_Designado,"
            mstrSQL = mstrSQL & " Horas,Valor,"
            mstrSQL = mstrSQL & " Porcentaje_Descuento, Monto_Descuento, "
            mstrSQL = mstrSQL & " SubTotal, Facturado)"
            mstrSQL = mstrSQL & " Values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
            mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & gstrSeccion & "',"
            mstrSQL = mstrSQL & " '" & Trim(lblIdMarca) & "','" & Trim(lblIdModelo) & "',"
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem) & "',"
            mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(6) & "'," & IIf(.SelectedItem.SubItems(8) = "", "NULL", " '" & .SelectedItem.SubItems(8) & "' ") & ", "
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.0"))) & " , " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.0"))) & " , "
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.0"))) & " ," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.0"))) & ","
            mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.0"))) & ",'" & .SelectedItem.SubItems(11) & "' )"
            If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                GuardaMecanica = False
                Exit Function
            End If
            Next
        Else
            GuardaMecanica = True
        End If
    End With
Else
    GuardaMecanica = False
    Exit Function
End If
End Function

Function letSql(strWhere As String, strOrder As String) As String
mstrSQL = "SELECT Id_OT, "
mstrSQL = mstrSQL & " Seccion_OT, "
mstrSQL = mstrSQL & " Patente, "
mstrSQL = mstrSQL & " Id_Garantia as TipoOT, "
mstrSQL = mstrSQL & " Folio_Garantia,"
mstrSQL = mstrSQL & " Id_Tipo_Cono, "
mstrSQL = mstrSQL & " Nro_Cono, "
mstrSQL = mstrSQL & " RealizadoPor, "
mstrSQL = mstrSQL & " Fecha_Emision,"
mstrSQL = mstrSQL & " Entrega_Estimada, "
mstrSQL = mstrSQL & " Hora_Entrega, "
mstrSQL = mstrSQL & " Nro_Siniestro, "
mstrSQL = mstrSQL & " Nro_Poliza,"
mstrSQL = mstrSQL & " Liquidador, "
mstrSQL = mstrSQL & " Deducible_UF, "
mstrSQL = mstrSQL & " Deducible_Pesos, "
mstrSQL = mstrSQL & " Id_Compañia_Seguro, "
mstrSQL = mstrSQL & " Solicitado_Por, "
mstrSQL = mstrSQL & " Total_Mecanica, "
mstrSQL = mstrSQL & " Total_Carroceria,"
mstrSQL = mstrSQL & " Total_Desabolladura, "
mstrSQL = mstrSQL & " Total_Pintura, "
mstrSQL = mstrSQL & " Total_Terceros,"
mstrSQL = mstrSQL & " Total_Materiales,"
mstrSQL = mstrSQL & " Total_Insumos,"
mstrSQL = mstrSQL & " Total_Repuestos , "
mstrSQL = mstrSQL & " Total_OT, "
mstrSQL = mstrSQL & " Estado, "
mstrSQL = mstrSQL & " Comentario, "
mstrSQL = mstrSQL & " ReparacionMantencion, "
mstrSQL = mstrSQL & " Estado_Reserva, "
mstrSQL = mstrSQL & " Id_Presupuesto "
mstrSQL = mstrSQL & " From Tllr_Presupuesto"
letSql = mstrSQL & " " & strWhere & " " & strOrder

End Function

Private Sub LeerCampos()

If mblnTablaVacia Then
    LimpiaCampos
    Exit Sub
End If
With AdoPrincipal
    If !Seccion_OT = "C" Then
        Me.optRecepcion(1).Value = True
    Else
        Me.optRecepcion(0).Value = True
    End If
    lblNroRecepcion.Text = ValorNulo(!Id_Presupuesto)
    mstrIdPresupuestoOrigen = ValorNulo(!Id_Presupuesto)
    lblPresupuesto = ValorNulo(!Id_OT)
    dtcGarantia.BoundText = !TipoOt
    gstrIdCargo = TraeCargo(!TipoOt)
    dtcTipoCono.BoundText = ValorNulo(!Id_Tipo_Cono)
    dtcRecepcionista.BoundText = ValorNulo(!RealizadoPor)
    txtNroCono = ValorNulo(!Nro_Cono)
    pckFechaAtencion.Value = !Fecha_Emision
    pckFechaEntrega.Value = !Entrega_Estimada
    cboHora.Text = ValorNulo(!Hora_Entrega)
    txtNroSiniestro = ValorNulo(!Nro_Siniestro)
    txtNroPoliza = ValorNulo(!Nro_Poliza)
    txtLiquidador = ValorNulo(!Liquidador)
    
    txtDeducibleUF = ValorNulo(!Deducible_UF)
    txtDeduciblePesos = ValorNulo(!deducible_pesos)
    lblCompañia.Tag = ValorNulo(!Id_Compañia_Seguro)
    lblCompañia = CiaSegDes(ValorNulo(!Id_Compañia_Seguro))
    
    txtComentario = ValorNulo(!Comentario)
    txtPatente = ValorNulo(!Patente)
    txtFolioGarantia = ValorNulo(!Folio_Garantia)
    txtSolicita = ValorNulo(!Solicitado_Por)
    gcurInsumo = ValorNulo(!Total_Insumos)
    'stbInsumos.Panels(2).Text = FormatoValor(!Total_Insumos, "", 0)
    If Not IsNull(!estado) Then
        lblEstadoOTValor.Caption = IIf(!estado = "V", "VIGENTE", IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "N", "NULA", IIf(!estado = "F" Or !estado = "B", "EMITIDA", IIf(!estado = "R", "RESERVA", IIf(!estado = "P", "PRESUPUESTO", ""))))))
        tlbBarraHerramientas.Buttons.Item(1).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", True, IIf(!estado = "R", True, False)))))
        tlbBarraHerramientas.Buttons.Item(2).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", True, IIf(!estado = "R", True, False)))))
        tlbBarraHerramientas.Buttons.Item(13).Enabled = IIf(!estado = "V", False, IIf(!estado = "L", True, IIf(!estado = "N", True, IIf(!estado = "F" Or !estado = "B", False, False))))    'ACTIVAR
        tlbBarraHerramientas.Buttons.Item(14).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, False))))    'ANULAR
        tlbBarraHerramientas.Buttons.Item(15).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", True, False))))    'LIQUIDAR
        tlbBarraHerramientas.Buttons.Item(20).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Separador
        tlbBarraHerramientas.Buttons.Item(21).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Confirmar Reserva
        tlbBarraHerramientas.Buttons.Item(22).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Eliminar Reserva
        tlbBarraHerramientas.Buttons.Item(24).Visible = IIf(!estado = "P", True, False) 'Liquidar presupuesto
        tlbBarraHerramientas.Buttons.Item(25).Visible = IIf(!estado = "P", True, False) 'Liquidar presupuesto
        Bloqueo ValorNulo(!estado)
    End If
    If ValorNulo(!Patente) <> "" Then DatosVehiculo !Patente
    '/////////////////////////////////////////////////////////////////////////////////
    FillConceptosVsCiaSeguro dtcConceptos, datConceptos, lblCompañia.Tag
    '/////////////////////////////////////////////////////////////////////////////////
    FillInventarioOT gstrIdEmpresa, gstrIdSucursal, !Id_Presupuesto, gstrSeccion
    '/////////////////////////////////////////////////////////////////////////////////
    FillMecanicaOT gstrIdEmpresa, gstrIdSucursal, !Id_Presupuesto, gstrSeccion
    AsignaTotal mcFichaMecanica, stbTotalMec
    
    
    '/////////////////////////////////////////////////////////////////////////////////
    'If !Seccion_OT = "C" Then
        FillCarroceriaOT gstrIdEmpresa, gstrIdSucursal, !Id_Presupuesto, gstrSeccion, lblCompañia.Tag
        AsignaTotal mcFichaCarroceria, stbTotalCarroceria
    'Else
    '    lvwServiciosCarroceria.ListItems.Clear
    '    frmRecepcion.stbTotalCarroceria.Panels(2).Text = 0
    'End If
    '/////////////////////////////////////////////////////////////////////////////////
    FillOtrosOT gstrIdEmpresa, gstrIdSucursal, !Id_Presupuesto, gstrSeccion
    AsignaTotal mcFichaOtros, stbTotalOtros
    '/////////////////////////////////////////////////////////////////////////////////
    FillTercerosOT gstrIdEmpresa, gstrIdSucursal, !Id_Presupuesto, gstrSeccion
    AsignaTotal mcFichaTerceros, stbTotalTerceros
    '/////////////////////////////////////////////////////////////////////////////////
    FillRepuestosOT gstrIdEmpresa, gstrIdSucursal, !Id_Presupuesto, gstrSeccion
    AsignaTotal mcFichaRepuestos, stbTotalRepuestos
'    stbTotalMateriales.Panels(2).Text = Format(CalculoMateriales(8))
    '/////////////////////////////////////////////////////////////////////////////////
    
        '//// Si no encuentra reserva de repuestos busca los repuestos de los servicios
        Dim i As Integer
        lvwRepuestosMantencion.ListItems.Clear
        For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
            mstrAgregaPresupuesto = False
            Repuestos_de_la_Mantencion Me.lblIdMarca, Me.lblIdModelo, lvwServiciosMecanica.ListItems(i), IIf(Me.lvwServiciosMecanica.ListItems(i).SubItems(12) = "S", True, False)
        Next
    
    TotalFinal
    '/////////////////////////////////////////////////////////////////////////////////
End With
End Sub

Function VerificaServicioCarroceria(strIdConcepto As String, strIdParte As String) As Boolean
VerificaServicioCarroceria = True
For intIndice = 1 To lvwServiciosCarroceria.ListItems.Count
    Set lvwServiciosCarroceria.SelectedItem = lvwServiciosCarroceria.ListItems(intIndice)
    If lvwServiciosCarroceria.SelectedItem.SubItems(1) = strIdConcepto Then
        If lvwServiciosCarroceria.SelectedItem.SubItems(4) = strIdParte Then
            VerificaServicioCarroceria = False
            Exit Function
        Else
            VerificaServicioCarroceria = True
        End If
    Else
        VerificaServicioCarroceria = True
    End If
Next intIndice
End Function

Private Sub cmdAnularReserva_Click()
Dim EstadoReserva As String

If Me.lvwRepuestosMantencion.ListItems.Count > 0 Then

    '/// valida que la reserva no haya pasado a Consumo
    EstadoReserva = Retorna_Valor_General("Select Estado_Reserva from Stck_Regularizacion Where Id_OT='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
    If EstadoReserva = "L" Then
        MsgBox "Esta Reserva ya paso a ser un Consumo...", vbInformation, "Anular Reserva de Repuestos"
        Exit Sub
    End If
    'Levanta listview con los repuestos de la mantencion
    If MsgBox(" Esta Seguro de Anular esta esta Reserva de Repuestos ", vbQuestion + vbYesNo, "Confirma Anulación") = vbYes Then
        NroRegularizacion = Retorna_Valor_General("Select Id_Regularizacion as Numero from Stck_Regularizacion where id_ot='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
        Call Actualiza_Saldos_VS_Detalle("S", "Select Canrtidad, Id_Empresa, Id_sucursal, Id_Bodega,Id_Ubicacion,Id_Item From Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & NroRegularizacion & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Empresa = '" & gstrIdEmpresa & "'")
        
        EliminaReservaRepuestos NroRegularizacion, lblNroRecepcion
        
        '/// Actualiza estado de reserva
        mstrSQL = "UPDATE Tllr_Presupuesto SET Estado_Reserva='N' "
        mstrSQL = mstrSQL & "Where Id_OT='" & frmRecepcion.lblNroRecepcion & "' "
        mstrSQL = mstrSQL & "And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Seccion_OT='" & gstrSeccion & "'"
        Conexion.SendHost mstrSQL, , , , gcTiempoEspera
        DesactivaBotonAnularReserva
        
    Else
        Exit Sub
    End If
End If

End Sub
Sub DesactivaBotonAnularReserva()
    cmdAnularReserva.Enabled = False
    cmdReserva.Enabled = True
End Sub
Private Sub cmdReserva_Click()
If Me.lvwRepuestosMantencion.ListItems.Count > 0 Then
    'Levanta listview con los repuestos de la mantencion
    If MsgBox(" Las Cantidades ya estan Confirmadas ? ", vbQuestion + vbYesNo, "Verifica Cantidades") = vbYes Then
        GrabaReservaRepuestosRecepcion
        frmRepuestosReservados.Show vbModal
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub dtcConceptos_Change()
txtSeccion = TipoConcepto(dtcConceptos.BoundText)
End Sub
Private Sub dtcGarantia_Change()
'mstrCargo = TraeCargo(dtcGarantia.BoundText)
'TipoOt dtcGarantia.BoundText
'gstrIdCargo = mstrCargo
End Sub
Private Sub dtcPartePieza_Change()
txtHorasCar = TraeHorasDefinidas(lblCompañia.Tag, dtcConceptos.BoundText, dtcPartePieza.BoundText)
txtValorDefCar = TraeValorDefinido(lblCompañia.Tag, dtcConceptos.BoundText, dtcPartePieza.BoundText)
txtValorFinCar = TraeValorDefinido(lblCompañia.Tag, dtcConceptos.BoundText, dtcPartePieza.BoundText)
End Sub

Private Sub dtcTipoCono_Click(Area As Integer)
If Area > 0 Then
    txtNroCono.SetFocus
End If
End Sub

Private Sub Form_Load()
    mblnSW = True
    gstrSeccion = "M"
    stbServicios.tab = 0
    gstrKmsAutoNuevo = ""
    mstrLiquidaPresupuesto = False
    'gcurInsumoDef = gcurInsumo
End Sub
Private Sub lblIdCliente_Change()
If DatosCliente(lblIdCliente) Then DoEvents
End Sub

Private Sub lblNroRecepcion_DblClick()
If gstrImpresion = "O" And Me.lblNroRecepcion <> "" Then
    gstrBusca = InputBox("Ingrese El Numero de O/T Deseado :", "Ir a....", CStr(Val(Mid(lblNroRecepcion, 6, Len(lblNroRecepcion) - 5))))
    gstrBusca = FormatPresupuesto(gstrBusca)
    If gstrBusca <> "" Then
        mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.ID_Presupuesto=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_OT"
        gstrSql = letSql(mstrWhere, mstrOrderBy)
        If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost AdoPrincipal
    End If
    Me.SetFocus
End If
End Sub

Private Sub lvwOtrosServicios_DblClick()
If mblnBloqueo = False And Me.lvwOtrosServicios.ListItems.Count > 0 Then
    With lvwOtrosServicios
        If .SelectedItem.SubItems(11) <> "S" Then
            If Not .SelectedItem Is Nothing Then
                frmEditaOtroServicio.Show vbModal
                AsignaTotal mcFichaOtros, stbTotalOtros
                TotalFinal
            End If
        Else
            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
        End If
    End With
End If
End Sub
Private Sub lvwOtrosServicios_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If mblnBloqueo = False Then
    Dim i As Integer
    Dim gstrBusca As String
    
    Select Case Button
        Case vbRightButton  '//BOTON DERECHO
            gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
            If IsNumeric(gstrBusca) Then
                If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
                    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
                        If Me.lvwOtrosServicios.ListItems(i).Selected Then
                            dblTotalInicial = Round(CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(2)) * CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(3)), 2)
                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                        End If
                        
                    Next
                    AsignaTotal mcFichaOtros, stbTotalOtros
                    TotalFinal
                Else
                    MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
                End If
            Else
                MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
            End If
    End Select
End If

End Sub

Private Sub lvwRepuestos_DblClick()
If mblnBloqueo = False And Me.lvwRepuestos.ListItems.Count > 0 Then
    With lvwRepuestos
        If .SelectedItem.SubItems(10) <> "S" Then
            If Not .SelectedItem Is Nothing Then
                frmEditaServicioRepuesto.Show vbModal
                gitmActual = .SelectedItem.Index
                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                TotalFinal
                Set .SelectedItem = .ListItems(gitmActual)
            End If
        Else
            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
        End If
    End With
End If
End Sub
Private Sub lvwRepuestos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If mblnBloqueo = False Then
    Dim i As Integer
    Dim gstrBusca As String
    
    Select Case Button
        Case vbRightButton  '//BOTON DERECHO
            gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
            If IsNumeric(gstrBusca) Then
                If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
                    For i = 1 To Me.lvwRepuestos.ListItems.Count
                        If Me.lvwRepuestos.ListItems(i).Selected Then
                            dblTotalInicial = Round(CDbl(Me.lvwRepuestos.ListItems.Item(i).SubItems(2)) * CDbl(Me.lvwRepuestos.ListItems.Item(i).SubItems(3)), 2)
                            Me.lvwRepuestos.ListItems.Item(i).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                            Me.lvwRepuestos.ListItems.Item(i).SubItems(8) = FormatoValor(dblTotalInicial - CDbl(Me.lvwRepuestos.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                            Me.lvwRepuestos.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                        End If
                        
                    Next
                    AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                    TotalFinal
                Else
                    MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
                End If
            Else
                MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
            End If
    End Select
End If
End Sub

Private Sub lvwRepuestosMantencion_DblClick()
If mblnBloqueo = False Then
If lvwRepuestosMantencion.ListItems.Count > 0 And Me.cmdReserva.Enabled = True Then
strMode = "Edit"
Set lsiItem = lvwRepuestosMantencion.SelectedItem
With frmEditaTempRepuesto
    .Caption = "Editar Repuesto"
    .txtMarca = frmRecepcion.lblMarca
    .txtModelo = frmRecepcion.lblModelo
    '.txtServicio = frmRecepcion.lvwServiciosMecanica.ListItems(1).SubItems(1)
    .txtCodigo = lsiItem
    .txtDescripcion = lsiItem.SubItems(1)
    .txtValor = SacarFormatoValor(lsiItem.SubItems(3), "")
    .txtCantidad = SacarFormatoValor(lsiItem.SubItems(2), "")
    .Show 1
End With
End If
End If
End Sub

Private Sub lvwServiciosCarroceria_DblClick()
If mblnBloqueo = False And Me.lvwServiciosCarroceria.ListItems.Count > 0 Then
    With lvwServiciosCarroceria
        If .SelectedItem.SubItems(17) <> "S" Then
            If Not .SelectedItem Is Nothing Then
                gitmActual = .SelectedItem.Index
                frmEditaTrabajoCarroceria.Show vbModal
                AsignaTotal mcFichaCarroceria, stbTotalCarroceria
                TotalFinal
                Set .SelectedItem = .ListItems(gitmActual)
            End If
        Else
            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
        End If
    End With
End If
End Sub

Private Sub lvwServiciosCarroceria_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If mblnBloqueo = False Then
    Dim i As Integer
    Dim gstrBusca As String
    
    Select Case Button
        Case vbRightButton  '//BOTON DERECHO
            gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
            If IsNumeric(gstrBusca) Then
                If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
                    For i = 1 To Me.lvwServiciosCarroceria.ListItems.Count
                        If Me.lvwServiciosCarroceria.ListItems(i).Selected Then
                            If Trim(Me.lvwServiciosCarroceria.ListItems.Item(i).SubItems(5)) <> "0.0" Then
                                dblTotalInicial = Round(CDbl(Me.lvwServiciosCarroceria.ListItems.Item(i).SubItems(5)) * CDbl(Me.lvwServiciosCarroceria.ListItems.Item(i).SubItems(9)), 2)
                                Me.lvwServiciosCarroceria.ListItems.Item(i).SubItems(11) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                                Me.lvwServiciosCarroceria.ListItems.Item(i).SubItems(16) = FormatoValor(dblTotalInicial - CDbl(Me.lvwServiciosCarroceria.ListItems.Item(i).SubItems(11)), "", gintDecimalesMoneda)
                                Me.lvwServiciosCarroceria.ListItems.Item(i).SubItems(10) = FormatoValor(Val(gstrBusca), "", 2)
                            End If
                        End If
                    Next
                    AsignaTotal mcFichaCarroceria, stbTotalCarroceria
                    TotalFinal
                Else
                    MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
                End If
            Else
                MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
            End If
    End Select
End If

End Sub

Private Sub lvwServiciosMecanica_DblClick()
If mblnBloqueo = False And Me.lvwServiciosMecanica.ListItems.Count > 0 Then
    With lvwServiciosMecanica
        If .SelectedItem.SubItems(11) <> "S" Then
            If Not .SelectedItem Is Nothing Then
                gitmActual = .SelectedItem.Index
                frmEditaServicioMecanica.Show vbModal
                AsignaTotal mcFichaMecanica, stbTotalMec
                TotalFinal
                Set .SelectedItem = .ListItems(gitmActual)
            End If
        Else
            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
        End If
    End With
End If
End Sub
Private Sub lvwServiciosMecanica_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If mblnBloqueo = False Then

    Dim i As Integer
    Dim gstrBusca As String
    
    Select Case Button
        Case vbRightButton  '//BOTON DERECHO
            gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
            If IsNumeric(gstrBusca) Then
                If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
                    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
                        If Me.lvwServiciosMecanica.ListItems(i).Selected Then
                            dblTotalInicial = Round(CDbl(Me.lvwServiciosMecanica.ListItems.Item(i).SubItems(2)) * CDbl(Me.lvwServiciosMecanica.ListItems.Item(i).SubItems(3)), 2)
                            Me.lvwServiciosMecanica.ListItems.Item(i).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                            Me.lvwServiciosMecanica.ListItems.Item(i).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(Me.lvwServiciosMecanica.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
                            Me.lvwServiciosMecanica.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
                        End If
                        
                    Next
                    AsignaTotal mcFichaMecanica, stbTotalMec
                    TotalFinal
                Else
                    MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
                End If
            Else
                MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
            End If
    End Select
End If
End Sub

Private Sub lvwServiciosTerceros_DblClick()
If mblnBloqueo = False And Me.lvwServiciosTerceros.ListItems.Count > 0 Then
With lvwServiciosTerceros
    If .SelectedItem.SubItems(15) <> "S" Then
        If Not .SelectedItem Is Nothing Then
            frmEditaServicioTercero.Show 1
            gitmActual = .SelectedItem.Index
            AsignaTotal mcFichaTerceros, stbTotalTerceros
            TotalFinal
            Set .SelectedItem = .ListItems(gitmActual)
        End If
    Else
        MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
    End If
End With
End If
End Sub

Private Sub lvwServiciosTerceros_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If mblnBloqueo = False Then
    Dim i As Integer
    Dim gstrBusca As String
    
    Select Case Button
        Case vbRightButton  '//BOTON DERECHO
            gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
            If IsNumeric(gstrBusca) Then
                If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
                    For i = 1 To Me.lvwServiciosTerceros.ListItems.Count
                        If Me.lvwServiciosTerceros.ListItems(i).Selected Then
                            If Trim(Me.lvwServiciosTerceros.ListItems.Item(i).SubItems(6)) <> "0.0" Then
                                dblTotalInicial = Round(CDbl(Me.lvwServiciosTerceros.ListItems.Item(i).SubItems(6)) * CDbl(Me.lvwServiciosTerceros.ListItems.Item(i).SubItems(9)), 2)
                                Me.lvwServiciosTerceros.ListItems.Item(i).SubItems(11) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
                                Me.lvwServiciosTerceros.ListItems.Item(i).SubItems(12) = FormatoValor(dblTotalInicial - CDbl(Me.lvwServiciosTerceros.ListItems.Item(i).SubItems(11)), "", gintDecimalesMoneda)
                                Me.lvwServiciosTerceros.ListItems.Item(i).SubItems(10) = FormatoValor(Val(gstrBusca), "", 2)
                            End If
                        End If
                    Next
                    AsignaTotal mcFichaTerceros, stbTotalTerceros
                    TotalFinal
                Else
                    MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
                End If
            Else
                MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
            End If
    End Select
End If

End Sub

Private Sub optRecepcion_Click(Index As Integer)
Select Case Index
Case 0
    stbServicios.tab = 0
    gstrSeccion = "M"
    Renovar
    'stbServicios.TabEnabled(3) = False
Case 1
    stbServicios.tab = 0
    gstrSeccion = "C"
    Renovar
    'stbServicios.TabEnabled(3) = True
End Select
End Sub

Private Sub tlbAddRep_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Agregar" ' ////////////////AGREGAR
        If Trim(txtPatente.Text) <> "" Then
            gstrProcedencia = "Movimientos"
            frmSelTempRepuestos.Show vbModal
            AsignaTotal mcFichaRepuestos, stbTotalRepuestos
            TotalFinal
        End If
    Case "Quitar" ' ////////////////QUITAR
        If Not lvwRepuestos.SelectedItem Is Nothing Then
            If AccesoEliminar(lvwRepuestos.SelectedItem) = True Then
                lvwRepuestos.ListItems.Remove (lvwRepuestos.SelectedItem.Index)
                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                TotalFinal
            Else
                MsgBox ""
            End If
        End If
    End Select
End Sub

Private Sub tlbAddServicioCar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case Is = "Agregar"
    If Trim(txtPatente) <> "" Then
        'If dtcConceptos.BoundText <> "" And dtcPartePieza.BoundText <> "" Then
            'If VerificaServicioCarroceria(dtcConceptos.BoundText, dtcPartePieza.BoundText) Then
             '   Call ServicioCarroceria(mAddItem)
             '   AsignaTotal mcFichaCarroceria, stbTotalCarroceria
             '   TotalFinal
             '   LimpiaLinea
                
                frmAddTrabajosCarroceria.Show vbModal
                AsignaTotal mcFichaCarroceria, stbTotalCarroceria
                'AsignaTotal mcFichaTerceros, stbTotalTerceros
                TotalFinal
                
            'Else
            '    MsgBox LoadResString(324) & Chr(13) & LoadResString(325)
            'End If
            
        'End If
    Else
        MsgBox LoadResString(301), vbOKOnly, LoadResString(4)
    End If
Case Is = "Quitar"
    'If MsgBox(LoadResString(801), vbYesNo, LoadResString(4)) = 6 Then
        Call ServicioCarroceria(mDelItem)
        AsignaTotal mcFichaCarroceria, stbTotalCarroceria
        TotalFinal
    'End If
Case Else
    DoEvents
End Select
End Sub
Private Sub tlbAddServicioMec_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer
Dim j As Integer
Dim lstrServicioMecanica As String

Select Case Button.Key
Case Is = "Agregar"
    If Trim(txtPatente) <> "" Then
        frmAddServiciosMarMod.Show 1
        lvwRepuestosMantencion.ListItems.Clear
        mstrAgregaPresupuesto = True
        For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
            Repuestos_de_la_Mantencion Me.lblIdMarca, Me.lblIdModelo, lvwServiciosMecanica.ListItems(i), IIf(Me.lvwServiciosMecanica.ListItems(i).SubItems(12) = "S", True, False)
            
        Next
        AsignaTotal mcFichaMecanica, stbTotalMec
        TotalFinal
        If lvwServiciosMecanica.ListItems.Count > 0 Then
            lvwServiciosMecanica.ListItems(lvwServiciosMecanica.ListItems.Count).SubItems(12) = "N"
        End If
    Else
        MsgBox LoadResString(301), vbOKOnly, LoadResString(4)
    End If
Case Is = "Quitar"
    If (lvwServiciosMecanica.ListItems.Count > 0 And Me.cmdReserva.Enabled = True) Or Me.dtcGarantia.BoundText = "PRE" Then
        lstrServicioMecanica = lvwServiciosMecanica.SelectedItem
       ' If MsgBox(LoadResString(801), vbYesNo, LoadResString(4)) = 6 Then
            If Not lvwServiciosMecanica.SelectedItem Is Nothing Then
                If Me.dtcGarantia.BoundText = "PRE" Then
                    '//// quita los repuestos que se agregaron a la ficha de repuestos
                    Quita_Repuestos_Mantencion Me.lblIdMarca, Me.lblIdModelo, lstrServicioMecanica
                End If
                lvwServiciosMecanica.ListItems.Remove (lvwServiciosMecanica.SelectedItem.Index)
                
                lvwRepuestosMantencion.ListItems.Clear
                For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
                    Repuestos_de_la_Mantencion Me.lblIdMarca, Me.lblIdModelo, lvwServiciosMecanica.ListItems(i), IIf(Me.lvwServiciosMecanica.ListItems(i).SubItems(12) = "S", True, False)
                Next
                AsignaTotal mcFichaMecanica, stbTotalMec
                TotalFinal
            Else
                MsgBox LoadResString(802), vbOKOnly, LoadResString(4)
            End If
       ' End If
    Else
        MsgBox "Si Tiene Una reserva de Repuestos no puede quitar el Servicio", vbExclamation, "Quitar Servicio de Mecanica"
    End If
Case Else
    DoEvents
End Select
End Sub

Private Sub tlbAddServicioOtr_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Agregar"
    If Trim(txtPatente.Text) <> "" Then
        frmAddOtrosServicios.Show vbModal
        AsignaTotal mcFichaOtros, stbTotalOtros
        TotalFinal
    End If
Case "Quitar"
    If lvwOtrosServicios.ListItems.Count > 0 Then
        If Not lvwOtrosServicios.SelectedItem Is Nothing Then
            lvwOtrosServicios.ListItems.Remove lvwOtrosServicios.SelectedItem.Index
            AsignaTotal mcFichaOtros, stbTotalOtros
            TotalFinal
        End If
    End If
End Select
End Sub

Private Sub tlbAddServicioTer_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Agregar" ' ////////////////AGREGAR
    If Trim(txtPatente.Text) <> "" Then
        frmAddTrabajosTercero.Show vbModal
        AsignaTotal mcFichaTerceros, stbTotalTerceros
        TotalFinal
    End If
Case "Quitar" ' ////////////////QUITAR
    If Not lvwServiciosTerceros.SelectedItem Is Nothing Then
        lvwServiciosTerceros.ListItems.Remove (lvwServiciosTerceros.SelectedItem.Index)
        AsignaTotal mcFichaTerceros, stbTotalTerceros
        TotalFinal
    End If
End Select
End Sub

Private Sub tlbAgregarRepuestos_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Agregar" ' ////////////////AGREGAR
        If Trim(txtPatente.Text) <> "" Then
            gstrProcedencia = "Movimientos"
            gstrProcedenciaRptos = "Mantencion"
            frmSelTempRepuestos.Show vbModal
            'AsignaTotal mcFichaRepuestos, stbTotalRepuestos
            'TotalFinal
        End If
    Case "Quitar" ' ////////////////QUITAR
        If Me.cmdReserva.Enabled = True Then
            If Not Me.lvwRepuestosMantencion.SelectedItem Is Nothing Then
                If AccesoEliminar(Me.lvwRepuestosMantencion.SelectedItem) = True Then
                    Me.lvwRepuestosMantencion.ListItems.Remove (Me.lvwRepuestosMantencion.SelectedItem.Index)
                    'AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                    'TotalFinal
                Else
                    MsgBox ""
                End If
            End If
        Else
            MsgBox "Si tiene una Reserva no puede Quitar Repuestos", vbExclamation, "Reserva de Repuestos"
        End If
    End Select

End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            AgregarRegistro
        Case "Grabar"
            GrabarRegistro
        Case "Cancelar"
            CancelarAgregaRegistro
        Case "Borrar"
            BorrarRegistro
        Case "Buscar"
            BuscarRegistro
        Case "Imprimir"
            PrintOT
        Case "Primero"
            PrimerRegistro
        Case "Anterior"
            RegistroAnterior
        Case "Siguiente"
            RegistroSiguiente
        Case "Ultimo"
            UltimoRegistro
        Case "Activar"
            EstadosOT gOTActivar
        Case "Anular"
            EstadosOT gOTAnular
        Case "Liquidar"
            EstadosOT gOTLiquidar
        Case "Renovar"
            Renovar
        Case "Cerrar"
            CerrarSalir
        Case "Confirmar"
            ConfirmarReserva
        Case "Vaciar"
            CancelaReserva
        Case "LiquidarPres"
            LiquidarPresupuesto
        Case "AnularPres"
            AnularPresupuesto
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSW Then
        mblnSW = False
        If Not Atributos("Glbl", "Tllr_20_0060", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If '/////////ojo
        
        FillConceptosInventario
        FillGarantia dtcGarantia, datGarantia, True
        FillRecepcionista dtcRecepcionista, datRecepcionista
        FillTipoCono dtcTipoCono, datTipoCono
        FillTime 9, 20, cboHora
        FillTipoCargo dtcCargoCar, datCargoCar
        FillMecanicos dtcMecanicoCar, datMecanico
        
        '//MODIFICADO POR FERNANDO DIAZ 29/11/2000  DESACTIVA LA FICHA DE CARROCERIA
        'stbServicios.TabEnabled(3) = False
        
        '//Crear registro por defecto...
        If gapAccion = apcrear Then
           AgregarRegistro
           lblNroRecepcion = gstrBusca
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
        
        '//Editar registro por defecto...
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.ID_OT='" & gstrBusca & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_OT"
                gstrSql = letSql(mstrWhere, mstrOrderBy)
                If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                        ActivaBotones
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            Me.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If gapAccion = apninguno Then
           Renovar
        End If
        
        optRecepcion(0).Value = True
    End If
    gapAccion = apninguno
    Screen.MousePointer = vbDefault
    '//AgregarRegistro
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            CancelarAgregaRegistro
        Case 14 And tlbBarraHerramientas.Buttons.Item("Crear").Enabled
            KeyAscii = 0
            AgregarRegistro
        Case 7 And tlbBarraHerramientas.Buttons.Item("Grabar").Enabled
            KeyAscii = 0
            GrabarRegistro
        Case 4 And tlbBarraHerramientas.Buttons.Item("Borrar").Enabled
            KeyAscii = 0
            BorrarRegistro
        Case 2 And tlbBarraHerramientas.Buttons.Item("Buscar").Enabled
            KeyAscii = 0
            BuscarRegistro
        Case 9 And tlbBarraHerramientas.Buttons.Item("Imprimir").Enabled
            KeyAscii = 0
            PrintOT
        Case 16 And tlbBarraHerramientas.Buttons.Item("Primero").Enabled
            KeyAscii = 0
            PrimerRegistro
        Case 1 And tlbBarraHerramientas.Buttons.Item("Anterior").Enabled
            KeyAscii = 0
            RegistroAnterior
        Case 19 And tlbBarraHerramientas.Buttons.Item("Siguiente").Enabled
            KeyAscii = 0
            RegistroSiguiente
        Case 21 And tlbBarraHerramientas.Buttons.Item("Ultimo").Enabled
            KeyAscii = 0
            UltimoRegistro
        Case 18 And tlbBarraHerramientas.Buttons.Item("Renovar").Enabled
            KeyAscii = 0
            Renovar
        Case 3 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    Bloqueo "V"
    ParametrosDefecto gstrIdEmpresa, gstrIdSucursal
    lblEstadoOTValor = ""
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    'dtcGarantia.BoundText = gstrIdTipoOtDefecto
    SetCheckOff lvwInventario
    lvwServiciosMecanica.ListItems.Clear
    lvwRepuestosMantencion.ListItems.Clear
    lvwServiciosCarroceria.ListItems.Clear
    lvwOtrosServicios.ListItems.Clear
    lvwServiciosTerceros.ListItems.Clear
    lvwRepuestos.ListItems.Clear
    LimpiaTotales
    stbServicios.tab = 0
    txtPatente.Enabled = True
    If fmePat.Enabled = True Then
        txtPatente.SetFocus
    End If
    
    '//// que obligatoriamente elija un tipo de OT
    If InStr(gstrEmpresa, "AUTO SUMMIT") = 1 Then
        frmElegirTipoOT.Show vbModal
        dtcGarantia.Enabled = False
    End If
    
    '////si es nuevo muestra la ot PRESUPUESTO
    dtcGarantia.BoundText = gstrIdTipoOtDefecto
    Me.Tag = "Crear"
'    gcurInsumoDef = gcurInsumo
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones                                                                       'AND Tllr_Presupuesto.ID_OT = Tllr_Presupuesto.ID_OT >'" & Trim(lblNroRecepcion) & "'
    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_OT DESC"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.ID_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_OT"
            gstrSql = letSql(mstrWhere, mstrOrderBy)
            If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub GrabarRegistro()
    If Not validacion() Then
        Exit Sub
    End If
    
    If Me.Tag = "Crear" Then
        If Me.dtcGarantia.BoundText <> "PRE" Then  '  And mstrLiquidaPresupuesto = True Then
            lblNroRecepcion = TraeCorrelativo(gcOrdenTrabajo, gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
            'mstrIdPresupuestoOrigen = ""
        Else
            lblNroRecepcion = "P-" & TraeCorrelativoPresupuesto(gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
            mstrIdPresupuestoOrigen = lblNroRecepcion
        End If
        gstrBusca = lblNroRecepcion
        mstrSQL = "INSERT INTO Tllr_Presupuesto "
        mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal, "
        mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
        mstrSQL = mstrSQL & " Id_Garantia, Folio_Garantia, "
        mstrSQL = mstrSQL & " Id_Tipo_Cono, Nro_Cono, "
        mstrSQL = mstrSQL & " Patente, RealizadoPor,"
        mstrSQL = mstrSQL & " Kilometros_Recepcion, Id_Compañia_seguro,"
        mstrSQL = mstrSQL & " Fecha_Proxima_Visita, "                           'Fecha_Liquidacion,"
        mstrSQL = mstrSQL & " Estado,Fecha_Emision, "
        mstrSQL = mstrSQL & " Entrega_Estimada, Hora_Entrega, "
        mstrSQL = mstrSQL & " Nro_Factura_Emitida,Nro_Presupuesto_Origen,"
        mstrSQL = mstrSQL & " Nro_Siniestro, Nro_Poliza, Liquidador, "
        mstrSQL = mstrSQL & " Comentario, Solicitado_Por,"
        mstrSQL = mstrSQL & " Deducible_UF , Deducible_Pesos, "
        mstrSQL = mstrSQL & " Total_Mecanica,Total_Carroceria,"
        mstrSQL = mstrSQL & " Total_Desabolladura,Total_Pintura,"
        mstrSQL = mstrSQL & " Total_Terceros,Total_Repuestos,"
        mstrSQL = mstrSQL & " Total_Materiales,Total_Insumos, "
        mstrSQL = mstrSQL & " Total_Otros,Total_Ot,"
        mstrSQL = mstrSQL & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor,"
        mstrSQL = mstrSQL & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto ) "
        mstrSQL = mstrSQL & " VALUES ("
        mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
        mstrSQL = mstrSQL & " '" & lblNroRecepcion & "', '" & gstrSeccion & "',"
        mstrSQL = mstrSQL & " '" & Trim(dtcGarantia.BoundText) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), "S/F") & "',"
        mstrSQL = mstrSQL & " '" & dtcTipoCono.BoundText & "', " & CLng(txtNroCono.Text) & ","
        mstrSQL = mstrSQL & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
        mstrSQL = mstrSQL & " " & CLng(txtKilAct) & ", '" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"   'OJO
        mstrSQL = mstrSQL & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
        mstrSQL = mstrSQL & " '" & IIf(Me.dtcGarantia.BoundText = "PRE", "P", "V") & "','" & CDate(pckFechaAtencion.Value) & "', "
        mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
        mstrSQL = mstrSQL & " '" & "S/N" & "', '" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "',"
        mstrSQL = mstrSQL & " '" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ','" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "','" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "' , "
        mstrSQL = mstrSQL & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), "S/S") & "' ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & " ," & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", " & gcurInsumo & ", "
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & ", " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & " ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & " ," & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ","
        mstrSQL = mstrSQL & " '" & lblIdCliente & "',"
        mstrSQL = mstrSQL & " '" & IIf(optMantencion.Value = True, "M", "R") & "',"
        mstrSQL = mstrSQL & " '" & IIf(cmdReserva.Enabled = False, "R", "N") & "',"
        mstrSQL = mstrSQL & " '" & mstrIdPresupuestoOrigen & "')"
    Else
        mstrSQL = "UPDATE Tllr_Presupuesto "
        mstrSQL = mstrSQL & " SET Id_Garantia='" & Trim(dtcGarantia.BoundText) & "', "
        mstrSQL = mstrSQL & " Folio_Garantia='" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), ".") & "', "
        mstrSQL = mstrSQL & " Id_Tipo_Cono='" & dtcTipoCono.BoundText & "', "
        mstrSQL = mstrSQL & " Nro_Cono=" & CLng(txtNroCono.Text) & ", "
        mstrSQL = mstrSQL & " Patente='" & txtPatente.Text & "', "
        mstrSQL = mstrSQL & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
        mstrSQL = mstrSQL & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
        mstrSQL = mstrSQL & " Entrega_Estimada='" & CDate(pckFechaEntrega) & "', "
        mstrSQL = mstrSQL & " Hora_Entrega='" & cboHora.Text & "', "
        mstrSQL = mstrSQL & " Nro_Siniestro='" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ', "
        mstrSQL = mstrSQL & " Nro_Poliza='" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "', "
        mstrSQL = mstrSQL & " Liquidador='" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "', "
        mstrSQL = mstrSQL & " Comentario='" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), ".") & "', "
        mstrSQL = mstrSQL & " Solicitado_Por='" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), ".") & "',"
        mstrSQL = mstrSQL & " Total_Mecanica=" & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Carroceria=" & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " Total_Desabolladura=" & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Pintura=" & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " Total_Terceros=" & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Repuestos=" & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " Total_Otros=" & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & "  ,"
        mstrSQL = mstrSQL & " Total_Materiales=" & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Insumos=" & gcurInsumo & ", "
        mstrSQL = mstrSQL & " Total_Ot=" & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo & "  ,"
        mstrSQL = mstrSQL & " Total_OT_Iva=" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & "  ,"
        mstrSQL = mstrSQL & " Total_IVA =" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & "  ,"
        mstrSQL = mstrSQL & " Deducible_UF = " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , "
        mstrSQL = mstrSQL & " Deducible_Pesos = " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
        mstrSQL = mstrSQL & " Nro_Presupuesto_Origen='" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "', "
        mstrSQL = mstrSQL & " Kilometros_Recepcion=" & CLng(txtKilAct) & ","
        mstrSQL = mstrSQL & " Id_Compañia_Seguro='" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"
        mstrSQL = mstrSQL & " Fecha_Proxima_Visita = '" & DateAdd("d", 365, pckFechaAtencion.Value) & "',"
        mstrSQL = mstrSQL & " Id_Cliente_Proveedor='" & lblIdCliente & "',"
        mstrSQL = mstrSQL & " ReparacionMantencion='" & IIf(Me.optMantencion.Value = True, "M", "R") & "',"
        mstrSQL = mstrSQL & " Estado_Reserva='" & IIf(Me.cmdReserva.Enabled = False, "R", "N") & "',"
        mstrSQL = mstrSQL & " Id_Presupuesto='" & mstrIdPresupuestoOrigen & "'"
        mstrSQL = mstrSQL & " WHERE Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal ='" & gstrIdSucursal & "' And Id_OT ='" & Trim(Trim(lblNroRecepcion)) & "' AND Seccion_OT ='" & gstrSeccion & "' "
    End If                                                                                                                                                                                                                                                                              ''" & pckFechaVenta.Value & "'
    
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        '/////////////////////////////// AQUI GUARDAR DATOS DEL VEHICULO
            mstrSQL = " Update Tllr_Vehiculo_Cliente "
            mstrSQL = mstrSQL & " Set Kilometros_Actuales = " & IIf(Trim(txtKilAct) <> "", CLng(txtKilAct), 0) & " , "
            mstrSQL = mstrSQL & " Concesionario='" & IIf(Trim(txtConcesionario) <> "", UCase(Trim(txtConcesionario)), "S/C") & "' ,"
            mstrSQL = mstrSQL & " Fecha_Venta='" & pckFecVta.Value & "'"
            mstrSQL = mstrSQL & " Where Patente='" & txtPatente & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
            MsgBox LoadResString(323)
        End If
        If GuardaInventario(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
            MsgBox LoadResString(322)
        End If
        If GuardaMecanica(lblNroRecepcion, gcOrdenTrabajo) = False Then
            MsgBox LoadResString(321)
        End If
        'If gstrSeccion = "C" Then
            If GuardaCarroceria(lblNroRecepcion, gstrSeccion, lblCompañia.Tag, gcOrdenTrabajo) = False Then
                MsgBox LoadResString(320)
            End If
        'End If
        If GuardaOtros(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
            MsgBox LoadResString(328)
        End If
        If GuardaTerceros(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
            MsgBox LoadResString(319)
        End If
        If GuardaRepuestos(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
            MsgBox LoadResString(318)
        End If
        
'//////////////////////////////////

        'actualiza datos de rent a car
        If Me.dtcGarantia.BoundText = "REN" And Me.optMantencion.Value = True Then
            gstrEstadoMantencion = Retorna_Valor_General("Select EstadoMantencion from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
            gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoMantencion & "'"
            gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
            If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
            End If
        End If
        If Me.dtcGarantia.BoundText = "REN" And Me.optReparacion.Value = True Then
            gstrEstadoReparacion = Retorna_Valor_General("Select EstadoReparacion from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
            gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoReparacion & "'"
            gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
            If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
            End If
        End If

'//////////////////////////////////
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
'//////////////////////////////////
        If lblEstadoOT.Visible = False Then
            If MsgBox("Imprimirá la OT Nº " & lblNroRecepcion & ", Confirma el Documento", 4 + 32, "Imprime OT(Recepción)") = vbYes Then
                PrintOT
                AgregarRegistro
            Else
                AgregarRegistro
            End If
        End If
    End If '//////////////
End Sub
Sub GrabarPresupuesto(NumeroPresupuesto As String, EstadoPresupuesto As String, MotivoAnula As String)

    mstrSQL = "INSERT INTO Tllr_Presupuesto "
    mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal, "
    mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
    mstrSQL = mstrSQL & " Id_Garantia, Folio_Garantia, "
    mstrSQL = mstrSQL & " Id_Tipo_Cono, Nro_Cono, "
    mstrSQL = mstrSQL & " Patente, RealizadoPor,"
    mstrSQL = mstrSQL & " Kilometros_Recepcion, Id_Compañia_seguro,"
    mstrSQL = mstrSQL & " Fecha_Proxima_Visita, "                           'Fecha_Liquidacion,"
    mstrSQL = mstrSQL & " Estado,Fecha_Emision, "
    mstrSQL = mstrSQL & " Entrega_Estimada, Hora_Entrega, "
    mstrSQL = mstrSQL & " Nro_Factura_Emitida,Nro_Presupuesto_Origen,"
    mstrSQL = mstrSQL & " Nro_Siniestro, Nro_Poliza, Liquidador, "
    mstrSQL = mstrSQL & " Comentario, Solicitado_Por,"
    mstrSQL = mstrSQL & " Deducible_UF , Deducible_Pesos, "
    mstrSQL = mstrSQL & " Total_Mecanica,Total_Carroceria,"
    mstrSQL = mstrSQL & " Total_Desabolladura,Total_Pintura,"
    mstrSQL = mstrSQL & " Total_Terceros,Total_Repuestos,"
    mstrSQL = mstrSQL & " Total_Materiales,Total_Insumos, "
    mstrSQL = mstrSQL & " Total_Otros,Total_Ot,"
    mstrSQL = mstrSQL & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor,"
    mstrSQL = mstrSQL & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto, Descripcion_Anula ) "
    mstrSQL = mstrSQL & " VALUES ("
    mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
    mstrSQL = mstrSQL & " '" & lblNroRecepcion & "', '" & gstrSeccion & "',"
    mstrSQL = mstrSQL & " '" & Trim(dtcGarantia.BoundText) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), "S/F") & "',"
    mstrSQL = mstrSQL & " '" & dtcTipoCono.BoundText & "', " & CLng(txtNroCono.Text) & ","
    mstrSQL = mstrSQL & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
    mstrSQL = mstrSQL & " " & CLng(txtKilAct) & ", '" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"   'OJO
    mstrSQL = mstrSQL & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
    mstrSQL = mstrSQL & " '" & EstadoPresupuesto & "','" & CDate(pckFechaAtencion.Value) & "', "
    mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
    mstrSQL = mstrSQL & " '" & "S/N" & "', '" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "',"
    mstrSQL = mstrSQL & " '" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ','" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "','" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "' , "
    mstrSQL = mstrSQL & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), "S/S") & "' ,"
    mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
    mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & " ," & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
    mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
    mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
    mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", " & gcurInsumo & ", "
    mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & ", " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & " ,"
    mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & " ," & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ","
    mstrSQL = mstrSQL & " '" & lblIdCliente & "',"
    mstrSQL = mstrSQL & " '" & IIf(optMantencion.Value = True, "M", "R") & "',"
    mstrSQL = mstrSQL & " '" & IIf(cmdReserva.Enabled = False, "R", "N") & "',"
    mstrSQL = mstrSQL & " '" & NumeroPresupuesto & "',"
    mstrSQL = mstrSQL & " '" & MotivoAnula & "')"
    
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        If GuardaInventario(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
            MsgBox LoadResString(322)
        End If
        If GuardaMecanica(NumeroPresupuesto, gcPresupuesto) = False Then
            MsgBox LoadResString(321)
        End If
        'If gstrSeccion = "C" Then
            If GuardaCarroceria(NumeroPresupuesto, gstrSeccion, lblCompañia.Tag, gcPresupuesto) = False Then
                MsgBox LoadResString(320)
            End If
        'End If
        If GuardaOtros(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
            MsgBox LoadResString(328)
        End If
        If GuardaTerceros(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
            MsgBox LoadResString(319)
        End If
        If GuardaRepuestos(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
            MsgBox LoadResString(318)
        End If
        
'//////////////////////////////////
        mblnTablaVacia = False
    End If '//////////////

End Sub
Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        '////////////////////////////////ELIMINAR SERVICIOS DE MECANICA///////////////////////////////////
        mstrSQL = "DELETE FROM Tllr_Mecanica_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        Conexion.SendHost mstrSQL, , , , gcTiempoEspera
        '////////////////////////////////ELIMINAR SERVICIOS DE CARRPCERIA///////////////////////////////////
        mstrSQL = "DELETE FROM Tllr_Carroceria_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        Conexion.SendHost mstrSQL, , , , gcTiempoEspera
        '////////////////////////////////////ELIMINAR INENTARIO///////////////////////////////
        mstrSQL = "DELETE FROM Tllr_Inventario_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        Conexion.SendHost mstrSQL, , , , gcTiempoEspera
        '//////////////////////////////////////ENCABEZADO/////////////////////////////
        mstrSQL = "DELETE FROM Tllr_Presupuesto WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
            mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.Id_OT > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_OT"
            gstrSql = letSql(mstrWhere, mstrOrderBy)
            If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.Id_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_OT"
                    gstrSql = letSql(mstrWhere, mstrOrderBy)
                    
                    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                            LeerCampos
                        Else
                            mblnTablaVacia = True
                            LimpiaCampos
                        End If
                    End If
                End If
            End If
        End If
        Conexion.CloseHost AdoPrincipal
    End If
End Sub
Private Sub BuscarRegistro()
Screen.MousePointer = 1
frmBuscaPresupuesto.Show vbModal
Screen.MousePointer = 1
If gstrBusca <> "" Then
    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.ID_Presupuesto=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_OT"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End If
Me.SetFocus

End Sub
Private Sub PrimerRegistro()
    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_Presupuesto"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroAnterior()
    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.Id_Presupuesto < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_Presupuesto DESC"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroSiguiente()
    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.Id_Presupuesto > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_Presupuesto "
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub UltimoRegistro()
    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_Presupuesto DESC"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub Renovar()
    mstrWhere = " WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_Presupuesto.Id_Presupuesto "
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        VerificaTablaVacia
        ActivaBotones
        If Not mblnTablaVacia Then
            PrimerRegistro
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub ActivaBotones()
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
        .Item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoEditar, True, False))
        .Item("Cancelar").Enabled = False
        .Item("Buscar").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Imprimir").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoImprimir, True, False))
        .Item("Primero").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Anterior").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Siguiente").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Ultimo").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Renovar").Enabled = True
        .Item("Cerrar").Enabled = True
    End With
End Sub
Private Sub DesactivaBotones()
With tlbBarraHerramientas.Buttons
    .Item("Crear").Enabled = False
    .Item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
    .Item("Cancelar").Enabled = True
    .Item("Buscar").Enabled = False
    .Item("Imprimir").Enabled = False
    .Item("Primero").Enabled = False
    .Item("Anterior").Enabled = False
    .Item("Siguiente").Enabled = False
    .Item("Ultimo").Enabled = False
    .Item("Renovar").Enabled = False
    .Item("Cerrar").Enabled = True
    .Item("Activar").Enabled = False
    .Item("Liquidar").Enabled = False
    .Item("Anular").Enabled = False
End With
End Sub
Private Sub VerificaTablaVacia()
    If (Not AdoPrincipal.BOF And Not AdoPrincipal.EOF) And AdoPrincipal.RecordCount > 0 Then
        mblnTablaVacia = False
    Else
        mblnTablaVacia = True
        LimpiaCampos
    End If
End Sub

Private Sub LimpiaCampos()
With Me
    SetCheckOff .lvwInventario
    .lvwServiciosCarroceria.ListItems.Clear
    .lvwServiciosMecanica.ListItems.Clear
    .lvwServiciosTerceros.ListItems.Clear
    .lvwRepuestos.ListItems.Clear
    .lblNroRecepcion.Text = ""
    .dtcGarantia.BoundText = ""
    .dtcGarantia.Enabled = True
    .pckFechaAtencion.Value = Now
    .txtPatente.Text = ""
    .lblMarca.Caption = "": .lblIdMarca = ""
    .lblModelo.Caption = "": .lblIdModelo = ""
    .txtAño.Text = ""
    .lblColorE.Caption = ""
    .lblChasis.Caption = ""
    .lblMotor.Caption = ""
    .lblCliente.Caption = ""
    .txtKilAct.Text = ""
    .txtConcesionario.Text = ""
    .pckFecVta.Value = Now
    .dtcTipoCono.BoundText = ""
    .txtNroCono.Text = ""
    .dtcRecepcionista.BoundText = ""
    .pckFechaEntrega.Value = Now
    .cboHora.Text = ""
    .lblCompañia.Caption = ""
    .lblCompañia.Tag = ""
    .txtDeducibleUF.Text = "0"
    .txtDeduciblePesos.Text = "0"
    .txtNroSiniestro.Text = ""
    .txtNroPoliza.Text = ""
    .txtLiquidador.Text = ""
    .lblFono.Caption = ""
    .lblVin.Caption = ""
    .txtSolicita.Text = ""
    .txtFolioGarantia.Text = ""
    .txtRut.Text = ""
    .txtComuna.Text = ""
    .txtDir.Text = ""
    .lblIdCliente.Caption = ""
    .txtComentario = ""
    .cmdAnularReserva.Enabled = False
    .cmdReserva.Enabled = True
    .lblPresupuesto = ""
End With
End Sub
Private Sub ValoresporDefecto()
    txtAño.Text = Year(Now)
    txtDeducibleUF.Text = "0"
    txtNroCono.Text = "0"
    txtDeduciblePesos.Text = "0"
    txtNroSiniestro.Text = "."
    txtNroPoliza.Text = "."
    txtLiquidador.Text = "."
    txtKilAct.Text = "0"
    lblEstadoOTValor = "VIGENTE"
    lblEstadoOTValor.Tag = "V"
End Sub
Private Function validacion() As Boolean
    validacion = True
With Me
    If .dtcGarantia.BoundText = "" Then
        MsgBox LoadResString(317), vbInformation, "Advertencia"
        dtcGarantia.Enabled = True
        dtcGarantia.SetFocus
        validacion = False
        Exit Function
    End If
    If .txtPatente = "" Then
        MsgBox LoadResString(316), vbInformation, "Advertencia"
        txtPatente.SetFocus
        validacion = False
        Exit Function
    Else
        If ExistePatente(txtPatente) = False Then
            MsgBox LoadResString(329), vbInformation, "Advertencia"
            txtPatente.SetFocus
            validacion = False
            Exit Function
        End If
    End If
    If gstrSeccion = "C" Then
        If .txtFolioGarantia = "" Then
            MsgBox LoadResString(315), vbInformation, "Advertencia"
            txtFolioGarantia.SetFocus
            validacion = False
            Exit Function
        End If
    End If
    If .txtSolicita = "" Then
        MsgBox LoadResString(314), vbInformation, "Advertencia"
        txtSolicita.SetFocus
        validacion = False
        Exit Function
    End If
'    If .txtKilAct = "" Then
'        MsgBox LoadResString(313), vbInformation, "Advertencia"
'        txtKilAct.SetFocus
'        Validacion = False
'        Exit Function
'    End If
    If (CDbl(.txtKilAct) = 0 Or .txtKilAct = "") Or (CDbl(.txtKilAct) <= KilometrajeEntrada) Then
        If UCase(Me.Tag) = "CREAR" And Me.lblEstadoOTValor <> "RESERVA" And gstrKmsAutoNuevo <> "Nuevo" And Me.dtcGarantia.BoundText <> "PRE" And mstrLiquidaPresupuesto = False Then
            'MsgBox LoadResString(313), vbInformation, "Advertencia"
            MsgBox "El Kilometraje de la última visita fué de " & CDbl(.txtKilAct) & Chr(13) & "Verifique el kilometraje ingresado...", vbInformation, "Advertencia"
            Me.Frame3.Enabled = True
            Me.Frame4.Enabled = True
            Me.Frame8.Enabled = True
            txtKilAct.Enabled = True
            'txtKilAct.SetFocus
            validacion = False
            Exit Function
        End If
    End If

    If .dtcTipoCono.BoundText = "" Then
        MsgBox LoadResString(312), vbInformation, "Advertencia"
        dtcTipoCono.SetFocus
        validacion = False
        Exit Function
    End If
    If .txtNroCono = "" Then
        MsgBox LoadResString(311), vbInformation, "Advertencia"
        txtNroCono.SetFocus
        validacion = False
        Exit Function
    End If
    If .dtcRecepcionista.BoundText = "" Then
        MsgBox LoadResString(310), vbInformation, "Advertencia"
        dtcRecepcionista.SetFocus
        validacion = False
        Exit Function
    End If
    If dtcGarantia.BoundText = "REN" Then
        If Me.optMantencion.Value = False And Me.optReparacion.Value = False Then
            MsgBox "Para Rent a Car Debe elegir Reparación o Mantención", vbInformation, "Advertencia"
            dtcGarantia.SetFocus
            validacion = False
            Exit Function
         End If
    End If
    If .optRecepcion(1).Value = True Then
        If .txtDeducibleUF.Text = "" Then
            MsgBox LoadResString(308), vbInformation, "Advertencia"
            txtDeducibleUF.SetFocus
            validacion = False
            Exit Function
        End If
        If .txtDeduciblePesos.Text = "" Then
            MsgBox LoadResString(307), vbInformation, "Advertencia"
            txtDeduciblePesos.SetFocus
            validacion = False
            Exit Function
        End If
        If .txtNroSiniestro.Text = "" Then
            MsgBox LoadResString(306), vbInformation, "Advertencia"
            txtNroSiniestro.SetFocus
            validacion = False
            Exit Function
        End If
        If .txtNroPoliza.Text = "" Then
            MsgBox LoadResString(305), vbInformation, "Advertencia"
            txtNroPoliza.SetFocus
            validacion = False
            Exit Function
        End If
        If .txtLiquidador.Text = "" Then
            MsgBox LoadResString(304), vbInformation, "Advertencia"
            txtLiquidador.SetFocus
            validacion = False
            Exit Function
        End If
    End If
    '//////////////////////////////////CARROCERIA
End With
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" And Me.lblEstadoOTValor <> "RESERVA" And Me.lblEstadoOTValor <> "PRESUPUESTO" Then
        Dim adoTemp As New ADODB.Recordset
        mstrSQL = "select ID_OT from Tllr_Presupuesto where SECCION_OT = '" & gstrSeccion & "' AND ID_OT ='" & lblNroRecepcion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción "
                validacion = False
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmRecepcion = Nothing
    gstrBusca = lblNroRecepcion.Text
End Sub
Private Sub RevizaAtributos()
    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True
End Sub

Private Sub tlbCiaSeg_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case Is = "Nueva"
    gstrProcedencia = "Movimientos"
    frmMantenedorCompañiaSeguro.Show 1
    
Case Is = "Buscar"
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Compañia_Seguro", "Id_Compañia_Seguro", "Nombre", "Busca Compañia de Seguro")
    gstrBusca = ""
    frmBuscarCiaSeguros.Show vbModal
    lblCompañia = NombreCiaSeg(gstrBusca)
    lblCompañia.Tag = gstrBusca
    FillConceptosVsCiaSeguro dtcConceptos, datConceptos, lblCompañia.Tag
    txtDeduciblePesos.SetFocus
End Select

End Sub

Private Sub tlbPatente_ButtonClick(ByVal Button As MSComctlLib.Button)

If Me.Tag = "Crear" Then
    Select Case Button.Key
    Case "Nuevo"
        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apcrear)
        DatosVehiculo txtPatente
    Case "Buscar"
        gstrProcedencia = "Movimientos"
        frmBuscaVehiculo.Show vbModal
    End Select
Else
    Select Case Button.Key
    Case "Nuevo"
        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apeditar)
        DatosVehiculo txtPatente
    End Select
End If

End Sub


Private Sub tlbPatente_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If tlbPatente.Buttons(1).Key = "Nuevo" Then
    tlbPatente.Buttons(1).ToolTipText = IIf(Me.Tag = "Crear", "Nuevo Vehiculo", "Editar Vehiculo")
Else
    tlbPatente.Buttons(2).ToolTipText = IIf(Me.Tag = "Crear", "Buscar Vehiculo", "Buscar Vehiculo")
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub txtConcesionario_GotFocus()
MarcaTexto txtConcesionario
End Sub

Private Sub txtDeduciblePesos_GotFocus()
MarcaTexto txtDeduciblePesos
End Sub


Private Sub txtDeducibleUF_GotFocus()
MarcaTexto txtDeducibleUF

End Sub


Private Sub txtFolioGarantia_GotFocus()
MarcaTexto txtFolioGarantia
End Sub

Private Sub txtKilAct_GotFocus()
MarcaTexto txtKilAct
End Sub

Private Sub txtNroCono_GotFocus()
MarcaTexto txtNroCono
End Sub

Private Sub txtPatente_GotFocus()
MarcaTexto txtPatente
End Sub

Private Sub txtPatente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MyRecordset As New ADODB.Recordset
'If Me.Tag = "Crear" Then
Dim str1 As String
Dim str2 As String
    If KeyCode = 13 Then
''       kjcv 24 - 01 - 12
''       CheckPatente txtPatente, str1, str2  '/// devuelve el rut de la patente
'        txtFolioGarantia = str2
        If txtPatente <> "" Then
            If Len(txtPatente) = 6 And lblPat.Caption = gstrNombrePatente Then
               ' If str2 = "-0" Then         '//// valida la patente
               '     MsgBox "Patente no Valida, Ingrese de nuevo la Patente", vbInformation, "Ingreso de Patente"
               '     Exit Sub
               ' End If
                
                If dtcGarantia.BoundText = "VHP" Then  '/// valida patente vehiculos propios
                    If ConsultaVehiculoPropio(txtPatente) = False Then
                        MsgBox gstrNombrePatente & " no EXISTE en Vehiculos Propios", vbInformation, "Ingreso de " & gstrNombrePatente
                        Exit Sub
                    End If
                End If
                
                If dtcGarantia.BoundText = "REN" Then
                    Set MyRecordset = cnnAux.Execute("EXEC RENT_ACTUALIZA_VEHICULO_CLIENTE '" & Me.txtPatente & "', '" & gstrIdUsuario & "', '" & Date & "'")
                End If
                
                If ConsultaVehiculo(txtPatente) = True Then
                    If MsgBox("La " & gstrNombrePatente & " " & txtPatente & " Ya Existe, Desea Desplegar los Datos", 4 + 32, gstrNombrePatente & " Existente") = vbYes Then
                        Call DatosVehiculo(txtPatente)
                    Else
                        LimpiaCampos
                    End If
                Else
                    gstrProcedencia = "Movimientos"
                    gapAccion = apcrear
                    gstrKmsAutoNuevo = "Nuevo"
                    frmMantenedorVehiculoCliente.Show vbModal
                End If
            
            ElseIf dtcGarantia.BoundText = "INW" Or dtcGarantia.BoundText = "INC" Then
                If ConsultaVinExistencia(txtPatente) = True Then
                    If ConsultaVehiculo(txtPatente) = True Then
                        If MsgBox("La Placa " & txtPatente & " Ya Existe, Desea Desplegar los Datos", 4 + 32, "Placa Existente") = vbYes Then
                            Call DatosVehiculo(txtPatente)
                        Else
                            LimpiaCampos
                        End If
                    Else
                        gstrProcedencia = "Movimientos"
                        gapAccion = apcrear
                        frmMantenedorVehiculoCliente.Show vbModal
                    End If
                End If
            Else
                MsgBox LoadResString(326)
            End If
            
        Else
            MsgBox LoadResString(327)
        End If
    End If
'End If
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'If Trim(lblPat.Caption) = gstrNombrePatente And dtcGarantia.BoundText <> "PEX" Then
'    If gstrValidaPatente = "S" Then
'        KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'    End If
'End If
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub

Private Sub txtSolicita_GotFocus()
MarcaTexto txtSolicita
End Sub
Private Sub ConfirmarReserva()
    Screen.MousePointer = vbHourglass
    gstrBuscaReserva = lblNroRecepcion
    Me.Tag = "Crear"
    If validacion() = True Then
        GrabarRegistro    '/// graba la ot de reserva en una ot definitiva
        
        '/// Asigna numero de ot a la reserva de hora
        mstrSQL = "Update Tllr_ReservaHora "
        mstrSQL = mstrSQL & " Set Id_OT = '" & gstrBusca & "'"
        mstrSQL = mstrSQL & " Where Id_Reserva='" & Mid(gstrBuscaReserva, 3, 5) & "'"
        mstrSQL = mstrSQL & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
            MsgBox "Error Al Actualizar Los Datos De La Reserva de Hora"
        End If
        
        '/// actualiza la reserva de repuestos
        mstrSQL = "Update Stck_Regularizacion "
        mstrSQL = mstrSQL & " Set Id_OT = '" & gstrBusca & "'"
        mstrSQL = mstrSQL & " Where Id_Ot='" & gstrSeccion & gstrBuscaReserva & "'"
        mstrSQL = mstrSQL & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
            MsgBox "Error Al Actualizar Los Datos De La Reserva de Repuestos"
        End If
        
        '/// actualiza repuestos reservados
        mstrSQL = "Update Tllr_Repuestos_Reservados "
        mstrSQL = mstrSQL & " Set Id_OT = '" & gstrBusca & "'"
        mstrSQL = mstrSQL & " Where Id_Ot='" & gstrBuscaReserva & "' And Seccion_OT='" & gstrSeccion & "'"
        mstrSQL = mstrSQL & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
            MsgBox "Error Al Actualizar Los Datos de los Repuestos Reservados"
        End If
        
        '/// actualiza repuestos faltantes
        mstrSQL = "Update Tllr_Repuestos_Faltantes "
        mstrSQL = mstrSQL & " Set Id_OT = '" & gstrBusca & "'"
        mstrSQL = mstrSQL & " Where Id_Ot='" & gstrBuscaReserva & "' And Seccion_OT='" & gstrSeccion & "'"
        mstrSQL = mstrSQL & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
            MsgBox "Error Al Actualizar Los Datos De Los Repuestos Faltantes"
        End If
    Else
        Exit Sub
    End If
    
    EliminaReserva gstrBuscaReserva          '/// elimina la ot de reserva que fue grabada anteriormente
    Screen.MousePointer = vbDefault
End Sub
Private Sub EliminaReserva(pstrNroReserva)
    '////////////////////////////////ELIMINAR SERVICIOS DE MECANICA///////////////////////////////////
    mstrSQL = "DELETE FROM Tllr_Mecanica_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '////////////////////////////////ELIMINAR OTROS SERVICIOS ///////////////////////////////////
    mstrSQL = "DELETE FROM Tllr_Presupuestoro_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '////////////////////////////////ELIMINAR SERVICIOS DE CARRPCERIA///////////////////////////////////
    mstrSQL = "DELETE FROM Tllr_Carroceria_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '////////////////////////////////ELIMINAR SERVICIOS DE TERCEROS///////////////////////////////////
    mstrSQL = "DELETE FROM Tllr_Terceros_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '////////////////////////////////////ELIMINAR INENTARIO///////////////////////////////
    mstrSQL = "DELETE FROM Tllr_Inventario_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    mstrSQL = "DELETE FROM Tllr_Repuestos_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '//////////////////////////////////////ENCABEZADO/////////////////////////////
    mstrSQL = "DELETE FROM Tllr_Presupuesto WHERE Tllr_Presupuesto.Seccion_OT = '" & gstrSeccion & "' AND Tllr_Presupuesto.Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    
End Sub
Private Sub EliminaReservaRepuestos(pstrNroRegularizacion As String, pstrNroOT As String)
    
    '//// Elimina Detalle de la Reserva de Repuestos
    mstrSQL = "DELETE FROM Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & pstrNroRegularizacion & "' AND Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '//// Elimina Cabezera de la Reserva de Repuestos
    mstrSQL = "DELETE FROM Stck_Regularizacion Where Id_Regularizacion = '" & pstrNroRegularizacion & "' AND Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '//// Elimina Repuestos Reservados
    mstrSQL = "DELETE FROM Tllr_Repuestos_Reservados WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & pstrNroOT & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
    '//// Elimina Repuestos Faltantes
    mstrSQL = "DELETE FROM Tllr_Repuestos_Faltantes WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & pstrNroOT & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
    
End Sub

Private Sub CancelaReserva()
Dim mstrMotivoCancela As String

    If MsgBox("¿ Realmente Desea eliminar esta Reserva de Hora?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        If TieneReservadeRepuestos Then
            NroRegularizacion = Retorna_Valor_General("Select Id_Regularizacion as Numero from Stck_Regularizacion where id_ot='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
            Call Actualiza_Saldos_VS_Detalle("S", "Select Canrtidad, Id_Empresa, Id_sucursal, Id_Bodega,Id_Ubicacion,Id_Item From Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & NroRegularizacion & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Empresa = '" & gstrIdEmpresa & "'")
            EliminaReservaRepuestos NroRegularizacion, lblNroRecepcion  'Regularizacion
        End If
        EliminaReserva lblNroRecepcion   'OT
    Else
        Exit Sub
    End If
    
    '/// ingresa motivo de cancelación
    mstrMotivoCancela = InputBox("Ingrese el motivo de Cancelacion de la Reserva", "Por que Cancela...")
    
    '/// Actualiza Estado Reserva de Hora
    mstrSQL = " Update Tllr_ReservaHora "
    mstrSQL = mstrSQL & " Set Estado = 'E',"
    mstrSQL = mstrSQL & " Fecha_Cancelacion='" & Date & "',"
    mstrSQL = mstrSQL & " Quien_Cancela='" & gstrUsuario & "',"
    mstrSQL = mstrSQL & " MotivoCancela='" & mstrMotivoCancela & "'"
    mstrSQL = mstrSQL & " Where Id_Reserva='" & Mid(lblNroRecepcion, 3, 5) & "'"
    mstrSQL = mstrSQL & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
        MsgBox "Error Al Actualizar Los Datos De La Reserva de Hora"
    End If
    
    CancelarAgregaRegistro
    
End Sub
Sub Repuestos_de_la_Mantencion(stridMarca As String, stridModelo As String, stridServicio As String, blnLlenaRepuestos As Boolean)
Dim adoTemp As New ADODB.Recordset
    
'lvwRepuestos.ListItems.Clear
mstrSQL = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
mstrSQL = mstrSQL & " Stck_Item.Descripcion AS NOMBRE, "
mstrSQL = mstrSQL & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
mstrSQL = mstrSQL & " Tllr_Actividad_Repuesto.Valor AS VLR, "
mstrSQL = mstrSQL & " Stck_Item.Id_Familia AS IDFAM, "
mstrSQL = mstrSQL & " Glbl_Familia.Descripcion AS FAMILIA "
mstrSQL = mstrSQL & " FROM Glbl_Familia RIGHT OUTER JOIN Stck_Item ON  Glbl_Familia.Id_Familia = Stck_Item.Id_Familia RIGHT OUTER JOIN Tllr_Actividad_Repuesto ON Stck_Item.Id_Item = Tllr_Actividad_Repuesto.Id_Item"
mstrSQL = mstrSQL & " WHERE Tllr_Actividad_Repuesto.Id_Marca = '" & stridMarca & "' AND Tllr_Actividad_Repuesto.Id_Modelo = '" & stridModelo & "' AND Tllr_Actividad_Repuesto.Id_Servicio = '" & stridServicio & "' "
    
    
    If Conexion.SendHost(mstrSQL, adoTemp, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoTemp
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwRepuestosMantencion.ListItems.Add(, , ValorNulo(!Codigo))
                    lsiItem.SubItems(1) = ValorNulo(!Nombre)
                    lsiItem.SubItems(2) = FormatoValor(!CANTY, "", 1)
                    lsiItem.SubItems(3) = FormatoValor(!VLR, "", gintDecimalesMoneda)
                    lsiItem.SubItems(4) = ValorNulo(!Familia)
                    lsiItem.SubItems(5) = Me.lvwServiciosMecanica.SelectedItem.SubItems(6)
                    
                    If Me.dtcGarantia.BoundText = "PRE" Then
                        Set itmAux = lvwRepuestos.ListItems.Add(, , ValorNulo(!Codigo))
                        itmAux.SubItems(1) = ValorNulo(!Nombre)
                        itmAux.SubItems(2) = FormatoValor(!CANTY, "", 1)
                        itmAux.SubItems(3) = FormatoValor(!VLR, "", gintDecimalesMoneda)
                        itmAux.SubItems(4) = FormatoValor(0, "", 2)
                        itmAux.SubItems(5) = FormatoValor(0, "", gintDecimalesMoneda)
                        itmAux.SubItems(6) = "" 'TraeCargoDes(gstrIdCargo)
                        itmAux.SubItems(7) = gstrIdCargo
                        itmAux.SubItems(8) = Format(Val(SacarFormatoValor(itmAux.SubItems(2), "")) * Val(SacarFormatoValor(itmAux.SubItems(3), "")), "###,##0.0")
                        itmAux.SubItems(9) = ValorNulo(!IDFAM)
                        itmAux.SubItems(10) = "N"
                        itmAux.SubItems(11) = "PRESUPUESTOS"
                    End If
                    
                    .MoveNext
                Wend
                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                TotalFinal
            End If
        End With
    End If
    Conexion.CloseHost adoTemp
End Sub
Sub Quita_Repuestos_Mantencion(stridMarca As String, stridModelo As String, stridServicio As String)
Dim i As Integer

'lvwRepuestos.ListItems.Clear
mstrSQL = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
mstrSQL = mstrSQL & " Stck_Item.Descripcion AS NOMBRE, "
mstrSQL = mstrSQL & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
mstrSQL = mstrSQL & " Tllr_Actividad_Repuesto.Valor AS VLR, "
mstrSQL = mstrSQL & " Stck_Item.Id_Familia AS IDFAM, "
mstrSQL = mstrSQL & " Glbl_Familia.Descripcion AS FAMILIA "
mstrSQL = mstrSQL & " FROM Glbl_Familia RIGHT OUTER JOIN Stck_Item ON  Glbl_Familia.Id_Familia = Stck_Item.Id_Familia RIGHT OUTER JOIN Tllr_Actividad_Repuesto ON Stck_Item.Id_Item = Tllr_Actividad_Repuesto.Id_Item"
mstrSQL = mstrSQL & " WHERE Tllr_Actividad_Repuesto.Id_Marca = '" & stridMarca & "' AND Tllr_Actividad_Repuesto.Id_Modelo = '" & stridModelo & "' AND Tllr_Actividad_Repuesto.Id_Servicio = '" & stridServicio & "' "
    
    
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set itmAux = lvwRepuestos.FindItem(!Codigo, lvwText, , 0)
                    If Not itmAux Is Nothing Then   ' Si no hay coincidencia
                        lvwRepuestos.ListItems.Remove (lvwRepuestos.FindItem(!Codigo).Index)
                        
                    End If
                    .MoveNext
                Wend
                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                TotalFinal
            End If
        End With
    End If
    
End Sub

Function TieneReservadeRepuestos() As Boolean
Dim lstrEstadoReserva As String

    TieneReservadeRepuestos = False

    lstrEstadoReserva = Retorna_Valor_General("Select estado_reserva As Codigo From Tllr_Presupuesto where id_ot='" & Me.lblNroRecepcion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Seccion_OT='" & gstrSeccion & "'")
    If lstrEstadoReserva = "R" Then
        TieneReservadeRepuestos = True
    End If
End Function
Sub GrabaReservaRepuestosRecepcion()
    If Me.Tag = "Crear" Then
        lblNroRecepcion = TraeCorrelativo(gcOrdenTrabajo, gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
        gstrBusca = lblNroRecepcion
        mstrSQL = "INSERT INTO Tllr_Presupuesto "
        mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal, "
        mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
        mstrSQL = mstrSQL & " Id_Garantia, Folio_Garantia, "
        mstrSQL = mstrSQL & " Id_Tipo_Cono, Nro_Cono, "
        mstrSQL = mstrSQL & " Patente, RealizadoPor,"
        mstrSQL = mstrSQL & " Kilometros_Recepcion, Id_Compañia_seguro,"
        mstrSQL = mstrSQL & " Fecha_Proxima_Visita, "                           'Fecha_Liquidacion,"
        mstrSQL = mstrSQL & " Estado,Fecha_Emision, "
        mstrSQL = mstrSQL & " Entrega_Estimada, Hora_Entrega, "
        mstrSQL = mstrSQL & " Nro_Factura_Emitida,Nro_Presupuesto_Origen,"
        mstrSQL = mstrSQL & " Nro_Siniestro, Nro_Poliza, Liquidador, "
        mstrSQL = mstrSQL & " Comentario, Solicitado_Por,"
        mstrSQL = mstrSQL & " Deducible_UF , Deducible_Pesos, "
        mstrSQL = mstrSQL & " Total_Mecanica,Total_Carroceria,"
        mstrSQL = mstrSQL & " Total_Desabolladura,Total_Pintura,"
        mstrSQL = mstrSQL & " Total_Terceros,Total_Repuestos,"
        mstrSQL = mstrSQL & " Total_Materiales,Total_Insumos, "
        mstrSQL = mstrSQL & " Total_Otros,Total_Ot,"
        mstrSQL = mstrSQL & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor, ReparacionMantencion, Estado_Reserva ) "
        mstrSQL = mstrSQL & " VALUES ("
        mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
        mstrSQL = mstrSQL & " '" & lblNroRecepcion & "', '" & gstrSeccion & "',"
        mstrSQL = mstrSQL & " '" & Trim(dtcGarantia.BoundText) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), "S/F") & "',"
        mstrSQL = mstrSQL & " '" & dtcTipoCono.BoundText & "', " & CLng(txtNroCono.Text) & ","
        mstrSQL = mstrSQL & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
        mstrSQL = mstrSQL & " " & CLng(txtKilAct) & ", '" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"   'OJO
        mstrSQL = mstrSQL & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
        mstrSQL = mstrSQL & " 'V','" & CDate(pckFechaAtencion.Value) & "', "
        mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
        mstrSQL = mstrSQL & " '" & "S/N" & "', '" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "',"
        mstrSQL = mstrSQL & " '" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ','" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "','" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "' , "
        mstrSQL = mstrSQL & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), "S/S") & "' ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & " ," & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", " & gcurInsumo & ", "
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & ", " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & " ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & " ," & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ","
        mstrSQL = mstrSQL & " '" & lblIdCliente & "',"
        mstrSQL = mstrSQL & " '" & IIf(optMantencion.Value = True, "M", "R") & "',"
        mstrSQL = mstrSQL & " '" & IIf(cmdReserva.Enabled = False, "R", "N") & "')"
    Else
        mstrSQL = "UPDATE Tllr_Presupuesto "
        mstrSQL = mstrSQL & " SET Id_Garantia='" & Trim(dtcGarantia.BoundText) & "', "
        mstrSQL = mstrSQL & " Folio_Garantia='" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), ".") & "', "
        mstrSQL = mstrSQL & " Id_Tipo_Cono='" & dtcTipoCono.BoundText & "', "
        mstrSQL = mstrSQL & " Nro_Cono=" & CLng(txtNroCono.Text) & ", "
        mstrSQL = mstrSQL & " Patente='" & txtPatente.Text & "', "
        mstrSQL = mstrSQL & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
        mstrSQL = mstrSQL & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
        mstrSQL = mstrSQL & " Entrega_Estimada='" & CDate(pckFechaEntrega) & "', "
        mstrSQL = mstrSQL & " Hora_Entrega='" & cboHora.Text & "', "
        mstrSQL = mstrSQL & " Nro_Siniestro='" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ', "
        mstrSQL = mstrSQL & " Nro_Poliza='" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "', "
        mstrSQL = mstrSQL & " Liquidador='" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "', "
        mstrSQL = mstrSQL & " Comentario='" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), ".") & "', "
        mstrSQL = mstrSQL & " Solicitado_Por='" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), ".") & "',"
        mstrSQL = mstrSQL & " Total_Mecanica=" & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Carroceria=" & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " Total_Desabolladura=" & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Pintura=" & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " Total_Terceros=" & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Repuestos=" & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " Total_Otros=" & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & "  ,"
        mstrSQL = mstrSQL & " Total_Materiales=" & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", "
        mstrSQL = mstrSQL & " Total_Insumos=" & gcurInsumo & ", "
        mstrSQL = mstrSQL & " Total_Ot=" & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo & "  ,"
        mstrSQL = mstrSQL & " Total_OT_Iva=" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & "  ,"
        mstrSQL = mstrSQL & " Total_IVA =" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & "  ,"
        mstrSQL = mstrSQL & " Deducible_UF = " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , "
        mstrSQL = mstrSQL & " Deducible_Pesos = " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
        mstrSQL = mstrSQL & " Nro_Presupuesto_Origen='" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "', "
        mstrSQL = mstrSQL & " Kilometros_Recepcion=" & CLng(txtKilAct) & ","
        mstrSQL = mstrSQL & " Id_Compañia_Seguro='" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"
        mstrSQL = mstrSQL & " Fecha_Proxima_Visita = '" & DateAdd("d", 365, pckFechaAtencion.Value) & "',"
        mstrSQL = mstrSQL & " Id_Cliente_Proveedor='" & lblIdCliente & "',"
        mstrSQL = mstrSQL & " ReparacionMantencion='" & IIf(Me.optMantencion.Value = True, "M", "R") & "',"
        mstrSQL = mstrSQL & " Estado_Reserva='" & IIf(Me.cmdReserva = False, "R", "N") & "'"
        mstrSQL = mstrSQL & " WHERE Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal ='" & gstrIdSucursal & "' And Id_OT ='" & Trim(Trim(lblNroRecepcion)) & "' AND Seccion_OT ='" & gstrSeccion & "' "
    End If                                                                                                                                                                                                                                                                              ''" & pckFechaVenta.Value & "'
    
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        If GuardaMecanica(lblNroRecepcion, gcOrdenTrabajo) = False Then
            MsgBox LoadResString(321)
        End If

'//////////////////////////////////
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
End Sub
Sub LiquidarPresupuesto()
    Screen.MousePointer = vbHourglass
    mstrLiquidaPresupuesto = True
    gstrBuscaReserva = lblNroRecepcion
    Me.Tag = "Crear"
    dtcGarantia.BoundText = "NGN"
    GrabarRegistro                                  '/// graba el presupuesto en una ot definitiva
    GrabarPresupuesto gstrBuscaReserva, "L", ""    '/// Graba presupuesto en tablas de presupuesto
    EliminaReserva gstrBuscaReserva                 '/// elimina el presupuesto que fue grabado anteriormente como OT
    '/// activa ot normal
    Me.dtcGarantia.Enabled = True
    Me.lblEstadoOTValor = "VIGENTE"
    If gstrImpresion = "O" Then
        tlbBarraHerramientas.Buttons.Item(14).Enabled = True
        tlbBarraHerramientas.Buttons.Item(15).Enabled = True
        tlbBarraHerramientas.Buttons.Item(24).Enabled = False
        tlbBarraHerramientas.Buttons.Item(25).Enabled = False
    End If
End Sub
Sub AnularPresupuesto()
Dim mstrMotivoAnula As String

    Screen.MousePointer = vbHourglass
    mstrMotivoAnula = InputBox("Ingrese El Motivo por que Anula :", "Por que Anula Presupuesto....")
    gstrBuscaReserva = lblNroRecepcion
    GrabarPresupuesto gstrBuscaReserva, "N", mstrMotivoAnula     '/// Graba presupuesto en tablas de presupuesto
    EliminaReserva gstrBuscaReserva          '/// elimina el presupuesto que fue grabado anteriormente como OT
    Renovar
    
End Sub
