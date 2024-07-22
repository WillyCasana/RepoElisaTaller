VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecepcion2 
   Caption         =   "Generación OT"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12765
   Icon            =   "frmRecepcion2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11970
      Begin VB.TextBox lblNroRecepcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   2100
      End
      Begin VB.TextBox txtTipo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10800
         TabIndex        =   1
         Top             =   165
         Width           =   975
      End
      Begin MSComCtl2.DTPicker pckFechaAtencion 
         Height          =   315
         Left            =   5070
         TabIndex        =   3
         Top             =   165
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93585409
         CurrentDate     =   36776
      End
      Begin VB.Label lblCorrelativo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción Nº :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Atención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3660
         TabIndex        =   7
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label lblEstadoOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado OT:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6420
         TabIndex        =   6
         Top             =   240
         Width           =   1035
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
         Left            =   7545
         TabIndex        =   5
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label lblTipo 
         Caption         =   "TIPO"
         Height          =   255
         Left            =   10200
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
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
            Key             =   "sep1"
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
            Key             =   "sep2"
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
            Key             =   "sep3"
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
            Key             =   "sep4"
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
            Object.ToolTipText     =   "Confirmar Reserva"
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
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ValoresCargo"
            Object.ToolTipText     =   "Valores por Cargo"
            ImageKey        =   "list"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Editar"
            Object.ToolTipText     =   "Ver Histórico de OT"
            ImageKey        =   "Editar"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stbServicios 
      Height          =   6135
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmRecepcion2.frx":038A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fmeCia"
      Tab(0).Control(1)=   "fmePat"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Inventario Recepción - Comentario"
      TabPicture(1)   =   "frmRecepcion2.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeInv"
      Tab(1).Control(1)=   "fmeCom"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Mecánica"
      TabPicture(2)   =   "frmRecepcion2.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeMec"
      Tab(2).Control(1)=   "stbTotalMec"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Carroceria"
      TabPicture(3)   =   "frmRecepcion2.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeCar"
      Tab(3).Control(1)=   "stbTotalDesabolladura"
      Tab(3).Control(2)=   "stbTotalArmeyDesarme"
      Tab(3).Control(3)=   "stbTotalCarroceria"
      Tab(3).Control(4)=   "stbTotalPintura"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Trabajos Adicionales"
      TabPicture(4)   =   "frmRecepcion2.frx":03FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeOtr"
      Tab(4).Control(1)=   "stbTotalOtros"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Servicios Externos"
      TabPicture(5)   =   "frmRecepcion2.frx":0416
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fmeTer"
      Tab(5).Control(1)=   "stbTotalTerceros"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Repuestos"
      TabPicture(6)   =   "frmRecepcion2.frx":0432
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "stbInsumos"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "stbTotalMateriales"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "stbTotalRepuestos"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "StbLubricantes"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "fmeRep"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).ControlCount=   5
      Begin VB.Frame fmeCia 
         Height          =   1545
         Left            =   -75000
         TabIndex        =   127
         Top             =   4545
         Width           =   12180
         Begin VB.TextBox txtNroSiniestro 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7320
            MaxLength       =   30
            TabIndex        =   136
            Top             =   330
            Width           =   2925
         End
         Begin VB.TextBox txtNroPoliza 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7320
            MaxLength       =   30
            TabIndex        =   135
            Top             =   750
            Width           =   1500
         End
         Begin VB.TextBox txtLiquidador 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7320
            MaxLength       =   50
            TabIndex        =   134
            Top             =   1125
            Width           =   4740
         End
         Begin VB.Frame Frame3 
            Caption         =   "Deducible"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   150
            TabIndex        =   129
            Top             =   765
            Width           =   5400
            Begin VB.TextBox txtDeducibleUF 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   720
               MaxLength       =   4
               TabIndex        =   131
               Top             =   240
               Width           =   1920
            End
            Begin VB.TextBox txtDeduciblePesos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3330
               MaxLength       =   8
               TabIndex        =   130
               Top             =   240
               Width           =   1920
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dólares"
               Height          =   195
               Index           =   20
               Left            =   105
               TabIndex        =   133
               Top             =   270
               Width           =   540
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Soles"
               Height          =   195
               Index           =   19
               Left            =   2730
               TabIndex        =   132
               Top             =   270
               Width           =   390
            End
         End
         Begin VB.TextBox txtOrdenReparacion 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10440
            TabIndex        =   128
            Top             =   720
            Width           =   1620
         End
         Begin MSComctlLib.Toolbar tlbCiaSeg 
            Height          =   330
            Left            =   5085
            TabIndex        =   137
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
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compañia de Seguro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   143
            Top             =   225
            Width           =   1815
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liquidador"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   6180
            TabIndex        =   142
            Top             =   1230
            Width           =   885
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Poliza"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   6165
            TabIndex        =   141
            Top             =   825
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Siniestro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   6180
            TabIndex        =   140
            Top             =   405
            Width           =   1020
         End
         Begin VB.Label lblCompañia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   150
            TabIndex        =   139
            Top             =   420
            Width           =   4890
         End
         Begin VB.Label Label4 
            Caption         =   "N° O. Reparación"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8880
            TabIndex        =   138
            Top             =   765
            Width           =   1575
         End
      End
      Begin VB.Frame fmeInv 
         Caption         =   "Inventario Recepciòn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   -74950
         TabIndex        =   123
         Top             =   350
         Width           =   4425
         Begin VB.ComboBox cmbBencina 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmRecepcion2.frx":044E
            Left            =   1200
            List            =   "frmRecepcion2.frx":0461
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   5160
            Width           =   2295
         End
         Begin MSComctlLib.ListView lvwInventario 
            Height          =   4815
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   8493
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
         Begin VB.Label Label6 
            Caption         =   "Gasolina"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   5160
            Width           =   855
         End
      End
      Begin VB.Frame fmeCom 
         Caption         =   "Comentario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   -70350
         TabIndex        =   121
         Top             =   350
         Width           =   6645
         Begin VB.TextBox txtComentario 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5300
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   122
            Top             =   240
            Width           =   6330
         End
      End
      Begin VB.Frame fmeMec 
         Height          =   5295
         Left            =   -74950
         TabIndex        =   112
         Top             =   350
         Width           =   11700
         Begin VB.CommandButton cmdReserva 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Reservar Repuestos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9690
            TabIndex        =   115
            Top             =   4785
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.CommandButton cmdAnularReserva 
            Appearance      =   0  'Flat
            Caption         =   "&Anular Reserva"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7920
            TabIndex        =   114
            Top             =   4800
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.CommandButton cmdConsultaSaldo 
            Appearance      =   0  'Flat
            Caption         =   "Consulta Saldos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   113
            Top             =   4800
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSComctlLib.Toolbar tlbAgregarRepuestos 
            Height          =   330
            Left            =   120
            TabIndex        =   116
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
         End
         Begin MSComctlLib.ListView lvwServiciosMecanica 
            Height          =   1740
            Left            =   45
            TabIndex        =   117
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   17
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
               Text            =   "P. Unitario"
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
               Text            =   "Monto Dscto."
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
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   13
               Text            =   "HorasReales"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "IdTarea"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "EstadoTarea"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "CentroCosto"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbAddServicioMec 
            Height          =   330
            Left            =   120
            TabIndex        =   118
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
         End
         Begin MSComctlLib.ListView lvwRepuestosMantencion 
            Height          =   1920
            Left            =   30
            TabIndex        =   119
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
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
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Saldo"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Repuestos Mantención"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            TabIndex        =   120
            Top             =   2565
            Width           =   1980
         End
      End
      Begin VB.Frame fmeRep 
         Height          =   4800
         Left            =   50
         TabIndex        =   105
         Top             =   350
         Width           =   11700
         Begin VB.CommandButton cmdConsultaStock 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Caption         =   "Consultar Stock"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   9840
            TabIndex        =   106
            Top             =   4400
            Width           =   1695
         End
         Begin MSComctlLib.ListView lvwRepuestos 
            Height          =   4065
            Left            =   0
            TabIndex        =   107
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   16
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
               Text            =   "P.  Unitario"
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
               Text            =   "Monto Dscto."
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
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Saldo"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "CentroCosto"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "PrecioVentaD"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   3
            Left            =   11805
            TabIndex        =   108
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
            TabIndex        =   109
            Top             =   4320
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
         End
      End
      Begin VB.Frame fmeTer 
         Height          =   5310
         Left            =   -74950
         TabIndex        =   101
         Top             =   350
         Width           =   11700
         Begin MSComctlLib.ListView lvwServiciosTerceros 
            Height          =   4400
            Left            =   50
            TabIndex        =   102
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   17
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
               Text            =   "Monto Recargo"
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
               Text            =   "Monto Dscto."
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
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "CentroCosto"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   2
            Left            =   0
            TabIndex        =   103
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
            TabIndex        =   104
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
         End
      End
      Begin VB.Frame fmeOtr 
         Height          =   5250
         Left            =   -74950
         TabIndex        =   98
         Top             =   350
         Width           =   11700
         Begin MSComctlLib.ListView lvwOtrosServicios 
            Height          =   4400
            Left            =   50
            TabIndex        =   99
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   16
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
               Text            =   "P. Unitario"
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
               Text            =   "Monto Dscto."
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
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Text            =   "Horas Reales"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "IdTarea"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "EstadoTarea"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "CentroCosto"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbAddServicioOtr 
            Height          =   330
            Left            =   105
            TabIndex        =   100
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
         End
      End
      Begin VB.Frame fmeCar 
         Height          =   4905
         Left            =   -74950
         TabIndex        =   74
         Top             =   350
         Width           =   11700
         Begin VB.TextBox txtMtoDesCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6450
            MaxLength       =   8
            TabIndex        =   80
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtPorcDesCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5955
            TabIndex        =   79
            Text            =   "00.0"
            Top             =   405
            Visible         =   0   'False
            Width           =   500
         End
         Begin VB.TextBox txtValorFinCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7455
            MaxLength       =   8
            TabIndex        =   78
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtSeccion 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1995
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   77
            Top             =   405
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox txtHorasCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   76
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtValorDefCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   75
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin MSDataListLib.DataCombo dtcCargoCar 
            Bindings        =   "frmRecepcion2.frx":049C
            Height          =   315
            Left            =   8460
            TabIndex        =   81
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
            Bindings        =   "frmRecepcion2.frx":04B6
            Height          =   315
            Left            =   9720
            TabIndex        =   82
            Top             =   360
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
            Left            =   120
            TabIndex        =   83
            Top             =   165
            Width           =   11475
            _ExtentX        =   20241
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   20
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
               Text            =   "Monto Recargo"
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
               Text            =   "Monto Desc."
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
            BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   19
               Text            =   "CentroCosto"
               Object.Width           =   0
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcPartePieza 
            Bindings        =   "frmRecepcion2.frx":04D0
            Height          =   315
            Left            =   2370
            TabIndex        =   84
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
            Bindings        =   "frmRecepcion2.frx":04EE
            Height          =   315
            Left            =   60
            TabIndex        =   85
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
            TabIndex        =   86
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
         End
         Begin MSComctlLib.Toolbar tlbTemparioCarroceria 
            Height          =   330
            Left            =   10200
            TabIndex        =   87
            Top             =   4440
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            ButtonWidth     =   2011
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Temparios"
                  Key             =   "Temparios"
                  Object.ToolTipText     =   "Temparios de Carroceria"
                  ImageIndex      =   24
               EndProperty
            EndProperty
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horas"
            Height          =   195
            Index           =   62
            Left            =   4485
            TabIndex        =   97
            Top             =   225
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Cargo"
            Height          =   195
            Index           =   60
            Left            =   8580
            TabIndex        =   96
            Top             =   195
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mecánico Asigado"
            Height          =   195
            Index           =   59
            Left            =   9930
            TabIndex        =   95
            Top             =   210
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% Desc."
            Height          =   195
            Index           =   49
            Left            =   5880
            TabIndex        =   94
            Top             =   195
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$ Desc."
            Height          =   195
            Index           =   48
            Left            =   6675
            TabIndex        =   93
            Top             =   195
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$ a Utilizar"
            Height          =   195
            Index           =   28
            Left            =   7530
            TabIndex        =   92
            Top             =   210
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$ Def."
            Height          =   195
            Index           =   27
            Left            =   5205
            TabIndex        =   91
            Top             =   210
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parte / Pieza"
            Height          =   195
            Index           =   26
            Left            =   2925
            TabIndex        =   90
            Top             =   225
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   25
            Left            =   1995
            TabIndex        =   89
            Top             =   210
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto"
            Height          =   195
            Index           =   24
            Left            =   720
            TabIndex        =   88
            Top             =   210
            Visible         =   0   'False
            Width           =   690
         End
      End
      Begin VB.Frame fmePat 
         Height          =   4275
         Left            =   -75000
         TabIndex        =   12
         Top             =   350
         Width           =   12180
         Begin VB.ComboBox cboHora 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4245
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   3870
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox txtKilAct 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   27
            Top             =   2535
            Width           =   1380
         End
         Begin VB.TextBox txtAño 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7875
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   26
            Top             =   1695
            Width           =   600
         End
         Begin VB.TextBox txtNroCono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4230
            MaxLength       =   3
            TabIndex        =   25
            Top             =   3435
            Width           =   930
         End
         Begin VB.TextBox txtSolicita 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8040
            MaxLength       =   50
            TabIndex        =   24
            Top             =   3870
            Width           =   3825
         End
         Begin VB.TextBox txtConcesionario 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4095
            TabIndex        =   23
            Top             =   2520
            Width           =   3210
         End
         Begin VB.TextBox txtDir 
            Height          =   315
            Left            =   435
            MaxLength       =   50
            TabIndex        =   22
            Top             =   4275
            Width           =   4185
         End
         Begin VB.TextBox txtComuna 
            Height          =   315
            Left            =   4710
            MaxLength       =   50
            TabIndex        =   21
            Top             =   4290
            Width           =   4185
         End
         Begin VB.TextBox txtRut 
            Height          =   315
            Left            =   8955
            MaxLength       =   50
            TabIndex        =   20
            Top             =   4305
            Width           =   2085
         End
         Begin VB.TextBox txtFonos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10920
            MaxLength       =   3
            TabIndex        =   18
            Top             =   2955
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtFolioGarantia 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            MaxLength       =   30
            TabIndex        =   17
            Top             =   285
            Width           =   1875
         End
         Begin VB.OptionButton optReparacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Reparación"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   7080
            TabIndex        =   16
            Top             =   980
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.OptionButton optMantencion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Mantención"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   8400
            TabIndex        =   15
            Top             =   980
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.TextBox txtPatente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   14
            Top             =   990
            Width           =   1200
         End
         Begin VB.TextBox txtNReferencia 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5880
            MaxLength       =   15
            TabIndex        =   13
            Text            =   "0"
            Top             =   1020
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker pckFecVta 
            Height          =   315
            Left            =   8880
            TabIndex        =   19
            Top             =   2505
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   93585409
            CurrentDate     =   36796
         End
         Begin MSComCtl2.DTPicker pckFechaEntrega 
            Height          =   315
            Left            =   1080
            TabIndex        =   29
            Top             =   3885
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   93585409
            CurrentDate     =   36733
         End
         Begin MSDataListLib.DataCombo dtcTipoCono 
            Bindings        =   "frmRecepcion2.frx":0509
            Height          =   315
            Left            =   1080
            TabIndex        =   30
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
            TabIndex        =   31
            Top             =   960
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nuevo"
                  Object.ToolTipText     =   "Nueva Placa"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Buscar"
                  Object.ToolTipText     =   "Buscar Placa"
                  ImageIndex      =   9
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Historial"
                  ImageIndex      =   5
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcRecepcionista 
            Bindings        =   "frmRecepcion2.frx":0523
            Height          =   315
            Left            =   8025
            TabIndex        =   32
            Top             =   3420
            Width           =   3840
            _ExtentX        =   6773
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
            Left            =   10920
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   28
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":0542
                  Key             =   "Crear"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":0654
                  Key             =   "Menos"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":0AAC
                  Key             =   "Mas"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":0F04
                  Key             =   "Persona"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":135C
                  Key             =   "Editar"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":146E
                  Key             =   "Grabar"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":1580
                  Key             =   "Cancelar"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":1692
                  Key             =   "Borrar"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":17A4
                  Key             =   "Buscar"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":18B6
                  Key             =   "Imprimir"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":19C8
                  Key             =   "Cerrar"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":1ADA
                  Key             =   "Ayuda"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":1BEC
                  Key             =   "Primero"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":1CFE
                  Key             =   "Anterior"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":1E10
                  Key             =   "Siguiente"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":1F22
                  Key             =   "Ultimo"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":2034
                  Key             =   "Renovar"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":2146
                  Key             =   "SortAsc"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":2258
                  Key             =   "SortDesc"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":236A
                  Key             =   "Seleccion"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":27BC
                  Key             =   "Seleccion1"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":2C0E
                  Key             =   "Copiar"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":2D20
                  Key             =   "Vaciar"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":3174
                  Key             =   "Confirmar"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":3490
                  Key             =   "LiquidarPres"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":38E8
                  Key             =   "AnularPres"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":3D3C
                  Key             =   "Salir"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion2.frx":408E
                  Key             =   "list"
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcGarantia 
            Bindings        =   "frmRecepcion2.frx":41A0
            Height          =   315
            Left            =   1080
            TabIndex        =   33
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
         Begin MSComctlLib.Toolbar tlbBusca 
            Height          =   330
            Index           =   4
            Left            =   6840
            TabIndex        =   34
            Top             =   2955
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nuevo"
                  Object.ToolTipText     =   "Nuevo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Buscar"
                  Object.ToolTipText     =   "Modificar Cliente"
                  ImageIndex      =   9
               EndProperty
            EndProperty
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   3
            X1              =   135
            X2              =   10980
            Y1              =   720
            Y2              =   720
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
         Begin VB.Label lblIdCliente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5820
            TabIndex        =   73
            Top             =   2970
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblIdModelo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5670
            TabIndex        =   72
            Top             =   1695
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblIdMarca 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1170
            TabIndex        =   71
            Top             =   1695
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Motor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   4020
            TabIndex        =   70
            Top             =   2100
            Width           =   840
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Entrega"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   14
            Left            =   3075
            TabIndex        =   69
            Top             =   3930
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ent."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   68
            Top             =   3930
            Width           =   885
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cono"
            BeginProperty Font 
               Name            =   "Verdana"
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
            Left            =   120
            TabIndex        =   67
            Top             =   3420
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concesionario"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   2835
            TabIndex        =   66
            Top             =   2535
            Width           =   1215
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   65
            Top             =   2955
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color Exterior"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   8535
            TabIndex        =   64
            Top             =   1725
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   7530
            TabIndex        =   63
            Top             =   1695
            Width           =   330
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   3165
            TabIndex        =   62
            Top             =   1695
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   61
            Top             =   1695
            Width           =   510
         End
         Begin VB.Label lblPat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Placa"
            BeginProperty Font 
               Name            =   "Verdana"
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
            TabIndex        =   60
            Top             =   1020
            Width           =   525
         End
         Begin VB.Label lblMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1080
            TabIndex        =   59
            Top             =   1695
            Width           =   1980
         End
         Begin VB.Label lblModelo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3795
            TabIndex        =   58
            Top             =   1695
            Width           =   3540
         End
         Begin VB.Label lblColorE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9060
            TabIndex        =   57
            Top             =   1695
            Width           =   2880
         End
         Begin VB.Label lblCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1080
            TabIndex        =   56
            Top             =   2955
            Width           =   5880
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Cono"
            BeginProperty Font 
               Name            =   "Verdana"
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
            TabIndex        =   55
            Top             =   3435
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recepcionista"
            BeginProperty Font 
               Name            =   "Verdana"
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
            Left            =   5640
            TabIndex        =   54
            Top             =   3465
            Width           =   1350
         End
         Begin VB.Label lblChasis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1080
            TabIndex        =   53
            Top             =   2085
            Width           =   2850
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chasis"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   52
            Top             =   2130
            Width           =   570
         End
         Begin VB.Label lblVin 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8640
            TabIndex        =   51
            Top             =   2085
            Width           =   3300
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VIN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   8160
            TabIndex        =   50
            Top             =   2115
            Width           =   315
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quién trajo el vehiculo?"
            BeginProperty Font 
               Name            =   "Verdana"
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
            Left            =   5640
            TabIndex        =   49
            Top             =   3855
            Width           =   2325
         End
         Begin VB.Label lblFono 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8520
            TabIndex        =   48
            Top             =   2955
            Width           =   2250
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fonos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   7920
            TabIndex        =   47
            Top             =   2985
            Width           =   495
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
         Begin VB.Label lblMotor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4935
            TabIndex        =   46
            Top             =   2100
            Width           =   2790
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Venta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   7440
            TabIndex        =   45
            Top             =   2595
            Width           =   1050
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   120
            X2              =   10965
            Y1              =   600
            Y2              =   600
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
            Index           =   1
            X1              =   195
            X2              =   11040
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folio Gtía."
            BeginProperty Font 
               Name            =   "Verdana"
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
            Left            =   4365
            TabIndex        =   44
            Top             =   285
            Width           =   990
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo OT"
            BeginProperty Font 
               Name            =   "Verdana"
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
            TabIndex        =   43
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kms. Act."
            BeginProperty Font 
               Name            =   "Verdana"
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
            Left            =   120
            TabIndex        =   42
            Top             =   2610
            Width           =   900
         End
         Begin VB.Label lblPresupuesto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9600
            TabIndex        =   41
            Top             =   285
            Width           =   2175
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Presupuesto N°"
            BeginProperty Font 
               Name            =   "Verdana"
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
            Left            =   7995
            TabIndex        =   40
            Top             =   285
            Width           =   1500
         End
         Begin VB.Label lblDocumentos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3360
            TabIndex        =   39
            Top             =   1020
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Documentos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   38
            Top             =   765
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000A&
            Caption         =   "Fecha Liq."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   9720
            TabIndex        =   37
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblFechaLiquidacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   10800
            TabIndex        =   36
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "N° Referencia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   35
            Top             =   765
            Width           =   1215
         End
      End
      Begin MSComctlLib.StatusBar StbLubricantes 
         Height          =   405
         Left            =   1665
         TabIndex        =   11
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
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar stbTotalDesabolladura 
         Height          =   405
         Left            =   -72375
         TabIndex        =   110
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
            Name            =   "Verdana"
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
         TabIndex        =   111
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
      Begin MSComctlLib.StatusBar stbTotalRepuestos 
         Height          =   405
         Left            =   6690
         TabIndex        =   144
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
            Name            =   "Verdana"
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
         TabIndex        =   145
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
            Name            =   "Verdana"
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
         TabIndex        =   146
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
            Name            =   "Verdana"
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
         Left            =   6690
         TabIndex        =   147
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
            Name            =   "Verdana"
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
         Left            =   1680
         TabIndex        =   148
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
            Name            =   "Verdana"
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
         TabIndex        =   149
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
            Name            =   "Verdana"
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
         TabIndex        =   150
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
            Name            =   "Verdana"
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
         TabIndex        =   151
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
            Name            =   "Verdana"
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
      Height          =   405
      Left            =   6705
      TabIndex        =   152
      Top             =   7080
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
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdImpresora 
      Left            =   600
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbSeguroTaller 
      Height          =   405
      Left            =   1680
      TabIndex        =   153
      Top             =   7080
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
            Text            =   "Seguro Taller"
            TextSave        =   "Seguro Taller"
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
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1830
      Left            =   240
      Picture         =   "frmRecepcion2.frx":41BA
      Top             =   7440
      Visible         =   0   'False
      Width           =   8745
   End
End
Attribute VB_Name = "frmRecepcion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Dim adoPrincipal As New ADODB.Recordset
''Dim mstrSql As String
''Dim mstrWhere As String
''Dim mstrOrderBy As String
''Dim mblnTablaVacia As Boolean
''Dim mblnAccesoCrear As Boolean
''Dim mblnAccesoEditar As Boolean
''Dim mblnAccesoBorrar As Boolean
''Dim mblnAccesoImprimir As Boolean
''Dim mblnSW As Boolean
''Dim itmAux As ListItem
''Dim lsiItem As ListItem
''Dim intIndice As Integer
''Dim curValor As Currency
''Dim mstrTipoCargo As String
''Dim mstrIdOT As String
''Dim mstrCargo As String
''Dim mdblTotalInicial As Double
''Dim mstrIdPresupuestoOrigen As String
''Dim mstrProcedencia As String
''Dim mblnBloqueo As Boolean
''Dim dblTotalInicial As Double
''Dim KilometrajeEntrada As Double 'Variable de ILeiva 07/02/2001 para conservar el kilometraje de entrada asi lo comparo con el que va a ingresar en la recepción debe ser mayor
''Dim gstrEstadoMantencion As String
''Dim gstrEstadoReparacion As String
''Dim gstrEstadoDisponible As String
''Dim gstrBuscaReserva As String
''Dim NroRegularizacion As String
''Dim gstrKmsAutoNuevo As String
''Dim mstrEstadoPresupuesto As String
''Dim mstrAgregaPresupuesto As Boolean
''Dim mstrLiquidaPresupuesto As Boolean
''Dim mblnOtFacturada As Boolean
''Dim curSumaInsumos As Currency
''Dim mstrProcedenciaAux As String
''Sub TipoOt(pstrTipoOt As String)
'''lblPat.Caption
''Select Case pstrTipoOt
''Case "GFB"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        If Me.fmePat.Enabled = False Then
''            fmePat.Enabled = True
''        End If
''        .txtFolioGarantia.Enabled = True
''        .txtFolioGarantia.SetFocus
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''    End With
''Case "CS"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "INA"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "INR"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "INS"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "INU"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "INW"
''    With Me
''        .lblPat.Caption = "V.I.N."
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "NGN"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "INC"
''    With Me
''        .lblPat.Caption = "V.I.N."
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "PEX"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "REN"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = True
''        .optReparacion.Visible = True
''
''        .tlbAgregarRepuestos.Visible = True
''    End With
''Case "PRE"
''    With Me
''        .lblPat.Caption = gstrNombrePatente
''        .txtFolioGarantia = "S/F"
''        .txtFolioGarantia.Enabled = False
''        .optMantencion.Visible = False
''        .optReparacion.Visible = False
''        .cmdAnularReserva.Visible = False
''        .cmdReserva.Visible = False
''        .tlbAgregarRepuestos.Visible = False
''        mstrEstadoPresupuesto = "ON"
''        mstrLiquidaPresupuesto = False
''        gcurInsumo = 0
''    End With
''End Select
''End Sub
''
''
''Function ExistePatente(pstrPatente As String) As Boolean
''
''mstrSql = "Select top 1 * From Tllr_Vehiculo_Cliente"
''mstrSql = mstrSql & " WHERE Tllr_Vehiculo_Cliente.Patente = '" & pstrPatente & "'"
''If Conexion.SendHost(mstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With gadoPrincipal
''        If Not .BOF And Not .EOF Then
''            ExistePatente = True
''        Else
''            ExistePatente = False
''        End If
''    End With
''End If
''
''End Function
''
''Sub Bloqueo(pstrEstado As String)
''If pstrEstado = "V" Or pstrEstado = "R" Or pstrEstado = "P" Then
''    fmePat.Enabled = True
''    fmeCia.Enabled = True
''    fmeInv.Enabled = True
''    fmeCom.Enabled = True
''    mblnBloqueo = False
''ElseIf pstrEstado = "B" Or pstrEstado = "F" Then
''    If mblnOtFacturada = True Then
''        fmePat.Enabled = True
''        fmeCia.Enabled = True
''        fmeInv.Enabled = True
''        fmeCom.Enabled = True
''        mblnBloqueo = False
''    Else
''        fmePat.Enabled = False
''        fmeCia.Enabled = False
''        fmeInv.Enabled = False
''        fmeCom.Enabled = False
''        mblnBloqueo = True
''    End If
''Else
''    fmePat.Enabled = False
''    fmeCia.Enabled = False
''    fmeInv.Enabled = False
''    fmeCom.Enabled = False
''    mblnBloqueo = True
''End If
''End Sub
''
''Function TotalSeccionCargo(pstrIdEmpresa As String, _
''                            pstrIdSucursal As String, _
''                            pstrIdOT As String, _
''                            pstrIdTipoCargo As String, _
''                            pstrTipoOt As String, _
''                            Seccion As SumSec) As Currency
''If pstrIdTipoCargo = "" Then
''    If Seccion = ssMec Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_MECANICA_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssOtr Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_OTRO_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssCar Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN  FROM TLLR_CARROCERIA_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssTer Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_TERCEROS_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssRep Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_REPUESTOS_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    End If
''Else
''    If Seccion = ssMec Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_MECANICA_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssOtr Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_OTRO_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssCar Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN  FROM TLLR_CARROCERIA_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssTer Then
''        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_TERCEROS_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    ElseIf Seccion = ssRep Then
''        gstrSql = "SELECT SUM(SUBTOTAL)  AS RESUMEN FROM TLLR_REPUESTOS_OT"
''        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    End If
''End If
''If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With gadoPrincipal
''        If Not .BOF And Not .EOF Then
''            .MoveFirst
''            If Not IsNull(!Resumen) Then
''                TotalSeccionCargo = !Resumen
''            Else
''                TotalSeccionCargo = 0
''            End If
''        End If
''        .Close
''    End With
''End If
''
''End Function
''Function VerificaLubricantesTipoCargo(pstrIdEmpresa As String, _
''                            pstrIdSucursal As String, _
''                            pstrIdOT As String, _
''                            pstrIdTipoCargo As String, _
''                            pstrTipoOt As String, _
''                            Seccion As SumSec) As Currency
''
''Dim SumaLubricantes As Currency
''Dim SumaMateriales As Currency
''Dim SumaInsumos As Currency
''
''
''If pstrIdTipoCargo = "" Then
''    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
''    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
''    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
''    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "'"
''    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoLubricantes & "'" '90'
''Else
''    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
''    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
''    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
''    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoLubricantes & "'" '90'
''End If
''SumaLubricantes = 0
''
''If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With gadoPrincipal
''        While Not .EOF
''            SumaLubricantes = SumaLubricantes + !SubTotal
''            .MoveNext
''        Wend
''        .Close
''    End With
''End If
''VerificaLubricantesTipoCargo = SumaLubricantes
''
'''///// MATERIALES
''If pstrIdTipoCargo = "" Then
''    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
''    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
''    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
''    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "'"
''    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoMateriales & "'" '85'
''Else
''    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
''    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
''    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
''    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoMateriales & "'" '85'
''
''End If
''SumaMateriales = 0
''
''If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With gadoPrincipal
''        While Not .EOF
''            SumaMateriales = SumaMateriales + !SubTotal
''            .MoveNext
''        Wend
''        .Close
''    End With
''End If
''gcurMateriales = SumaMateriales
''
''
'''///// Insumos
''If pstrIdTipoCargo = "" Then
''    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
''    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
''    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
''    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "'"
''    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoInsumos & "'" '85'
''Else
''    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
''    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
''    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
''    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
''    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
''    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
''    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoInsumos & "'" '85'
''
''End If
''SumaInsumos = 0
''
''If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With gadoPrincipal
''        While Not .EOF
''            SumaInsumos = SumaInsumos + !SubTotal
''            .MoveNext
''        Wend
''        .Close
''    End With
''End If
''curSumaInsumos = SumaInsumos
''
''
''End Function
''
''Function AccesoEliminar(itmSeleccionado As ListItem) As Boolean
'''If itmSeleccionado.SubItems(5) = "85" Then
''    AccesoEliminar = True
'''Else
'''    AccesoEliminar = False
'''End If
''End Function
''
''Sub PrintOT()
''Dim mstrIdCargo As String
''Dim mcurTNeto As Currency
''Dim mcurTMec As Currency
''Dim mcurTOtr As Currency
''Dim mcurTCar As Currency
''Dim mcurTTer As Currency
''Dim mcurTRep As Currency
''Dim mcurTMat As Currency
''Dim mcurTIns As Currency
''Dim mcurTLub As Currency
''Dim mcurDeducible As Currency
''Dim lstrArchivoIni As String
''lstrArchivoIni = Command()
''gstrPathReporte = LetConnectionString("TLLR", "RPT", lstrArchivoIni, 256)
''
'''/// MODIFICADO POR FDO DIAZ EL 11/12/2000
'''/// PREGUNTA PRIMERO SI ES UNA RECEPCION Y DESPUES PREGUNTA DE QUE TIPO DE IMPRESION ES.
'''/// SI ES PREIMPRESO COMO AUTOSUMMIT O UNA IMPRESION EN BLANCO
''
''On Error GoTo Solucion
''
''If gstrImpresion = "R" Then
''
''    If TipoImpresion = "C" Then  'LA LETRA "C" ES PREIMPRESO AUTOSUMMIT
''        ImprimirDocumento gRecepcion
''    ElseIf TipoImpresion = "P" Then  'LA LETRA "P" ES PREIMPRESO PIAMONTE
''        ImprimirDocumentoPiamonte gRecepcion
''    ElseIf TipoImpresion = "K" Then  'LA LETRA "K" ES PREIMPRESO klassik car
''        ImprimirDocumentoKlassik gRecepcion
''    Else
''       ImprimirDocumentoRecepcion gRecepcion  ' // FORMATO RECEPCION STANDARD(HOJA EN BLANCO)
''    End If
''
''ElseIf gstrImpresion = "O" Then
''
''  If Me.dtcGarantia.BoundText <> "PRE" Then
''
''    If Val(txtDeduciblePesos) = 0 And Val(txtDeducibleUF) = 0 Then
''
''        gstrSql = "SELECT ID_TIPO_CARGO FROM TLLR_TIPO_CARGO"
''        If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''            With gadoPrincipal
''                If Not .BOF And Not .EOF Then
''                    .MoveFirst
''                    While Not .EOF
''                        mstrIdCargo = !Id_Tipo_Cargo
''                        mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssMec)
''                        mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssOtr)
''                        mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssCar)
''                        mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssTer)
''
''                        mcurTLub = VerificaLubricantesTipoCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep)
''                        mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep) '- IIf(mstrIdCargo = "01", gcurMateriales, 0)
''                        'mcurTIns = CalculoInsumos(8) + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0)
''                        mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0) + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0)
''                        'mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + (mcurTRep - gcurMateriales) + IIf(mstrIdCargo = "01", gcurInsumo, 0) + IIf(mstrIdCargo = "01", gcurMateriales, 0) + IIf(mstrIdCargo = "01", gcurSeguroTaller, 0)
''
''                        If mcurTNeto > 0 Then
''                          ' Antes
''                            With rptOT
''
''                            Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
''                            Me.cdImpresora.CancelError = True
''                            Me.cdImpresora.Action = 5
''
''                            .CopiesToPrinter = cdImpresora.Copies
''                            If gstrServiciosMarca = "S" Then
''                                .ReportFileName = gstrPathReporte & "\OTMM.rpt"
''                            Else
''                                .ReportFileName = gstrPathReporte & "\OT.rpt"
''                            End If
''                            .Destination = crptToPrinter
''                            .WindowState = crptMaximized
''
''                                .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
''                                .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
''                                .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
''                                .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
''                                .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
''                                .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
''                                .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
''
''                                .Formulas(7) = "TMecanica=" & mcurTMec & ""
''                                .Formulas(8) = "TOtros=" & mcurTOtr & ""
''                                .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
''                                .Formulas(10) = "TRepuesto=" & mcurTRep - (mcurTLub + gcurMateriales + curSumaInsumos) & ""
''                                .Formulas(11) = "TDyP=" & mcurTCar & ""
''                                .Formulas(12) = "TTerceros=" & mcurTTer & ""
''
''                                .Formulas(13) = "TMateriales=" & gcurMateriales & "" '& IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
''                                .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo + curSumaInsumos, 0) & ""
''                                .Formulas(15) = ""
''                                .Formulas(16) = ""
''                                .Formulas(17) = "TNetoOT=" & mcurTNeto & ""
''                                .Formulas(18) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
''                                .Formulas(19) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
''                                .Formulas(20) = "TLubricantes=" & mcurTLub & ""
''                                .Formulas(21) = "SeguroTaller=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0) & ""
''                                .Formulas(22) = "NotaRecepcion='" & IIf(gstrNotaRecepcion = "", "", "OK") & "'"
''                                .Formulas(23) = "TipoCargo='" & mstrIdCargo & "'"
''                                .Formulas(24) = "NombreIva='" & gstrNombreIva & "'"
''                                .Formulas(25) = "Tdecimal=" & gintDecimalesMoneda & ""
''                                .Formulas(26) = "NombreRut='" & gstrNombreRut & "'"
''                                .Formulas(27) = "NombrePatente='" & gstrNombrePatente & "'"
''                                .Formulas(28) = "FamiliaInsumos='" & gstrCodigoInsumos & "'"
''                                .Formulas(29) = "FamiliaLubricantes='" & gstrCodigoLubricantes & "'"
''                                .Formulas(30) = "FamiliaMateriales='" & gstrCodigoMateriales & "'"
''                                .Formulas(31) = "EditaRut='" & gstrEditaRut & "'"
''                                .Formulas(32) = "TipodeOt='" & Me.dtcGarantia.Text & "'"
''                                .Connect = "Driver={SQL Server};Server=wiracocha;UID=sa;PWD=Llosa1936;Database=elisa;" 'Conexion.ConnectionString
'''                                .Connect = "Driver={SQL Server};Server=wiracocha;UID=sa;PWD=Llosa1936;Database=Prueba;" 'Conexion.ConnectionString
''                                .SelectionFormula = "{Tllr_OT.Id_Empresa}='" & gstrIdEmpresa & "' And {Tllr_OT.Id_Sucursal}='" & gstrIdSucursal & "' And {Tllr_OT.Id_OT}='" & lblNroRecepcion & "' And {Tllr_OT.Seccion_OT}='" & gstrSeccion & "'"
''
''
''                                .Action = True
''                            End With
''                            .MoveNext
''                        Else
''                            .MoveNext
''                        End If
''                        mcurTMec = 0
''                        mcurTOtr = 0
''                        mcurTCar = 0
''                        mcurTTer = 0
''                        mcurTRep = 0
''                        mcurTLub = 0
''                        mcurTNeto = 0
''                    Wend
''                End If
''            End With
''        Else
''            DoEvents
''            Exit Sub
''        End If
''    Else    '/////////////////////////////////////////////////deducible <>0
''        Dim mblndeducible As Boolean
''        mcurDeducible = CCur(Val(txtDeduciblePesos))
''        gstrSql = "SELECT ID_TIPO_CARGO FROM TLLR_TIPO_CARGO"
''        If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''            With gadoPrincipal
''                If Not .BOF And Not .EOF Then
''                    .MoveLast
''                    While Not .BOF
''                        mstrIdCargo = !Id_Tipo_Cargo
''                        mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssMec)
''                        mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssOtr)
''                        mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssCar)
''                        mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssTer)
''                        'MODIFICADO POR FDO DIAZ EL 04/01/2001
''                        mcurTLub = VerificaLubricantesTipoCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep)
''                        mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep) '- IIf(mstrIdCargo = "01", gcurMateriales, 0)
''                        'mcurTIns = CalculoInsumos(8)
''                        mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0) + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0)
''
''                        'si solo existe deducible
''                        If mcurTNeto = 0 Then
''                            If mstrIdCargo = gstrCargoDeducibleMas Then
''                                mblndeducible = True
''                            End If
''                        End If
''                        If mcurTNeto > 0 Or mblndeducible = True Then
''                            With rptOT
''
''                                Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
''                                Me.cdImpresora.CancelError = True
''                                Me.cdImpresora.Action = 5
''
''                                .CopiesToPrinter = cdImpresora.Copies
''                                If gstrServiciosMarca = "S" Then
''                                    .ReportFileName = gstrPathReporte & "\OTCDMM.rpt"
''                                Else
''                                    .ReportFileName = gstrPathReporte & "\OTCD.rpt"
''                                End If
''                                .Destination = crptToPrinter
''                                .WindowState = crptMaximized
''                                If gstrIdEmpresa = "832207004" Or InStr(gstrEmpresa, "SERINFO") = 1 Then
''                                    .Destination = crptToWindow
''                                End If
''                                .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
''                                .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
''                                .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
''                                .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
''                                .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
''                                .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
''                                .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
''
''                                .Formulas(7) = "TMecanica=" & mcurTMec & ""
''                                .Formulas(8) = "TOtros=" & mcurTOtr & ""
''                                .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
''                                .Formulas(10) = "TRepuesto=" & mcurTRep - (mcurTLub + gcurMateriales + curSumaInsumos) & ""
''                                '.Formulas(10) = "TRepuesto=" & mcurTRep - mcurTLub - mcurTIns & ""
''                                .Formulas(11) = "TDyP=" & mcurTCar & ""
''                                .Formulas(12) = "TTerceros=" & mcurTTer & ""
''
''                                .Formulas(13) = "TMateriales=" & gcurMateriales & ""  '& IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
''                                .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo + curSumaInsumos, 0) & ""
''                                .Formulas(20) = "TLubricantes=" & mcurTLub & ""
''
''                                If mstrIdCargo = gstrCargoDeducibleMenos Then
''                                    If mcurDeducible <= mcurTNeto Then
''                                        .Formulas(15) = "Anexo= 'Deducible ( - )'"
''                                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
''                                        mcurTNeto = mcurTNeto - mcurDeducible
''                                    End If
''                                ElseIf mstrIdCargo = gstrCargoDeducibleMas Then
''                                        .Formulas(15) = "Anexo= 'Deducible ( + )'"
''                                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
''                                        mcurTNeto = mcurTNeto + mcurDeducible
''                                Else
''                                    .Formulas(15) = ""
''                                    .Formulas(16) = ""
''                                End If
''                                .Formulas(17) = "TNetoOT=" & mcurTNeto & ""
''                                .Formulas(18) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
''                                .Formulas(19) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
''                                .Formulas(21) = "SeguroTaller=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0) & ""
''                                .Formulas(22) = "NotaRecepcion='" & IIf(gstrNotaRecepcion = "", "", "OK") & "'"
''                                .Formulas(23) = "TipoCargo='" & mstrIdCargo & "'"
''                                .Formulas(24) = "NombreIva='" & gstrNombreIva & "'"
''                                .Formulas(25) = "Tdecimal=" & gintDecimalesMoneda & ""
''                                .Formulas(26) = "NombreRut='" & gstrNombreRut & "'"
''                                .Formulas(27) = "NombrePatente='" & gstrNombrePatente & "'"
''                                .Formulas(28) = "FamiliaInsumos='" & gstrCodigoInsumos & "'"
''                                .Formulas(29) = "FamiliaLubricantes='" & gstrCodigoLubricantes & "'"
''                                .Formulas(30) = "FamiliaMateriales='" & gstrCodigoMateriales & "'"
''                                .Formulas(31) = "EditaRut='" & gstrEditaRut & "'"
''                                .Formulas(32) = "TipodeOt='" & Me.dtcGarantia.Text & "'"
'''                                .Connect = Conexion.ConnectionString
''                                .Connect = "Driver={SQL Server};Server=wiracocha;UID=sa;PWD=Llosa1936;Database=elisa;" 'Conexion.ConnectionString
''                                .SelectionFormula = "{Tllr_OT.Id_Empresa}='" & gstrIdEmpresa & "' And {Tllr_OT.Id_Sucursal}='" & gstrIdSucursal & "' And {Tllr_OT.Id_OT}='" & lblNroRecepcion & "' And {Tllr_OT.Seccion_OT}='" & gstrSeccion & "'"
''
''                                .Action = True
''
''                            End With
''                            .MovePrevious
''                        Else
''                            .MovePrevious
''                        End If
''                        mcurTMec = 0
''                        mcurTOtr = 0
''                        mcurTCar = 0
''                        mcurTTer = 0
''                        mcurTRep = 0
''                        mcurTLub = 0
''                        mcurTNeto = 0
''                    Wend
''                End If
''            End With
''        Else
''            DoEvents
''            Exit Sub
''        End If
''    End If
''
''  Else  '//// es presupuesto
''        mcurDeducible = CCur(Val(txtDeduciblePesos))
''        mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssMec)
''        mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssOtr)
''        mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssCar)
''        mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssTer)
''        'MODIFICADO POR FDO DIAZ EL 04/01/2001
''        mcurTLub = VerificaLubricantesTipoCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssRep)
''        mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssRep) '- IIf(mstrIdCargo = "01", gcurMateriales, 0)
''        mcurTIns = CalculoInsumos(8)
''        mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep
''
''  'kjcv 05.09.16 base formada en access
''        Dim Dbsnueva As Database
''        Dim Tabla As DAO.Recordset
''        Dim i As Integer
''        Dim GcamBaseTem As String
''
''        gstrNombreRecepcionista = NombreRecepcionista(dtcRecepcionista.BoundText)
'''        gstrNombreRecepLlamado = NombreRecepcionista(dtcRecepcionista.BoundText)
''
''                Dim rc As Long
''                Dim WinPath As String
''                WinPath = Space$(300)
''                rc = GetWindowsDirectory(WinPath, 300)
''                GcamBaseTem = Trim$(WinPath)
''                GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
''                '---------------------------------------
''
''                Dim wrkPredeterminado As Workspace
''                Dim prpBucle As Property
''                Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
''                If Dir(gstrPathReporte & "\BDNuevaPresu.mdb") <> "" Then Kill gstrPathReporte & "\BDNuevaPresu.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
''                Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNuevaPresu.mdb", dbLangGeneral) ' Crea a una base de datos nueva
''                Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (OT text, Seccion text,Recepcionista text,FLiquida text, Cliente text,Direccion text,DNI text,Telefono text,Marca text, Modelo text,Patente text,VIN text,Color text,año text,Motor text,Siniestro text,Poliza text, Liquidador text,Compañia text, DeduSoles text,DeduDolar text, Observaciones memo)"
''                Dbsnueva.Execute "CREATE TABLE T_TOTALES(OT text,TManoObra text, TRepuestos text, TPyP text,TTerceros text, Insumos text,Lubricantes text, TOtros text, SubTotal text,IVA text,Total text)"
''                Dbsnueva.Execute "CREATE TABLE T_PARAMECANICA (OT text,IdServicio text,Descripcion text,Cargo text,Horas text,Porcentaje_Dscto text, Monto_Dscto text, MSubtotal text)"
''                Dbsnueva.Execute "CREATE TABLE T_PARASERVICIO (OT text,IdOtroServicio text,Servicio text,Cargo text, Horas text,Porcentaje_Dscto text,Monto_Dscto text, OSubTotal text)"
''                Dbsnueva.Execute "CREATE TABLE T_PARACARROCERIA (OT text,Carroceria text,CSubtotal text)"
''                Dbsnueva.Execute "CREATE TABLE T_PARATERCEROS(OT text,IdServicioTercero text, Tercero text, Cargo text,Porcentaje_Dscto text,Monto_Dscto text,TSubtotal text)"
''                Dbsnueva.Execute "CREATE TABLE T_PARAREPUESTOS(OT text,IdItem text, Saldo text, Pieza text, Cargo text, Valor text, Cantidad text,RSubtotal double)"
''
''                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
''                Tabla.AddNew
''                Tabla!OT = Me.lblNroRecepcion
''                Tabla!Seccion = gstrSeccion
''                Tabla!Recepcionista = gstrNombreRecepcionista
''                Tabla!FLiquida = Me.lblFechaLiquidacion
''                Tabla!Cliente = Me.lblCliente.Caption
''                Tabla!Direccion = TraeDireccion(Me.lblIdCliente)
''                Tabla!DNI = Me.lblIdCliente
''                Tabla!Telefono = Me.lblFono
''                Tabla!Marca = Me.lblMarca
''                Tabla!Modelo = Me.lblModelo
''                Tabla!Patente = Me.txtPatente
''                Tabla!VIN = Me.lblVin
''                Tabla!Color = Me.lblColorE
''                Tabla!Año = Me.txtAño
''                Tabla!motor = Me.lblMotor
''                Tabla!Siniestro = Me.txtNroSiniestro
''                Tabla!Poliza = Me.txtNroPoliza
''                Tabla!Liquidador = Me.txtLiquidador
''                Tabla!Compañia = Me.lblCompañia
''                Tabla!DeduSoles = Me.txtDeduciblePesos
''                Tabla!DeduDolar = Me.txtDeducibleUF
''                Tabla!Observaciones = Me.txtComentario
''                Tabla.Update
''                Tabla.Close
''
''                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAMECANICA")
''                For i = 1 To lvwServiciosMecanica.ListItems.Count
''                    Set lvwServiciosMecanica.SelectedItem = lvwServiciosMecanica.ListItems(i)
''                    Tabla.AddNew
''                    Tabla!idServicio = lvwServiciosMecanica.ListItems(i)
''                    Tabla!Descripcion = IIf(lvwServiciosMecanica.SelectedItem.SubItems(1) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(1))
''                    Tabla!CARGO = IIf(lvwServiciosMecanica.SelectedItem.SubItems(7) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(7))
''                    Tabla!Horas = IIf(lvwServiciosMecanica.SelectedItem.SubItems(2) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(2))
''                    Tabla!Porcentaje_Dscto = IIf(lvwServiciosMecanica.SelectedItem.SubItems(4) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(4))
''                    Tabla!monto_Dscto = IIf(lvwServiciosMecanica.SelectedItem.SubItems(5) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(5))
''                    Tabla!MSubtotal = IIf(lvwServiciosMecanica.SelectedItem.SubItems(10) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(10))
''                    Tabla.Update
''                Next i
''                Tabla.Close
''
''                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARASERVICIO")
''                For i = 1 To lvwOtrosServicios.ListItems.Count
''                    Set lvwOtrosServicios.SelectedItem = lvwOtrosServicios.ListItems(i)
''                    Tabla.AddNew
''                    Tabla!IdOtroServicio = lvwOtrosServicios.ListItems(i)
''                    Tabla!servicio = IIf(lvwOtrosServicios.SelectedItem.SubItems(1) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(1))
''                    Tabla!CARGO = IIf(lvwOtrosServicios.SelectedItem.SubItems(7) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(7))
''                    Tabla!Horas = IIf(lvwOtrosServicios.SelectedItem.SubItems(2) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(2))
''                    Tabla!Porcentaje_Dscto = IIf(lvwOtrosServicios.SelectedItem.SubItems(4) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(4))
''                    Tabla!monto_Dscto = IIf(lvwOtrosServicios.SelectedItem.SubItems(5) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(5))
''                    Tabla!OSubtotal = IIf(lvwOtrosServicios.SelectedItem.SubItems(10) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(10))
''                    Tabla.Update
''                Next i
''                Tabla.Close
''
''                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARACARROCERIA")
''                For i = 1 To lvwServiciosCarroceria.ListItems.Count
''                    Set lvwServiciosCarroceria.SelectedItem = lvwServiciosCarroceria.ListItems(i)
''                    Tabla.AddNew
''                    Tabla!Carroceria = lvwServiciosCarroceria.ListItems(2)
''                    Tabla!CSubtotal = IIf(lvwServiciosCarroceria.SelectedItem.SubItems(16) = "", " ", lvwServiciosCarroceria.SelectedItem.SubItems(16))
''                    Tabla.Update
''                Next i
''                Tabla.Close
''
''                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARATERCEROS")
''                For i = 1 To lvwServiciosTerceros.ListItems.Count
''                    Set lvwServiciosTerceros.SelectedItem = lvwServiciosTerceros.ListItems(i)
''                    Tabla.AddNew
''                    Tabla!IdServicioTercero = lvwServiciosTerceros.ListItems(i)
''                    Tabla!Tercero = IIf(lvwServiciosTerceros.SelectedItem.SubItems(3) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(3))
''                    Tabla!CARGO = IIf(lvwServiciosTerceros.SelectedItem.SubItems(13) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(13))
''                    Tabla!Porcentaje_Dscto = IIf(lvwServiciosTerceros.SelectedItem.SubItems(10) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(10))
''                    Tabla!monto_Dscto = IIf(lvwServiciosTerceros.SelectedItem.SubItems(11) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(11))
''                    Tabla!TSubtotal = IIf(lvwServiciosTerceros.SelectedItem.SubItems(12) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(12))
''                    Tabla.Update
''                Next i
''                Tabla.Close
''
''
''                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPUESTOS")
''                For i = 1 To lvwRepuestos.ListItems.Count
''                    Set lvwRepuestos.SelectedItem = lvwRepuestos.ListItems(i)
''                    Tabla.AddNew
''                    Tabla!IdItem = lvwRepuestos.ListItems(i)
''                    Tabla!Saldo = IIf(lvwRepuestos.SelectedItem.SubItems(12) = "", " ", lvwRepuestos.SelectedItem.SubItems(12))
''                    Tabla!pieza = IIf(lvwRepuestos.SelectedItem.SubItems(1) = "", " ", lvwRepuestos.SelectedItem.SubItems(1))
''                    Tabla!CARGO = IIf(lvwRepuestos.SelectedItem.SubItems(6) = "", " ", lvwRepuestos.SelectedItem.SubItems(6))
''                    Tabla!Valor = IIf(lvwRepuestos.SelectedItem.SubItems(3) = "", " ", lvwRepuestos.SelectedItem.SubItems(3))
''                    Tabla!cantidad = IIf(lvwRepuestos.SelectedItem.SubItems(2) = "", " ", lvwRepuestos.SelectedItem.SubItems(2))
''                    Tabla!RSubtotal = IIf(lvwRepuestos.SelectedItem.SubItems(8) = "", " ", lvwRepuestos.SelectedItem.SubItems(8))
''                    Tabla.Update
''                Next i
''                Tabla.Close
''
''                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_TOTALES")
''
''                    Tabla.AddNew
''                    Tabla!TManoObra = mcurTMec
''                    Tabla!TRepuestos = mcurTRep
''                    Tabla!TPyP = mcurTCar
''                    Tabla!TTerceros = mcurTTer
''                    Tabla!Insumos = mcurTIns
''                    Tabla!Lubricantes = mcurTLub
''                    Tabla!TOtros = mcurTOtr
''                    Tabla!SubTotal = mcurTNeto
''                    Tabla!IVA = Round(mcurTNeto * 0.18, 2)
''                    Tabla!Total = Round(1.18 * mcurTNeto, 2)
''                    Tabla.Update
''
''                Tabla.Close
''
''                Dbsnueva.Close
''
''
''
''        If mcurTNeto > 0 Then
''            With rptOT
''                            Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
''                            Me.cdImpresora.CancelError = True
''                            Me.cdImpresora.Action = 5
''
''                            .CopiesToPrinter = cdImpresora.Copies
''
''
''
''                If gstrServiciosMarca = "S" Then
''                    .ReportFileName = gstrPathReporte & "\OTPresupuestoMM.rpt"
''                Else
''                    .ReportFileName = gstrPathReporte & "\PresuPrueba.rpt"
''                End If
''
''                .Destination = crptToWindow
''                .WindowState = crptMaximized
''                .DataFiles(0) = gstrPathReporte & "\BDNuevaPresu.mdb"
''
'''                .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
'''                .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
'''                .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
'''                .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
'''                .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
'''                .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
'''                .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
'''
'''                .Formulas(7) = "TMecanica=" & mcurTMec & ""
'''                .Formulas(8) = "TOtros=" & mcurTOtr & ""
'''                .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
'''                .Formulas(10) = "TRepuesto=" & mcurTRep - (mcurTLub + gcurMateriales + mcurTIns) & ""
'''                .Formulas(11) = "TDyP=" & mcurTCar & ""
'''                .Formulas(12) = "TTerceros=" & mcurTTer & ""
'''
'''                .Formulas(13) = "TMateriales=" & gcurMateriales     '& IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
'''                .Formulas(14) = "TInsumos=" & mcurTIns              '& IIf(mstrIdCargo = "01", gcurInsumo, 0) & ""
'''                .Formulas(20) = "TLubricantes=" & mcurTLub & ""
'''                .Formulas(21) = "TelefonoE='Fono: " & gstrTelefono & " Fax: " & gstrFax & "'"
'''
'''                If mstrIdCargo = gstrCargoDeducibleMenos Then
'''                    If mcurDeducible <= mcurTNeto Then
'''                        .Formulas(15) = "Anexo= 'Deducible ( - )'"
'''                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
'''
'''                    End If
'''                ElseIf mstrIdCargo = gstrCargoDeducibleMas Then
'''                        .Formulas(15) = "Anexo= 'Deducible ( + )'"
'''                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
'''
'''                End If
'''                .Formulas(17) = "TNetoOT=" & mcurTNeto & ""
'''                .Formulas(18) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
'''                .Formulas(19) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
'''                .Formulas(22) = "NombreIva='" & gstrNombreIva & "'"
'''                .Formulas(23) = "Tdecimal=" & gintDecimalesMoneda & ""
'''                .Formulas(24) = "NombreRut='" & gstrNombreRut & "'"
'''                .Formulas(25) = "NombrePatente='" & gstrNombrePatente & "'"
'''                .Formulas(26) = "EditaRut='" & gstrEditaRut & "'"
'''                .Formulas(27) = "FamiliaInsumos='" & gstrCodigoInsumos & "'"
'''                .Formulas(28) = "FamiliaLubricantes='" & gstrCodigoLubricantes & "'"
'''                .Formulas(29) = "FamiliaMateriales='" & gstrCodigoMateriales & "'"
'''                .Connect = "Driver={SQL Server};Server=wiracocha;UID=sa;PWD=Llosa1936;Database=elisa;" 'Conexion.ConnectionString
'''                .SelectionFormula = "{Tllr_OT.Id_Empresa}='" & gstrIdEmpresa & "' And {Tllr_OT.Id_Sucursal}='" & gstrIdSucursal & "' And {Tllr_OT.Id_OT}='" & lblNroRecepcion & "' And {Tllr_OT.Seccion_OT}='" & gstrSeccion & "'"
''                .Destination = crptToWindow
''                .Action = True
''            End With
''            mcurTMec = 0
''            mcurTOtr = 0
''            mcurTCar = 0
''            mcurTTer = 0
''            mcurTRep = 0
''            mcurTLub = 0
''            mcurTNeto = 0
''        End If
''  End If
''End If
''
''Solucion:
''    If Err.Number = 32755 Then
''        MsgBox "Impresión Cancelada por el usuario", vbInformation, "Advertencia"
''        Screen.MousePointer = 1
''        Exit Sub
''    End If
''    If Err.Number <> 0 Then
''        MsgBox "Se ha producido el siguiente error " & Chr(13) & Err.Number & " " & Err.Description, vbExclamation, "Advertencia"
''        Screen.MousePointer = 1
''        Exit Sub
''    End If
''End Sub
''Sub ImprimeCompletaSinDeducible()
''Dim mstrIdCargo As String
''Dim mcurTNeto As Currency
''Dim mcurTMec As Currency
''Dim mcurTOtr As Currency
''Dim mcurTCar As Currency
''Dim mcurTTer As Currency
''Dim mcurTRep As Currency
''Dim mcurTMat As Currency
''Dim mcurTIns As Currency
''Dim mcurDeducible As Currency
''    mstrIdCargo = ""
''    mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssMec)
''    mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssOtr)
''    mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssCar)
''    mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssTer)
''    mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep) - IIf(mstrIdCargo = "01", gcurMateriales, 0)
''    mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = "01", gcurMateriales, 0) + IIf(mstrIdCargo = "01", gcurInsumo, 0)
''
''    With rptOT
''        .ReportFileName = gstrPathReporte & "\OTSTD" & ".rpt"
''        .Destination = crptToPrinter
''        .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
''        .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
''        .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
''        .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
''        .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
''        .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
''        .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
''
''        .Formulas(7) = "TMecanica=" & mcurTMec & ""
''        .Formulas(8) = "TOtros=" & mcurTOtr & ""
''        .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
''        .Formulas(10) = "TRepuesto=" & mcurTRep & ""
''        .Formulas(11) = "TDyP=" & mcurTCar & ""
''        .Formulas(12) = "TTerceros=" & mcurTTer & ""
''
''        .Formulas(13) = "TMateriales=" & IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
''        .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = "01", gcurInsumo, 0) & ""
''        .Formulas(15) = "TNetoOT=" & mcurTNeto & ""
''        .Formulas(16) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
''        .Formulas(17) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
''        .Action = True
''    End With
''End Sub
''
''Function ValidaDatos() As Boolean
''Dim j As Integer
''Dim i As Integer
''Dim cont As Integer
''Dim tablaParam As New ADODB.Recordset
''Dim lstrSQL As String
''Dim SW As Integer
''Dim val_real As Double
''cont = 0
''
''ValidaDatos = True
''
'''kjcv 19.01.16
''    For i = 1 To Me.lvwServiciosTerceros.ListItems.Count
''        If Trim(Me.lvwServiciosTerceros.ListItems(i).SubItems(4)) = "" Then
''            MsgBox "No existe Numero Factura en la Línea " & i & " de los Servicios de Terceros" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''    'kjcv 20.07.16
''    'Asignacion de Mecanico
''    'mecanica
''    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
''        If Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(8)) = gstrMecanicoDefectoSecMec Then
''            MsgBox "Debe asignar un Mecánico de los Servicios de Mecánica" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''     'otros servicios
''    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
''        If Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(8)) = gstrMecanicoDefectoSecMec Then
''            MsgBox "Debe asignar un Mecánico de los Otros Servicios" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''    'kjcv 20.07.16
''    'Valida nro Horas
''    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
''        If Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(2)) = "0.0" And Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(6)) <> gstrIdCargoInterno Then
''            MsgBox "Debe ingresar Nro Horas en Servicios de Mecánica" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''    'otros servicios
''    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
''        If Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(2)) = "0.00" And Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(6)) <> gstrIdCargoInterno Then
''            MsgBox "Debe ingresar Nro Horas en Otros Servicios" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''
'''valida subtotales en cero (0) según parametro
''If gblnValidaServiciosCero = True Then
''
''    'mecanica
''    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
''        If Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(10)) = "0" Then
''            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Servicios de Mecanica" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''    'carrocería
''    For i = 1 To Me.lvwServiciosCarroceria.ListItems.Count
''        If Trim(Me.lvwServiciosCarroceria.ListItems(i).SubItems(16)) = "0" Then
''            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Servicios de Carrocería" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''    'otros servicios
''    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
''        If Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(10)) = "0" Then
''            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Otros Servicios" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''    'terceros
''    For i = 1 To Me.lvwServiciosTerceros.ListItems.Count
''        If Trim(Me.lvwServiciosTerceros.ListItems(i).SubItems(12)) = "0" Then
''            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Servicios de Terceros" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''
''
''
''    'repuestos
''    For i = 1 To Me.lvwRepuestos.ListItems.Count
''        If Trim(Me.lvwRepuestos.ListItems(i).SubItems(8)) = "0" Then
''            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Repuestos" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
''            ValidaDatos = False
''            Exit Function
''        End If
''    Next
''End If
''
''With lvwRepuestos
''    i = 1
''    j = .ListItems.Count
''    For i = 1 To .ListItems.Count
''
''       If Trim(lvwRepuestos.ListItems(j).SubItems(11)) = "PRESUPUESTO" Then
''
''            If Trim(lvwRepuestos.ListItems(j).SubItems(2)) <> Trim(lvwRepuestos.ListItems(j).SubItems(13)) Then
''                SW = 1
''                val_real = Val(Trim(lvwRepuestos.ListItems(j).SubItems(2)))
''            End If
''
''            If MsgBox("El Repuesto " & Me.lvwRepuestos.ListItems(j).SubItems(1) & " No esta Descontado de Stock-Pro" & Chr(13) & "¿Desea Eliminarlo de la OT?", vbQuestion + vbYesNo + vbDefaultButton2, "Advertencia") = vbYes Then
''                lvwRepuestos.ListItems.Remove (j)
''                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''                TotalFinal
''            End If
''       End If
''
''        j = j - 1
''    Next
''
''End With
''
''If SW = 1 Then
''    GrabarRegistro
''    MsgBox "Se han actualizado las cantidades ", vbInformation
''End If
''
''End Function
''
''Sub EstadosOT(ModeAction As gAccionEstadoOT)
''
''Dim SW As Integer
''SW = 1
''gflag = False
''
''If ModeAction = gOTActivar Then
''    '//////////////////////////////////////VERIFICAR
''    Act = 1
''    If VeriLiq() = True And gflag = True Then
''        gstrSql = "UPDATE TLLR_OT SET ESTADO = 'V' ,"
'''        gstrSql = gstrSql & "Fecha_Activacion = '" & CDate(pckFechaAtencion.Value) & "' , "
'''kjcv 28.05.13 Graba la fecha en que se genera la activacion
''        gstrSql = gstrSql & "Fecha_Activacion = '" & CDate(Now) & "' , "
''        'kjcv 06.06.16
''        gstrSql = gstrSql & "Usr_Activacion = '" & gUsr_Activacion & "' ,"
''        gstrSql = gstrSql & "Fecha_Activa = '" & CDate(Now) & "' , "
''
''        gstrSql = gstrSql & "Quien_Activa = '" & gstrIdUsuario & "' "
''        gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_OT.Id_OT = '" & lblNroRecepcion & "' AND Tllr_OT.Seccion_OT = '" & gstrSeccion & "' "
''        If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''            lblEstadoOTValor = "VIGENTE"
''            tlbBarraHerramientas.Buttons.item(2).Enabled = True     'guardar
''            tlbBarraHerramientas.Buttons.item(13).Enabled = False   'ACTIVAR
''            tlbBarraHerramientas.Buttons.item(14).Enabled = True    'ANULAR
''            tlbBarraHerramientas.Buttons.item(15).Enabled = True    'LIQUIDAR
''        End If
''        EliminaRegistros gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, gstrSeccion
''        MsgBox "La OT Nº " & lblNroRecepcion & " Fue Activada"
''        Bloqueo "V"
''    Else
''        MsgBox "Lo siento, La Contraseña Ingresada no es la Correcta"
''    End If
''ElseIf ModeAction = gOTAnular Then
'''kjcv 10.02.14
'''Validacion si tiene perfil para anular OT, se creo nuevo perfil desde BD opcion_sistema
''
''    If Not Atributos("Glbl", "Tllr_20_0170", False, False, False, False) Then
''        MsgBox "Ud. No cuenta con Acceso para realizar esta operación...", vbInformation, "Advertencia"
'''        Unload Me
''        Exit Sub
''    End If
'''kjcv 28.04.15 se comenta diferencia del total- insumo
''    'stbTotalOT.Panels(2) = CDbl(stbTotalOT.Panels(2)) - gcurInsumo
''    If stbTotalOT.Panels(2) <= 0 Then  ' valida que no existan valores cargados a la OT
''
''        If VeriLiq() = True Then
''            gstrSql = "UPDATE TLLR_OT SET ESTADO = 'N' ,"
'''            gstrSql = gstrSql & "Fecha_Anulacion = '" & CDate(pckFechaAtencion.Value) & "' , "
''            'kjcv 28.05.13 Graba la fecha en que se genera la activacion
''            gstrSql = gstrSql & "Fecha_Anulacion = '" & CDate(Now) & "' , "
''            gstrSql = gstrSql & "Quien_Anula = '" & gstrIdUsuario & "' "
''            gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_OT.Id_OT = '" & lblNroRecepcion & "' AND Tllr_OT.Seccion_OT = '" & gstrSeccion & "' "
''            If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''                lblEstadoOTValor = "NULA"
''                tlbBarraHerramientas.Buttons.item(2).Enabled = False 'guardar
''                tlbBarraHerramientas.Buttons.item(13).Enabled = True    'ACTIVAR
''                tlbBarraHerramientas.Buttons.item(14).Enabled = False 'ANULAR
''                tlbBarraHerramientas.Buttons.item(15).Enabled = False  'LIQUIDAR
''            End If
''            MsgBox "La OT Nº " & lblNroRecepcion & " Fue Anulada"
''        Else
''            MsgBox "Lo siento, La Contraseña Ingresada no es la Correcta"
''        End If
''    Else
''        MsgBox "No puede Anular una OT que Tenga Valor mayor que 0", vbExclamation, "Anular OT"
''    End If
''ElseIf ModeAction = gOTLiquidar And SW = 1 Then
''
''    Dim lcurInsumos As Double
''
''    'guardo el parametro, porque mas adelante si lo cambia lo hace en la variable global
''    lcurInsumos = gcurInsumo
''
''    If ValidaDatos = False Then
''        Exit Sub
''    End If
''    Act = 0
''
''
''    frmLiquidacion.Show 1
''    If gblnCierraLiq = True Then
''        GrabarRegistro
''        If VeriLiq() = True And gflag = True Then
''            EliminaRegistros gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, gstrSeccion
''            gstrSql = "UPDATE TLLR_OT SET ESTADO = 'L' ,"
''            gstrSql = gstrSql & "Fecha_Liquidacion = '" & CDate(Format(Now, "dd/mm/yyyy")) & "' , "
''            gstrSql = gstrSql & "Quien_Liquida = '" & gstrIdUsuario & "' ,"
''            gstrSql = gstrSql & "Total_Insumos=" & gcurInsumo & " ,"
''            gstrSql = gstrSql & "Total_Materiales=" & gcurMateriales & " ,"
''            gstrSql = gstrSql & "Total_Iva=" & Round(gcurTotalIVA, gintDecimalesMoneda) & " ,"
''            gstrSql = gstrSql & "Total_OT_IVA=" & Round(gcurTotalNetoMasIVA, gintDecimalesMoneda) & " ,"
''            gstrSql = gstrSql & "Total_OT=" & Round(gcurTotalNeto, gintDecimalesMoneda) & " "
''            gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_OT.Id_OT = '" & lblNroRecepcion & "' AND Tllr_OT.Seccion_OT = '" & gstrSeccion & "' "
''            If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''                lblEstadoOTValor = "LIQUIDADA"
''                tlbBarraHerramientas.Buttons.item(2).Enabled = False 'guardar
''                tlbBarraHerramientas.Buttons.item(13).Enabled = True 'ACTIVAR
''                tlbBarraHerramientas.Buttons.item(14).Enabled = False 'ANULAR
''                tlbBarraHerramientas.Buttons.item(15).Enabled = False 'LIQUIDAR
''                gstrImpresion = "O"
''                Dim FechaLiquidacion As Date
''                FechaLiquidacion = CDate(Format(Now, "dd/mm/yyyy"))
''                GeneraRegistroFactura gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, gstrSeccion, txtPatente, lblMarca, lblModelo, lblCliente, gcurInsumo, gcurMateriales, gcurSeguroTaller, lblIdCliente, FechaLiquidacion
''
''                'actualizar datos de rent a car
''                If Me.dtcGarantia.BoundText = "REN" And Me.optMantencion.Value = True Then
''                    gstrEstadoDisponible = Retorna_Valor_General("Select EstadoDisponible from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
''                    gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoDisponible & "', "
''                    gstrSql = gstrSql & " KilometrajeActual=" & Me.txtKilAct
''                    gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
''                    If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''                    End If
''
''                    'actualiza valores en auto stock
''                    gstrSql = "UPDATE Rent_Anexo_Auto_Stock SET Fecha_Ultima_Mantencion='" & Me.pckFechaEntrega & "',"
''                    gstrSql = gstrSql & " Kilometraje_Ultima_Mantencion='" & Me.txtKilAct & "'"
''                    gstrSql = gstrSql & " Where Id_Cajon_Pedido='" & Me.lblVin & "'"
''                    If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''                    End If
''
''                End If
''                If Me.dtcGarantia.BoundText = "REN" And Me.optReparacion.Value = True Then
''                    gstrEstadoDisponible = Retorna_Valor_General("Select EstadoDisponible from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
''                    gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoDisponible & "'"
''                    gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
''                    If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''                    End If
''                End If
''
''                PrintOT
''
''                'vuelve al estado original del parametro
''                gcurInsumo = lcurInsumos
''
''            End If
''            MsgBox "La OT Nº " & lblNroRecepcion & " Fue Liquidada"
''            Bloqueo "L"
''        Else
''            MsgBox "Lo siento, La Contraseña Ingresada no es la Correcta"
''        End If
''    End If
''Else
''    DoEvents
''End If
''End Sub
''
''
''Function CalculoMateriales(IndiceSubItem As Integer) As Double
''Dim intS As Integer
''Dim dblPreSuma As Double
''dblPreSuma = 0
''With lvwRepuestos
''    For intS = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intS)
''        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoMateriales Then '"85"
''            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
''        End If
''    Next
''End With
''CalculoMateriales = dblPreSuma
''End Function
''Function CalculoInsumos(IndiceSubItem As Integer) As Double
''Dim intS As Integer
''Dim dblPreSuma As Double
''dblPreSuma = 0
''With lvwRepuestos
''    For intS = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intS)
''        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoInsumos Then '"80"
''            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
''        End If
''    Next
''End With
''CalculoInsumos = dblPreSuma
''End Function
''Function CalculoLubricantes(IndiceSubItem As Integer) As Double
''Dim intS As Integer
''Dim dblPreSuma As Double
''dblPreSuma = 0
''With lvwRepuestos
''    For intS = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intS)
''        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoLubricantes Then '"90"
''            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
''        End If
''    Next
''End With
''CalculoLubricantes = dblPreSuma
''End Function
''
''Sub LimpiaLinea()
''With Me
''    .dtcConceptos.BoundText = ""
''    .txtSeccion = ""
''    .dtcPartePieza.BoundText = ""
''    .txtHorasCar = ""
''    .txtValorDefCar = ""
''    .txtPorcDesCar = ""
''    .txtMtoDesCar = ""
''    .txtValorFinCar = ""
''    '.dtcCargoCar.BoundColumn = ""
''    .dtcMecanicoCar.BoundText = ""
''    .dtcConceptos.SetFocus
''End With
''End Sub
''
''
''Sub ServicioCarroceria(Accion As mAccionItem)
''If Accion = mAddItem Then
''    Set itmAux = lvwServiciosCarroceria.ListItems.Add(, , dtcConceptos.Text)
''    Set lvwServiciosCarroceria.SelectedItem = itmAux
''    itmAux.SubItems(1) = dtcConceptos.BoundText
''    itmAux.SubItems(2) = txtSeccion.Text
''    itmAux.SubItems(3) = dtcPartePieza.Text
''    itmAux.SubItems(4) = dtcPartePieza.BoundText
''    itmAux.SubItems(5) = FormatoValor(IIf(txtHorasCar <> "", txtHorasCar, 0), "", gintDecimalesMoneda)
''    itmAux.SubItems(6) = FormatoValor(IIf(txtValorDefCar <> "", txtValorDefCar, 0), "", gintDecimalesMoneda)
''    itmAux.SubItems(7) = FormatoValor(IIf(txtPorcDesCar <> "", txtPorcDesCar, 0), "", 2)
''    itmAux.SubItems(8) = FormatoValor(IIf(txtMtoDesCar <> "", txtMtoDesCar, 0), "", gintDecimalesMoneda)
''    itmAux.SubItems(9) = FormatoValor(IIf(txtValorFinCar <> "", txtValorFinCar, 0), "", gintDecimalesMoneda)
''    itmAux.SubItems(10) = IIf(dtcCargoCar = "", TraeCargoDes(gstrIdCargo), dtcCargoCar.Text)
''    itmAux.SubItems(11) = IIf(dtcCargoCar = "", gstrIdCargo, dtcCargoCar.BoundText)
''    itmAux.SubItems(12) = dtcMecanicoCar.Text  'TraeNombreMecanico(gstrMecanicoDefectoSecCar)
''    itmAux.SubItems(13) = dtcMecanicoCar.BoundText  'gstrMecanicoDefectoSecCar
''    itmAux.SubItems(14) = FormatoValor(CalculoSubTotal(mcFichaCarroceria), "", gintDecimalesMoneda)
''    itmAux.SubItems(15) = "N"
''End If
''If Accion = mDelItem Then
''    If lvwServiciosCarroceria.ListItems.Count > 0 Then
''        If Me.lvwServiciosCarroceria.SelectedItem.SubItems(17) = "N" Then
''            lvwServiciosCarroceria.ListItems.Remove lvwServiciosCarroceria.SelectedItem.Index
''        End If
''    End If
''End If
''End Sub
''Sub AsignaTotal(Seccion As mcFicha, Objeto As statusBar)
''Dim Resta As Double
''
''If Seccion = mcFichaMecanica Then '///////////total mecanica
''    With Objeto
''        .Panels(2).Text = FormatoValor(TotalSeccion(lvwServiciosMecanica, 10), "", gintDecimalesMoneda)
''    End With
''ElseIf Seccion = mcFichaCarroceria Then '///////////total carroceria
''    With Objeto
''        .Panels(2).Text = FormatoValor(TotalSeccion(lvwServiciosCarroceria, 16), "", gintDecimalesMoneda)
''        stbTotalDesabolladura.Panels(2).Text = FormatoValor(SubTotalDesabolladura, "", gintDecimalesMoneda)
''        stbTotalPintura.Panels(2).Text = FormatoValor(SubTotalPintura, "", gintDecimalesMoneda)
''        stbTotalArmeyDesarme.Panels(2).Text = FormatoValor(SubTotalArmeDesarme, "", gintDecimalesMoneda)
''    End With
''ElseIf Seccion = mcFichaTerceros Then '///////////total terceros
''    With Objeto
''        .Panels(2).Text = FormatoValor(TotalSeccion(lvwServiciosTerceros, 12), "", gintDecimalesMoneda)
''    End With
''ElseIf Seccion = mcFichaRepuestos Then '///////////total repuestos
''    With Objeto
''        gcurMateriales = CalculoMateriales(8)
''        gcurLubricantes = CalculoLubricantes(8)
''        Resta = CalculoInsumos(8) + gcurMateriales + gcurLubricantes
''        .Panels(2).Text = FormatoValor(TotalSeccion(lvwRepuestos, 8) - Resta, "", gintDecimalesMoneda)
''        stbTotalMateriales.Panels(2).Text = FormatoValor(gcurMateriales, "", gintDecimalesMoneda)   '// sumo insumos a materiales
''        StbLubricantes.Panels(2).Text = FormatoValor(gcurLubricantes, "", gintDecimalesMoneda)
''        stbInsumos.Panels(2).Text = FormatoValor(gcurInsumo + CalculoInsumos(8), "", gintDecimalesMoneda)
''    End With
''ElseIf Seccion = mcFichaOtros Then '///////////total otros
''    With Objeto
''        .Panels(2).Text = FormatoValor(TotalSeccion(lvwOtrosServicios, 10), "", gintDecimalesMoneda)
''    End With
''End If
''End Sub
''Sub LimpiaTotales()
''With Me
''    .stbTotalMec.Panels(2).Text = "0"
''    .stbTotalCarroceria.Panels(2).Text = "0"
''    .stbTotalDesabolladura.Panels(2).Text = "0"
''    .stbTotalPintura.Panels(2).Text = "0"
''    .stbTotalArmeyDesarme.Panels(2) = "0"
''    .stbTotalOtros.Panels(2).Text = "0"
''    .stbTotalTerceros.Panels(2).Text = "0"
''    .stbTotalRepuestos.Panels(2).Text = "0"
''    .stbTotalMateriales.Panels(2).Text = "0"
''    .StbLubricantes.Panels(2).Text = "0"
''    .stbTotalOT.Panels(2).Text = "0"
''    .stbInsumos.Panels(2).Text = "0"
''End With
''End Sub
''
''Function SubTotalDesabolladura() As Double
''Dim intS As Integer
''Dim dblPreSuma As Double
''
''dblPreSuma = 0
''With lvwServiciosCarroceria
''    For intS = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intS)
''        If .SelectedItem.SubItems(3) = "D" Then
''            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(.SelectedItem.SubItems(16), ""))
''        End If
''    Next
''End With
''SubTotalDesabolladura = dblPreSuma
''End Function
''
''Function SubTotalPintura() As Double
''Dim intS As Integer
''Dim dblPreSuma As Double
''
''dblPreSuma = 0
''With lvwServiciosCarroceria
''    For intS = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intS)
''        If .SelectedItem.SubItems(3) = "P" Then
''            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(.SelectedItem.SubItems(16), ""))
''        End If
''    Next
''End With
''SubTotalPintura = dblPreSuma
''
''End Function
''Function SubTotalArmeDesarme() As Double
''Dim intS As Integer
''Dim dblPreSuma As Double
''
''dblPreSuma = 0
''With lvwServiciosCarroceria
''    For intS = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intS)
''        If .SelectedItem.SubItems(3) = "A" Then
''            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(.SelectedItem.SubItems(16), ""))
''        End If
''    Next
''End With
''SubTotalArmeDesarme = dblPreSuma
''
''End Function
''
''Sub TotalFinal()
''    stbTotalOT.Panels(2).Text = FormatoValor(TotalOT, "", gintDecimalesMoneda)
''End Sub
''
''Function TotalOT() As Double
''Dim dblSemiTotal As Double
''With Me
''    dblSemiTotal = Val(SacarFormatoValor(.stbTotalMec.Panels(2).Text, ""))
''    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalCarroceria.Panels(2).Text, ""))
''    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalOtros.Panels(2).Text, ""))
''    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalTerceros.Panels(2).Text, ""))
''    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalRepuestos.Panels(2).Text, ""))
''    'dblSemiTotal = dblSemiTotal + IIf(Not IsNull(gcurInsumo), gcurInsumo, 0)
''    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbInsumos.Panels(2).Text, ""))
''    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalMateriales.Panels(2).Text, ""))
''    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.StbLubricantes.Panels(2).Text, ""))
''End With
''TotalOT = dblSemiTotal
''End Function
''Function CalculoSubTotal(Ficha As mcFicha) As Double
''Dim Total As Double
''
''Total = 0
''If Ficha = mcFichaMecanica Then
''    With lvwServiciosMecanica
''        If .ListItems.Count > 0 Then
''        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(2), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))
''        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
''        End If
''    End With
''ElseIf Ficha = mcFichaCarroceria Then
''    With lvwServiciosCarroceria
''        If .ListItems.Count > 0 Then
''            Total = Val(SacarFormatoValor(.SelectedItem.SubItems(5), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(9), ""))
''            Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(8), ""))
''        End If
''    End With
''ElseIf Ficha = mcFichaTerceros Then
''    With lvwServiciosTerceros
''        If .ListItems.Count > 0 Then
''        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(4), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
''        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(7), ""))
''        End If
''    End With
''ElseIf Ficha = mcFichaRepuestos Then
''    With lvwRepuestos
''        If .ListItems.Count > 0 Then
''        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(2), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))
''        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
''        End If
''    End With
''ElseIf Ficha = mcFichaOtros Then
''    With lvwOtrosServicios
''        If .ListItems.Count > 0 Then
''        Total = Val(SacarFormatoValor(.SelectedItem.SubItems(2), "")) * Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))
''        Total = Total - Val(SacarFormatoValor(.SelectedItem.SubItems(5), ""))
''        End If
''    End With
''End If
''    CalculoSubTotal = Total
''End Function
''
''Function TotalSeccion(lvwObjeto As ListView, IndiceSubItem As Integer) As Double
''Dim intS As Integer
''Dim dblPreSuma As Double
''dblPreSuma = 0
''With lvwObjeto
''    For intS = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intS)
''        dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
''    Next
''End With
''TotalSeccion = dblPreSuma
''End Function
''Function DatosCliente(strIdCliente As String) As Boolean
''If strIdCliente <> "" Then
''    mstrSql = "SELECT Glbl_Cliente_Proveedor.Razon_Social as NOMBRE, Glbl_Cliente_Proveedor.Direccion AS DIREC, Glbl_Comuna.Descripcion AS COMUNA, Glbl_Cliente_Proveedor.Rut AS RUT ,Glbl_Cliente_Proveedor.Telefono AS FONO FROM Glbl_Cliente_Proveedor INNER JOIN Glbl_Comuna ON Glbl_Cliente_Proveedor.Id_Comuna = Glbl_Comuna.Id_Comuna "
''    mstrSql = mstrSql & " AND Glbl_Cliente_Proveedor.Id_Ciudad = Glbl_Comuna.Id_Ciudad "
''    mstrSql = mstrSql & " Where Glbl_Cliente_Proveedor.Id_Cliente_Proveedor='" & strIdCliente & "'"
''    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''        With adoPrincipal
''            If Not .BOF And Not .EOF Then
''                lblCliente = IIf(Not IsNull(!Nombre), !Nombre, "")
''                txtDir = IIf(Not IsNull(!DirEC), !DirEC, "")
''                txtComuna = IIf(Not IsNull(!Comuna), !Comuna, "")
''                txtRut = IIf(Not IsNull(!rut), !rut, "")
''                lblFono = ValorNulo(!FONO)
''            End If
''        End With
''    End If
''    Conexion.CloseHost adoPrincipal
''End If
''End Function
''
''Function ExisteRegistro(IdCiaSeguro As String, IdConcepto As String, IdPtePza As String) As Boolean
''Dim adoTemp As New ADODB.Recordset
''ExisteRegistro = False
''mstrSql = "SELECT top 1 * From Tllr_CiaSeguro_Concepto_Parte_Pieza"
''mstrSql = mstrSql & " WHERE Id_Compañia_Seguro = '" & IdCiaSeguro & "'  AND Id_Concepto = '" & IdConcepto & "' AND Id_Parte_Pieza = '" & IdPtePza & "'"
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''        ExisteRegistro = True
''    Else
''        mstrSql = "Insert into Tllr_CiaSeguro_Concepto_Parte_Pieza (Id_Compañia_Seguro, Id_Concepto, Id_Parte_Pieza, Valor, Horas) Values ('" & IdCiaSeguro & "' ,'" & IdConcepto & "' ,'" & IdPtePza & "',0,0)"
''        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''            ExisteRegistro = True
''        Else
''            ExisteRegistro = False
''        End If
''    End If
''End If
''End Function
''
''Sub FillInventarioOT(strIdEmpresa As String, strIdSucursal As String, strIdRecepcion As String, strSeccion As String)
''
''SetCheckOff lvwInventario
''
''mstrSql = "Exec Tllr_CargaInventario_Ot " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdRecepcion & "'"
''
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        While Not .EOF
''            Set lvwInventario.SelectedItem = lvwInventario.FindItem(CStr(!Codigo), , , 1)
''            lvwInventario.SelectedItem.Checked = True
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''End Sub
''
''Sub FillMecanicaOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
''
''    lvwServiciosMecanica.ListItems.Clear
''
''    If gstrServiciosMarca = "S" Then
''        mstrSql = "Exec Tllr_CargaServicios_Mecanica_MM " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
''    Else
''        mstrSql = "Exec Tllr_CargaServicios_Mecanica " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
''    End If
''    Screen.MousePointer = 11
''    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        .MoveFirst
''        While Not .EOF
''            Set itmAux = lvwServiciosMecanica.ListItems.Add(, , ValorNulo(!ID))
''            Set lvwServiciosMecanica.SelectedItem = itmAux
''            itmAux.SubItems(1) = ValorNulo(!Descripcion)
''            itmAux.SubItems(2) = FormatoValor(!Horas, "", 1)
''            itmAux.SubItems(3) = FormatoValor(!Valor, "", gintDecimalesMoneda)
''            itmAux.SubItems(4) = FormatoValor(!PORC, "", 2)
''            itmAux.SubItems(5) = FormatoValor(!MONTO, "", gintDecimalesMoneda)
''            itmAux.SubItems(6) = ValorNulo(!IDCARGO)
''            itmAux.SubItems(7) = IIf(ValorNulo(!CARGO) = "", "(Ninguno)", !CARGO)
''            itmAux.SubItems(8) = ValorNulo(!idmec)
''            itmAux.SubItems(9) = IIf(ValorNulo(!mec) = "", "(Ninguno)", !mec)
''            itmAux.SubItems(10) = FormatoValor(!Total, "", gintDecimalesMoneda)
''            itmAux.SubItems(11) = ValorNulo(!Facturado)
''            If ValorNulo(!Facturado) = "N" Then
''                mblnOtFacturada = True
''            End If
''            itmAux.SubItems(13) = ValorNulo(!HorasReales)
''            itmAux.SubItems(14) = ValorNulo(!Id_tarea)
''            itmAux.SubItems(15) = ValorNulo(!estado_tarea)
''
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''End Sub
''Sub FillRepuestosReservados(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String, strTipo As String)
''If strTipo <> "Q" Then
''    lvwRepuestosMantencion.ListItems.Clear
''End If
''
''mstrSql = "SELECT Tllr_Repuestos_Reservados.Id_Item, "
''mstrSql = mstrSql & "Stck_Item.Descripcion, Stck_Item.Id_Familia,Tllr_Repuestos_Reservados.Solicitado, "
''mstrSql = mstrSql & "Tllr_Repuestos_Reservados.Precio_Unitario, Tllr_Repuestos_Reservados.Estado, "
''mstrSql = mstrSql & "Glbl_Familia.Descripcion AS Familia, "
''mstrSql = mstrSql & "Tllr_Repuestos_Reservados.Id_OT "
''mstrSql = mstrSql & "FROM Tllr_Repuestos_Reservados INNER JOIN "
''mstrSql = mstrSql & "Stck_Item ON "
''mstrSql = mstrSql & "Tllr_Repuestos_Reservados.Id_Item = Stck_Item.Id_Item INNER "
''mstrSql = mstrSql & "Join "
''mstrSql = mstrSql & "Glbl_Familia ON "
''mstrSql = mstrSql & "Stck_Item.Id_Familia = Glbl_Familia.Id_Familia "
''mstrSql = mstrSql & " WHERE (Tllr_Repuestos_Reservados.Id_Empresa = '" & strIdEmpresa & "') AND"
''mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Id_Sucursal = '" & strIdSucursal & "') AND"
''mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Id_OT = '" & strIdDocumento & "') AND"
''mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Seccion_OT = '" & strSeccion & "') AND"
''If strTipo <> "Q" Then
''    mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Tipo <> 'Q')"
''Else
''    mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Tipo = 'Q')"
''End If
''
''
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        .MoveFirst
''        While Not .EOF
''            Set itmAux = frmRecepcion.lvwRepuestosMantencion.FindItem(!Id_Item, lvwText, , 0)
''            If itmAux Is Nothing Then   ' Si no hay coincidencia
''                Set itmAux = lvwRepuestosMantencion.ListItems.Add(, , ValorNulo(!Id_Item))
''                Set lvwRepuestosMantencion.SelectedItem = itmAux
''                itmAux.SubItems(1) = ValorNulo(!Descripcion)
''                itmAux.SubItems(2) = FormatoValor(!Solicitado, "", 2)
''                itmAux.SubItems(3) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
''                itmAux.SubItems(4) = ValorNulo(!Familia)
''
''                If Me.lvwServiciosMecanica.ListItems.Count > 0 Then
''                    itmAux.SubItems(5) = Me.lvwServiciosMecanica.SelectedItem.SubItems(6)
''                Else
''                    itmAux.SubItems(5) = gstrIdCargo
''                End If
''
''                If !estado = "S" Then
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HFF0000
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HFF0000
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HFF0000
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HFF0000
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HFF0000
''                   ' lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HFF0000
''                End If
''                If !estado = "P" Then
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HC0&
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
''                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
''                   ' lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
''                End If
''            End If
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''End Sub
''Sub FillRepuestosFaltantes(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
''
''mstrSql = "SELECT Tllr_Repuestos_Faltantes.Id_Item, "
''mstrSql = mstrSql & "Stck_Item.Descripcion, Tllr_Repuestos_Faltantes.Solicitado, "
''mstrSql = mstrSql & "Tllr_Repuestos_Faltantes.Precio_Unitario, "
''mstrSql = mstrSql & "Glbl_Familia.Descripcion AS Familia, "
''mstrSql = mstrSql & "Tllr_Repuestos_Faltantes.Id_OT "
''mstrSql = mstrSql & "FROM Tllr_Repuestos_Faltantes INNER JOIN "
''mstrSql = mstrSql & "Stck_Item ON "
''mstrSql = mstrSql & "Tllr_Repuestos_Faltantes.Id_Item = Stck_Item.Id_Item INNER "
''mstrSql = mstrSql & "Join "
''mstrSql = mstrSql & "Glbl_Familia ON "
''mstrSql = mstrSql & "Stck_Item.Id_Familia = Glbl_Familia.Id_Familia "
''mstrSql = mstrSql & " WHERE (Tllr_Repuestos_Faltantes.Id_Empresa = '" & strIdEmpresa & "') AND"
''mstrSql = mstrSql & " (Tllr_Repuestos_Faltantes.Id_Sucursal = '" & strIdSucursal & "') AND"
''mstrSql = mstrSql & " (Tllr_Repuestos_Faltantes.Id_OT = '" & strIdDocumento & "') AND"
''mstrSql = mstrSql & " (Tllr_Repuestos_Faltantes.Seccion_OT = '" & strSeccion & "')"
''
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        .MoveFirst
''        While Not .EOF
''            Set itmAux = lvwRepuestosMantencion.ListItems.Add(, , ValorNulo(!Id_Item))
''            Set lvwRepuestosMantencion.SelectedItem = itmAux
''            itmAux.SubItems(1) = ValorNulo(!Descripcion)
''            itmAux.SubItems(2) = FormatoValor(!Solicitado, "", 2)
''            itmAux.SubItems(3) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
''            itmAux.SubItems(4) = ValorNulo(!Familia)
''            'itmAux.SubItems(5) = lvwServiciosMecanica.SelectedItem.SubItems(6)
''
''            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HC0&
''            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
''            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
''            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
''            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
''            'lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
''
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''End Sub
''
''
''Sub FillCarroceriaOT(strIdEmpresa As String, strIdSucursal As String, strIdRecepcion As String, strSeccion As String, strIdCiaSeguro As String)
''
''lvwServiciosCarroceria.ListItems.Clear
''
''mstrSql = "Exec Tllr_CargaServicios_Carroceria " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdRecepcion & "'"
''
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        .MoveFirst
''        While Not .EOF
''            Set itmAux = lvwServiciosCarroceria.ListItems.Add(, , "")          '///des concepto
''            itmAux.SubItems(1) = IIf(IsNull(!IDCONCEP), "", !IDCONCEP)                                            '///id concepto
''            itmAux.SubItems(2) = IIf(IsNull(!DescCarr), "", !DescCarr)                   '///d_p
''            itmAux.SubItems(3) = IIf(IsNull(!D_P), "", !D_P)                                               '/// des parte
''            itmAux.SubItems(4) = IIf(IsNull(!IDPARTE), "", !IDPARTE)                                             '///idparte
''            itmAux.SubItems(5) = FormatoValor(!Horas, "", 1)                              '///valor definido Format(ValorNulo(!HORAS), "#0.0")
''            itmAux.SubItems(6) = FormatoValor(!Valor, "", gintDecimalesMoneda)
''            itmAux.SubItems(7) = FormatoValor(!PORCREC, "", 2)
''            itmAux.SubItems(8) = FormatoValor(!MONTOREC, "", gintDecimalesMoneda)
''            itmAux.SubItems(9) = FormatoValor(!DEFINIDO, "", gintDecimalesMoneda)
''            itmAux.SubItems(10) = FormatoValor(!PORC, "", 2)
''            itmAux.SubItems(11) = FormatoValor(!MONTO, "", gintDecimalesMoneda)
''            itmAux.SubItems(12) = !CARGO
''            itmAux.SubItems(13) = !IDCARGO
''            itmAux.SubItems(14) = IIf(ValorNulo(!Provee) = "", "(Ninguno)", !Provee)
''            itmAux.SubItems(15) = ValorNulo(!IDPROV)
''            itmAux.SubItems(16) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
''            itmAux.SubItems(17) = ValorNulo(!Facturado)
''            itmAux.SubItems(18) = IIf(IsNull(!Codigo), 1, !Codigo)
''            If ValorNulo(!Facturado) = "N" Then
''                mblnOtFacturada = True
''            End If
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''End Sub
''Sub FillOtrosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
''
''lvwOtrosServicios.ListItems.Clear
''
''mstrSql = "Exec Tllr_CargaServicios_Otro " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
''
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        .MoveFirst
''        While Not .EOF
''            Set itmAux = lvwOtrosServicios.ListItems.Add(, , !ID)            '///des concepto
''            itmAux.SubItems(1) = !Des                                              '///id concepto
''            itmAux.SubItems(2) = FormatoValor(!TIEMPO, "", 2)                                                 '///d_p
''            itmAux.SubItems(3) = FormatoValor(!UNITARIO, "", gintDecimalesMoneda)                                               '/// des parte)
''            itmAux.SubItems(4) = FormatoValor(!PORCDESC, "", 2)                                 '///valor definido Format(ValorNulo(!HORAS), "#0.0")
''            itmAux.SubItems(5) = FormatoValor(!MTODESC, "", gintDecimalesMoneda)
''            itmAux.SubItems(6) = !IDCARGO
''            itmAux.SubItems(7) = TraeCargoDes(!IDCARGO)
''            itmAux.SubItems(8) = ValorNulo(!idmec)
''            itmAux.SubItems(9) = MecanicoD(ValorNulo(!idmec))
''            itmAux.SubItems(10) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
''            itmAux.SubItems(11) = ValorNulo(!Facturado)
''            If ValorNulo(!Facturado) = "N" Then
''                mblnOtFacturada = True
''            End If
''            itmAux.SubItems(12) = ValorNulo(!HorasReales)
''            itmAux.SubItems(13) = ValorNulo(!Id_tarea)
''            itmAux.SubItems(14) = ValorNulo(!estado_tarea)
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''
''End Sub
''
''
''Sub FillTercerosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
''
''lvwServiciosTerceros.ListItems.Clear
''
''mstrSql = "Exec Tllr_CargaServicios_Terceros " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
''
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        .MoveFirst
''        While Not .EOF
''            Set itmAux = lvwServiciosTerceros.ListItems.Add(, , !idServicio)            '///des concepto
''            itmAux.SubItems(1) = ValorNulo(!Proveedor)  '///id concepto
''            itmAux.SubItems(2) = ValorNulo(!IDPROV)
''            itmAux.SubItems(3) = ValorNulo(!servicio) '/// des parte
''            itmAux.SubItems(4) = ValorNulo(!NROFACT)
''            itmAux.SubItems(5) = FormatoValor(!PREUNI, "", gintDecimalesMoneda)
''            itmAux.SubItems(6) = FormatoValor(!CANTY, "", 1)                                 '///valor definido Format(ValorNulo(!HORAS), "#0.0")
''            itmAux.SubItems(7) = FormatoValor(!PRECARGO, "", 2)
''            itmAux.SubItems(8) = FormatoValor(!MRECARGO, "", gintDecimalesMoneda)
''            itmAux.SubItems(9) = FormatoValor(!PREFIN, "", gintDecimalesMoneda)
''            itmAux.SubItems(10) = FormatoValor(IIf(IsNull(!PDSCTO), "0", !PDSCTO), "", 2)
''            itmAux.SubItems(11) = FormatoValor(IIf(IsNull(!MDSCTO), "0", !MDSCTO), "", gintDecimalesMoneda)
''            itmAux.SubItems(12) = FormatoValor(!STotal, "", gintDecimalesMoneda)
''            itmAux.SubItems(13) = TraeCargoDes(!IDCARGO)
''            itmAux.SubItems(14) = !IDCARGO
''            itmAux.SubItems(15) = ValorNulo(!Facturado)
''            If ValorNulo(!Facturado) = "N" Then
''                mblnOtFacturada = True
''            End If
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''mstrSql = ""
''End Sub
''Sub FillRepuestosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
''
''lvwRepuestos.ListItems.Clear
''
''mstrSql = "Exec Tllr_CargaServicios_Repuestos " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
''
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''    If Not .BOF And Not .EOF Then
''        .MoveFirst
''        While Not .EOF
''            'If !CanTY > 0 Then  '///valores > 0
''                Set itmAux = lvwRepuestos.ListItems.Add(, , !ID)            '///des concepto
''                itmAux.SubItems(1) = ValorNulo(!item)                                              '///id concepto
''                itmAux.SubItems(2) = FormatoValor(!CANTY, "", 2)
''                itmAux.SubItems(3) = FormatoValor(!Valor, "", gintDecimalesMoneda)
''                itmAux.SubItems(4) = FormatoValor(!PORCDES, "", 2)
''                itmAux.SubItems(5) = FormatoValor(!MTODES, "", gintDecimalesMoneda)
''                itmAux.SubItems(6) = TraeCargoDes(ValorNulo(!IDCARGO))
''                itmAux.SubItems(7) = ValorNulo(!IDCARGO)
''                itmAux.SubItems(8) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
''                itmAux.SubItems(9) = FamiliaRep(!ID)
''                itmAux.SubItems(10) = ValorNulo(!Facturado)
''                itmAux.SubItems(11) = IIf(IsNull(!Consumo), "STOCK", IIf(!Consumo = "C", "STOCK", "PRESUPUESTO"))
''                '//LREYES
''                itmAux.SubItems(12) = FormatoValor(0, "", 0)
''                itmAux.SubItems(13) = FormatoValor(IIf(IsNull(!realy), 0, !realy), "", 2)
''                'kjcv 18.03.16
''                itmAux.SubItems(15) = !PrecioVentaD
''
''                If ValorNulo(!Facturado) = "N" Then
''                    mblnOtFacturada = True
''                End If
''            'End If
''            .MoveNext
''        Wend
''    End If
''    End With
''End If
''Conexion.CloseHost adoPrincipal
''End Sub
''Sub DatosVehiculo(strPatente As String)
''If strPatente <> "" Then
''    mstrSql = "SELECT Tllr_Vehiculo_Cliente.Patente,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Marca AS IDMARCA,"
''    mstrSql = mstrSql & " Glbl_Marca.Descripcion AS MARCA,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Modelo AS IDMODELO,"
''    mstrSql = mstrSql & " Glbl_Modelo.Descripcion AS MODELO,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Año,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Color_Exterior AS IDCOLOR,"
''    mstrSql = mstrSql & " Glbl_Color_Exterior.Descripcion AS COLOR,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Kilometros_Actuales AS KILACT,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Nro_Motor AS MOTOR,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Nro_Chasis AS CHASIS,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.VIN AS VIN,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor AS IDCLI,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Fecha_Venta AS FECVTA,"
''    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Concesionario AS CONCES"
''    mstrSql = mstrSql & " FROM Glbl_Cliente_Proveedor RIGHT OUTER JOIN Glbl_Color_Exterior RIGHT OUTER JOIN Tllr_Vehiculo_Cliente ON Glbl_Color_Exterior.Id_Color_Exterior = Tllr_Vehiculo_Cliente.Id_Color_Exterior LEFT OUTER JOIN Glbl_Modelo LEFT OUTER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca ON Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor"
''    '///NEO
''    'mstrSql = mstrSql & " WHERE Tllr_Vehiculo_Cliente.Patente='" & txtPatente & "'"
''    mstrSql = mstrSql & " WHERE Tllr_Vehiculo_Cliente.Patente='" & strPatente & "'"
''    '///
''    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''            With adoPrincipal
''                lblMarca = ValorNulo(!Marca)
''                lblIdMarca = ValorNulo(!IdMarca)
''                lblModelo = ValorNulo(!Modelo)
''                lblIdModelo = ValorNulo(!IdModelo)
''                lblChasis = ValorNulo(!chasis)
''                lblMotor = ValorNulo(!motor)
''                lblVin = ValorNulo(!VIN)
''                txtAño = ValorNulo(!Año)
''                lblColorE = ValorNulo(!Color)
''                'lblCliente = ValorNulo(!idCLI)
''                txtConcesionario = ValorNulo(!CONCES)
''                pckFecVta.Value = IIf(Not IsNull(!FECVTA), !FECVTA, Now)
''                txtKilAct = IIf(Not IsNull(!kilact), !kilact, "0")
''                lblIdCliente = ValorNulo(!idCLI)
''                KilometrajeEntrada = txtKilAct 'Variable de ileiva 07/02/2001
''            End With
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End If
''End Sub
''Sub FillConceptosInventario()
''mstrSql = "SELECT Id_Estado_Recepcion AS Codigo, Descripcion AS Nombre FROM Tllr_Estado_Recepcion WHERE Vigencia = 'S' Order By Id_Estado_Recepcion"
''If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''    With adoPrincipal
''        If Not .BOF And Not .EOF Then
''            .MoveFirst
''            While Not .EOF
''                Set itmAux = lvwInventario.ListItems.Add(, , !Codigo)
''                itmAux.SubItems(1) = !Nombre
''                .MoveNext
''            Wend
''        End If
''    End With
''End If
''End Sub
''Private Function GuardaCarroceria(strIdDocumento As String, strSeccion As String, strCiaSeguro As String, gParametro As gcParametro) As Boolean
''Dim mstrNombreTabla As String
''
''If gParametro = gcOrdenTrabajo Then
''    mstrNombreTabla = "Tllr_Carroceria_OT"
''ElseIf gParametro = gcPresupuesto Then
''    mstrNombreTabla = "Tllr_Carroceria_Presupuesto"
''End If
''
''GuardaCarroceria = True
''mstrSql = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' "
''If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''    With lvwServiciosCarroceria
''        If .ListItems.Count > 0 Then
''            For intIndice = 1 To .ListItems.Count
''                Set .SelectedItem = .ListItems(intIndice)
''                '/////////////////////////////////////////////////VALIDAR SI EXISTE EN PARENT
''                'If ExisteRegistro(strCiaSeguro, .SelectedItem.SubItems(1), .SelectedItem.SubItems(4)) = True Then
''                    mstrSql = "INSERT INTO " & mstrNombreTabla
''                    mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''                    mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''                    mstrSql = mstrSql & " Id_Compañia_Seguro, "
''                    mstrSql = mstrSql & " Id_Concepto, "
''                    mstrSql = mstrSql & " D_P,"
''                    mstrSql = mstrSql & " Id_Parte_Pieza, "
''                    mstrSql = mstrSql & " Id_Tipo_Cargo, Mecanico_Designado,"
''                    mstrSql = mstrSql & " Horas, Valor,Valor_Definido ,"
''                    mstrSql = mstrSql & " Porcentaje_Descuento,Monto_Descuento,"
'''                    mstrSQL = mstrSQL & " SubTotal,Facturado,Porcentaje_Recargo,Monto_Recargo,Id_Proveedor,Descripcion,Id_Servicio_Carroceria)"
''                    'kjcv 21.05.15
''                    mstrSql = mstrSql & " SubTotal,Facturado,Porcentaje_Recargo,Monto_Recargo,Id_Proveedor,Descripcion,ID_GRUPO_CENTRO_COSTO,Id_Servicio_Carroceria)"
''                    mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "       '///empresa, sucursal
''                    mstrSql = mstrSql & " '" & strIdDocumento & "', '" & strSeccion & "',"                  '///nro ot, seccion
''                    mstrSql = mstrSql & " '" & strCiaSeguro & "', "                                         '///cia seguro
''                    mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(1)) & "', "                      '///concepto
''                    mstrSql = mstrSql & " '" & .SelectedItem.SubItems(3) & "',"                                                   'Trim(.SelectedItem.SubItems(2)) ///d_p
''                    mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(4)) & "', "                      '///parte y pieza
''                    mstrSql = mstrSql & " '" & .SelectedItem.SubItems(13) & "','" & gstrMecanicoDefectoSecCar & "',"            '///mecanico designado
''                    mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(5), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(6), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(9), "######.00"))) & " ,"
''                    mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "######.00"))) & ","
''                    mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(16), "######.00"))) & ",'" & .SelectedItem.SubItems(17) & "',"
''                    mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "######.00"))) & ","
''                    mstrSql = mstrSql & " " & IIf(.SelectedItem.SubItems(15) = "", "NULL" & ",", " '" & .SelectedItem.SubItems(15) & "',")
''                    mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(2)) & "',"
''                    'kjcv 21.05.15
''                    mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(19)) & "',"
''                    mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(18)) & "')"
''                    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                        GuardaCarroceria = False
''                        Exit Function
''                    End If
''                'End If
''            Next
''        Else
''            GuardaCarroceria = True
''        End If
''    End With
''Else
''    GuardaCarroceria = False
''    Exit Function
''End If
''End Function
''Private Function GuardaTerceros(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
''Dim mstrNombreTabla As String
''
''If gParametro = gcOrdenTrabajo Then
''    mstrNombreTabla = "Tllr_Terceros_OT"
''ElseIf gParametro = gcPresupuesto Then
''    mstrNombreTabla = "Tllr_Terceros_Presupuesto"
''End If
''
''GuardaTerceros = True
''mstrSql = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' "
''If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''    With lvwServiciosTerceros
''        If .ListItems.Count > 0 Then
''            For intIndice = 1 To .ListItems.Count
''                Set .SelectedItem = .ListItems(intIndice)
''                mstrSql = "INSERT INTO " & mstrNombreTabla
''                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''                mstrSql = mstrSql & " Id_Proveedor, "
''                mstrSql = mstrSql & " Id_Servicio_Tercero,"
''                mstrSql = mstrSql & " Id_Tipo_Cargo, "
''                mstrSql = mstrSql & " Cantidad,Valor,"
''                mstrSql = mstrSql & " Porcentaje_Recargo,Monto_Recargo,"
''                mstrSql = mstrSql & " Precio_Final,"
''                mstrSql = mstrSql & " Descripcion , NroFarctura, "
''                mstrSql = mstrSql & " SubTotal, Facturado, "
'''                mstrSQL = mstrSQL & " Porcentaje_Dscto, Monto_Dscto)"
''                'kjcv 21.05.15
''                mstrSql = mstrSql & " Porcentaje_Dscto,ID_GRUPO_CENTRO_COSTO, Monto_Dscto)"
''                mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
''                mstrSql = mstrSql & " '" & strIdDocumento & "', '" & strSeccion & "',"
''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(2) & "', "
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem) & "', "
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(14)) & "', "
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(6), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(9), "#####0.00"))) & ","
''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(3) & "', "
''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(4) & "', "
'''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(17) & "', "
''                mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(12), "#####0.00"))) & ",'" & .SelectedItem.SubItems(15) & "',"
'''                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "#####0.00"))) & ")"
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & .SelectedItem.SubItems(16) & "'," & CCur(Val(Format(.SelectedItem.SubItems(11), "#####0.00"))) & ")"
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaTerceros = False
''                    Exit Function
''                End If
''            Next
''        Else
''            GuardaTerceros = True
''        End If
''    End With
''Else
''    GuardaTerceros = False
''    Exit Function
''End If
''End Function
''
''Private Function GuardaOtros(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
''Dim mstrNombreTabla As String
''
''If gParametro = gcOrdenTrabajo Then
''    mstrNombreTabla = "Tllr_Otro_OT"
''ElseIf gParametro = gcPresupuesto Then
''    mstrNombreTabla = "Tllr_Otro_Presupuesto"
''End If
''
''GuardaOtros = True
''mstrSql = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' "
''If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''    With lvwOtrosServicios
''        If .ListItems.Count > 0 Then
''            For intIndice = 1 To .ListItems.Count
''                Set .SelectedItem = .ListItems(intIndice)
''                mstrSql = "INSERT INTO " & mstrNombreTabla
''                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''                mstrSql = mstrSql & " Id_Otro_Servicio, "
''                mstrSql = mstrSql & " Id_Tipo_Cargo,"
''                mstrSql = mstrSql & " Mecanico_Asignado, "
''                mstrSql = mstrSql & " Horas,Valor,"
''                mstrSql = mstrSql & " Porcentaje_Descuento,Monto_Descuento,"
'''                mstrSQL = mstrSQL & " SubTotal,Descripcion_Otro,Facturado,HorasReales,Id_Tarea,Estado_Tarea)"
''                'kjcv 21.05.15
''                mstrSql = mstrSql & " SubTotal,Descripcion_Otro,Facturado,HorasReales,Id_Tarea,ID_GRUPO_CENTRO_COSTO,Estado_Tarea)"
''                '
''                mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
''                mstrSql = mstrSql & " '" & strIdDocumento & "', '" & strSeccion & "',"
''                mstrSql = mstrSql & " '" & .SelectedItem & "', "
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(6)) & "', "
''                mstrSql = mstrSql & " '" & IIf(Trim(.SelectedItem.SubItems(8)) = "", "SIN", Trim(.SelectedItem.SubItems(8))) & "', "
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & UCase(Trim(.SelectedItem.SubItems(1))) & "','" & UCase(Trim(.SelectedItem.SubItems(11))) & "',"
''                If .SelectedItem.SubItems(12) = "" Then
''                    mstrSql = mstrSql & " " & 0 & ","
''                Else
''                    mstrSql = mstrSql & " " & CDbl(.SelectedItem.SubItems(12)) & ","
''                End If
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(13)) & "',"
''                  'kjcv 21.05.15
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(15)) & "',"
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(14)) & "')"
''
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaOtros = False
''                    Exit Function
''                End If
''            Next
''        Else
''            GuardaOtros = True
''        End If
''    End With
''Else
''    GuardaOtros = False
''    Exit Function
''End If
''End Function
''Private Function GuardaRepuestos(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
''Dim mstrNombreTabla As String
''Dim adoTemp As New ADODB.Recordset
''Dim j As Integer
''
'''valida si los repuestos no han sido devueltos
'''y no ha sido refrescada la pantalla
''If gstrProcedencia = "Movimientos" Then
''    j = Me.lvwRepuestos.ListItems.Count
''    For intIndice = 1 To Me.lvwRepuestos.ListItems.Count
''        If Me.lvwRepuestos.ListItems(j).SubItems(11) = "STOCK" Then
''            mstrSql = "Select count(id_item) as Cuenta from Tllr_Repuestos_Ot WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' And Consumo='C' and id_item='" & Me.lvwRepuestos.ListItems(j) & "'"
''            If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''                If adoTemp!cuenta = 0 Then
''                    lvwRepuestos.ListItems.Remove (j)
''                End If
''            End If
''        End If
''        j = j - 1
''    Next
''End If
''
''If gParametro = gcOrdenTrabajo Then
''    mstrNombreTabla = "Tllr_Repuestos_OT"
''ElseIf gParametro = gcPresupuesto Then
''    mstrNombreTabla = "Tllr_Repuestos_Presupuesto"
''End If
''
''GuardaRepuestos = True
''
'''elimina solo si son presupuestos
''mstrSql = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' And Consumo='P'"
''Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''With lvwRepuestos
''    If .ListItems.Count > 0 Then
''        For intIndice = 1 To .ListItems.Count
''            Set .SelectedItem = .ListItems(intIndice)
''            If VerificaRepuesto(.SelectedItem, lblNroRecepcion, strSeccion, mstrNombreTabla) = True Then
''                mstrSql = "UPDATE " & mstrNombreTabla
''                mstrSql = mstrSql & " SET Id_Tipo_Cargo='" & Trim(.SelectedItem.SubItems(7)) & "',"
''                mstrSql = mstrSql & " Cantidad = " & CDbl(Val(Format(.SelectedItem.SubItems(13), "#####0.00"))) & ", "
''                mstrSql = mstrSql & " Valor = " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
''                mstrSql = mstrSql & " cantidad_real = " & CCur(Val(Format(.SelectedItem.SubItems(13), "#####0.00"))) & ","
''                mstrSql = mstrSql & " Porcentaje_Descuento = " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & ","
''                mstrSql = mstrSql & " Monto_Descuento = " & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''                mstrSql = mstrSql & " SubTotal = " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
''                mstrSql = mstrSql & " Facturado = " & UCase(Trim(IIf(.SelectedItem.SubItems(10) = "", "'N'", "'" & .SelectedItem.SubItems(10) & "'"))) & ","
''                mstrSql = mstrSql & " Consumo = '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "',"
''                mstrSql = mstrSql & " ID_GRUPO_CENTRO_COSTO = '" & .SelectedItem.SubItems(14) & "',"
''                'kjcv 05.02.16
''                mstrSql = mstrSql & " PrecioVentaD = '" & Round(.SelectedItem.SubItems(15), 2) & "',"
''                mstrSql = mstrSql & " Saldo = '" & .SelectedItem.SubItems(12) & "'"
''                mstrSql = mstrSql & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
''                mstrSql = mstrSql & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
''                mstrSql = mstrSql & " Id_OT = '" & strIdDocumento & "' AND  "
''                mstrSql = mstrSql & " Seccion_OT = '" & strSeccion & "' AND "
''                mstrSql = mstrSql & " Id_Item = '" & .SelectedItem & "' "
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaRepuestos = False
''                    Exit Function
''                End If
''            Else
''                '///////////////////////////////////VALIDAR SI EXISTE EN PARENT
''                mstrSql = "INSERT INTO " & mstrNombreTabla
''                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''                mstrSql = mstrSql & " Id_Item, "
''                mstrSql = mstrSql & " Id_Tipo_Cargo, "
''                mstrSql = mstrSql & " Cantidad, Valor,"
''                mstrSql = mstrSql & " Porcentaje_Descuento,Monto_Descuento,"
'''                mstrSql = mstrSql & " SubTotal,Facturado,Consumo,Saldo)"
''                'kjcv 27.02.13
'''                mstrSQL = mstrSQL & " SubTotal,Facturado,Consumo,Saldo,precioventaD)"
''                'kjcv 21.05.15
''                mstrSql = mstrSql & " SubTotal,Facturado,Consumo,Saldo,ID_GRUPO_CENTRO_COSTO, precioventaD)"
''                mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
''                mstrSql = mstrSql & " '" & strIdDocumento & "', '" & strSeccion & "',"
''                mstrSql = mstrSql & " '" & .SelectedItem & "', "
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(7)) & "', "
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ",'" & .SelectedItem.SubItems(10) & "',"
''                mstrSql = mstrSql & " '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "',"
'''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(12) & "')"
''  'kjcv 21.05.15
''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(14) & "',"
''                'kjcv 27.02.13
''                 mstrSql = mstrSql & " '" & .SelectedItem.SubItems(12) & "',"
'''                mstrSql = mstrSql & Retorna_Valor_General("Select Precio_Venta From Stck_Item Where Id_Item = '" & .SelectedItem & "'") & ")"
'''kjcv 05.02.16
''                mstrSql = mstrSql & Round(.SelectedItem.SubItems(15), 2) & ")"
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaRepuestos = False
''                    Exit Function
''                End If
''            End If
''        Next
''    Else
''        GuardaRepuestos = True
''    End If
''End With
''End Function
''Private Function GuardaInventario(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
''Dim mstrNombreTabla As String
''
''If gParametro = gcOrdenTrabajo Then
''    mstrNombreTabla = "Tllr_Inventario_OT"
''ElseIf gParametro = gcPresupuesto Then
''    mstrNombreTabla = "Tllr_Inventario_Presupuesto"
''End If
''
''
''GuardaInventario = True
''mstrSql = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND " & mstrNombreTabla & ".ID_OT='" & strIdDocumento & "' and " & mstrNombreTabla & ".Seccion_OT = '" & strSeccion & "'"
''If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''    For intIndice = 1 To lvwInventario.ListItems.Count
''        Set lvwInventario.SelectedItem = lvwInventario.ListItems(intIndice)
''        If lvwInventario.SelectedItem.Checked = True Then
''            mstrSql = "Insert Into " & mstrNombreTabla
''            mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,Id_Estado_Recepcion, Id_OT, Seccion_OT) "
''            mstrSql = mstrSql & " values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "','" & lvwInventario.SelectedItem & "', '" & strIdDocumento & "', '" & strSeccion & "' )"
''            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                GuardaInventario = False
''                Exit Function
''            End If
''        End If
''    Next
''    GuardaInventario = True
''Else
''    GuardaInventario = False
''    Exit Function
''End If
''End Function
''Private Function GuardaMecanica(strIdDocumento As String, gParametro As gcParametro) As Boolean
''Dim mstrNombreTabla As String
''
''If gParametro = gcOrdenTrabajo Then
''    mstrNombreTabla = "Tllr_Mecanica_OT"
''ElseIf gParametro = gcPresupuesto Then
''    mstrNombreTabla = "Tllr_Mecanica_Presupuesto"
''End If
''
''GuardaMecanica = True
''mstrSql = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And ID_OT='" & strIdDocumento & "' And Seccion_OT='" & gstrSeccion & "'"
''If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''    With lvwServiciosMecanica
''        If .ListItems.Count > 0 Then
''            For intIndice = 1 To .ListItems.Count
''            Set .SelectedItem = .ListItems(intIndice)
''            mstrSql = "Insert Into " & mstrNombreTabla
''            mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''            mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''            mstrSql = mstrSql & " Id_Marca, Id_Modelo, "
''            mstrSql = mstrSql & " Id_Servicio, "
''            mstrSql = mstrSql & " Id_Tipo_Cargo,Mecanico_Designado,"
''            mstrSql = mstrSql & " Horas,Valor,"
''            mstrSql = mstrSql & " Porcentaje_Descuento, Monto_Descuento, "
'''            mstrSQL = mstrSQL & " SubTotal, Facturado,HorasReales,Id_Tarea,Estado_Tarea)"
''            'kjcv 21.05.15
''            mstrSql = mstrSql & " SubTotal, Facturado,HorasReales,Id_Tarea,ID_GRUPO_CENTRO_COSTO,Estado_Tarea)"
''            mstrSql = mstrSql & " Values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
''            mstrSql = mstrSql & " '" & strIdDocumento & "', '" & gstrSeccion & "',"
''            mstrSql = mstrSql & " '" & Trim(lblIdMarca) & "','" & Trim(lblIdModelo) & "',"
''            mstrSql = mstrSql & " '" & Trim(.SelectedItem) & "',"
''            mstrSql = mstrSql & " '" & .SelectedItem.SubItems(6) & "'," & IIf(.SelectedItem.SubItems(8) = "", "NULL", " '" & .SelectedItem.SubItems(8) & "' ") & ", "
''            mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & " , " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & " , "
''            mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & " ," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''            mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & .SelectedItem.SubItems(11) & "',"
''            If .SelectedItem.SubItems(13) = "" Then
''                mstrSql = mstrSql & " " & 0 & ","
''            Else
''                mstrSql = mstrSql & " " & CDbl(.SelectedItem.SubItems(13)) & ","
''            End If
''            mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(14)) & "',"
''            'kjcv 21.05.15
''            mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(16)) & "',"
''            mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(15)) & "')"
''
''            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                GuardaMecanica = False
''                Exit Function
''            End If
''            Next
''        Else
''            GuardaMecanica = True
''        End If
''    End With
''Else
''    GuardaMecanica = False
''    Exit Function
''End If
''End Function
''
''Function letSql(strWhere As String, strOrder As String) As String
''mstrSql = "SELECT Top 1 Id_OT, "
''mstrSql = mstrSql & " Seccion_OT, "
''mstrSql = mstrSql & " Patente, "
''mstrSql = mstrSql & " Id_Garantia as TipoOT, "
''mstrSql = mstrSql & " Folio_Garantia,"
''mstrSql = mstrSql & " Id_Tipo_Cono, "
''mstrSql = mstrSql & " Nro_Cono, "
''mstrSql = mstrSql & " RealizadoPor, "
''mstrSql = mstrSql & " Fecha_Emision,"
''mstrSql = mstrSql & " Entrega_Estimada, "
''mstrSql = mstrSql & " Hora_Entrega, "
''mstrSql = mstrSql & " Kilometros_Recepcion, "
''mstrSql = mstrSql & " Nro_Siniestro, "
''mstrSql = mstrSql & " Nro_Poliza,"
''mstrSql = mstrSql & " Liquidador, "
''mstrSql = mstrSql & " Deducible_UF, "
''mstrSql = mstrSql & " Deducible_Pesos, "
''mstrSql = mstrSql & " Id_Compañia_Seguro, "
''mstrSql = mstrSql & " Solicitado_Por, "
''mstrSql = mstrSql & " Total_Mecanica, "
''mstrSql = mstrSql & " Total_Carroceria,"
''mstrSql = mstrSql & " Total_Desabolladura, "
''mstrSql = mstrSql & " Total_Pintura, "
''mstrSql = mstrSql & " Total_Terceros,"
''mstrSql = mstrSql & " Total_Materiales,"
''mstrSql = mstrSql & " Total_Insumos,"
''mstrSql = mstrSql & " Total_Repuestos , "
''mstrSql = mstrSql & " Total_OT, "
''mstrSql = mstrSql & " Estado, "
''mstrSql = mstrSql & " Comentario, "
''mstrSql = mstrSql & " ReparacionMantencion, "
''mstrSql = mstrSql & " Estado_Reserva, "
''mstrSql = mstrSql & " Id_Presupuesto, "
''mstrSql = mstrSql & " Fecha_Liquidacion, "
''mstrSql = mstrSql & " OrdenReparacion, "
''mstrSql = mstrSql & " Nro_Presupuesto_Origen, "
''mstrSql = mstrSql & " NroReferencia, Bencina "
''mstrSql = mstrSql & " ,PDI"
''mstrSql = mstrSql & " From Tllr_OT"
''letSql = mstrSql & " " & strWhere & " " & strOrder
''
''End Function
''
''Private Sub LeerCampos()
''
'''/// inicializa variable para verificar si la ot esta totalmente facturada
''mblnOtFacturada = False
''
''If mblnTablaVacia Then
''    LimpiaCampos
''    Exit Sub
''End If
''With adoPrincipal
''    If !Seccion_OT = "C" Then
''        Me.optRecepcion(1).Value = True
''    Else
''        Me.optRecepcion(0).Value = True
''    End If
''    If !ReparacionMantencion = "M" Then
''        Me.optMantencion.Value = True
''    Else
''        Me.optReparacion.Value = True
''    End If
''    If !Estado_Reserva = "R" Then
''        Me.cmdReserva.Enabled = False
''        Me.cmdAnularReserva.Enabled = True
''    Else
''        Me.cmdReserva.Enabled = True
''        Me.cmdAnularReserva.Enabled = False
''    End If
''    lblNroRecepcion.Text = !Id_OT
''    mstrIdPresupuestoOrigen = ValorNulo(!Id_Presupuesto)
''    lblPresupuesto = ValorNulo(!Nro_Presupuesto_Origen)
''    lblFechaLiquidacion = IIf(!estado <> "N" Or !estado <> "V", ValorNulo(!Fecha_Liquidacion), "")
''    dtcGarantia.BoundText = !TipoOt
''    txtNReferencia = ValorNulo(!NroReferencia)
''    Me.cmbBencina.ListIndex = IIf(IsNull(!Bencina), -1, !Bencina)
''    If !TipoOt = "PRE" Then
''        dtcGarantia.Enabled = False
''    Else
''        dtcGarantia.Enabled = True
''    End If
''    gstrIdCargo = TraeCargo(!TipoOt)
''    dtcTipoCono.BoundText = !Id_Tipo_Cono
''    dtcRecepcionista.BoundText = !RealizadoPor
''    txtNroCono = !Nro_Cono
''
''    pckFechaAtencion.Value = !Fecha_Emision
''    pckFechaEntrega.Value = !Entrega_Estimada
''    cboHora.Text = ValorNulo(!Hora_Entrega)
''
''    txtNroSiniestro = ValorNulo(!Nro_Siniestro)
''    txtNroPoliza = ValorNulo(!Nro_Poliza)
''    txtLiquidador = ValorNulo(!Liquidador)
''    txtOrdenReparacion = ValorNulo(!OrdenReparacion)
''
''    txtDeducibleUF = !Deducible_UF
''    txtDeduciblePesos = !deducible_pesos
''    lblCompañia.Tag = !Id_Compañia_Seguro
''    gstrIdCompañiaSeg = !Id_Compañia_Seguro
''    lblCompañia = CiaSegDes(!Id_Compañia_Seguro)
''
''    txtComentario = !Comentario
''    txtPatente = ValorNulo(!Patente)
''    txtFolioGarantia = !Folio_Garantia
''    txtSolicita = !Solicitado_Por
''    gcurInsumo = !Total_Insumos
''    'gcurMateriales = !Total_Materiales
''    'stbInsumos.Panels(2).Text = FormatoValor(!Total_Insumos, "", 0)
''    'kjcv 12.11.13 para Bloquear el Buscar Placa
''    tlbPatente.Buttons(2).Enabled = False
''    If Not IsNull(!estado) Then
''        If gstrProcedencia = "Movimientos" Then
''            lblEstadoOTValor.Caption = IIf(!estado = "V", "VIGENTE", IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "N", "NULA", IIf(!estado = "F" Or !estado = "B", "EMITIDA", IIf(!estado = "R", "RESERVA", IIf(!estado = "P", "PRESUPUESTO", ""))))))
''            'kjcv 24 10.13
''            txtTipo.Text = IIf(!PDI = "S", "PDI", "")
''            tlbBarraHerramientas.Buttons.item(2).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", True, IIf(!estado = "R", True, IIf(!estado = "P", True, False))))))
''            tlbBarraHerramientas.Buttons.item(13).Enabled = IIf(!estado = "V", False, IIf(!estado = "L", True, IIf(!estado = "N", True, IIf(!estado = "F" Or !estado = "B", False, False))))    'ACTIVAR
''            tlbBarraHerramientas.Buttons.item(14).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, False))))    'ANULAR
''            tlbBarraHerramientas.Buttons.item(15).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", True, False))))    'LIQUIDAR
''            tlbBarraHerramientas.Buttons.item(20).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Separador
''            tlbBarraHerramientas.Buttons.item(21).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Confirmar Reserva
''            tlbBarraHerramientas.Buttons.item(22).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Eliminar Reserva
''            tlbBarraHerramientas.Buttons.item(24).Visible = IIf(!estado = "P", True, False) 'Liquidar presupuesto
''            tlbBarraHerramientas.Buttons.item(25).Visible = IIf(!estado = "P", True, False) 'Liquidar presupuesto
''        Else
'''            tlbBarraHerramientas.Buttons.item(2).Enabled = False
'''kjcv 27.02.13
''            tlbBarraHerramientas.Buttons.item(2).Enabled = True
''        End If
''    End If
''
''    'busca numeros de documentos asociados
''    lblDocumentos = IIf(!estado = "F" Or !estado = "B", NumerosDocumentos(!Id_OT, gstrSeccion), "")
''
''    If ValorNulo(!Patente) <> "" Then DatosVehiculo !Patente
''    txtKilAct = !Kilometros_Recepcion 'trae los kilometros de la OT
''    '/////////////////////////////////////////////////////////////////////////////////
''    FillConceptosVsCiaSeguro dtcConceptos, datConceptos, lblCompañia.Tag
''    '/////////////////////////////////////////////////////////////////////////////////
''    FillInventarioOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
''    '/////////////////////////////////////////////////////////////////////////////////
''    FillMecanicaOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
''    AsignaTotal mcFichaMecanica, stbTotalMec
''
''    '/////////////////////////////////////////////////////////////////////////////////
''    'If !Seccion_OT = "C" Then
''        FillCarroceriaOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion, lblCompañia.Tag
''        AsignaTotal mcFichaCarroceria, stbTotalCarroceria
''    'Else
''    '    lvwServiciosCarroceria.ListItems.Clear
''    '    frmRecepcion.stbTotalCarroceria.Panels(2).Text = 0
''    'End If
''    '/////////////////////////////////////////////////////////////////////////////////
''    FillOtrosOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
''    AsignaTotal mcFichaOtros, stbTotalOtros
''    '/////////////////////////////////////////////////////////////////////////////////
''    FillTercerosOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
''    AsignaTotal mcFichaTerceros, stbTotalTerceros
''    '/////////////////////////////////////////////////////////////////////////////////
''    FillRepuestosOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
''    AsignaTotal mcFichaRepuestos, stbTotalRepuestos
'''    stbTotalMateriales.Panels(2).Text = Format(CalculoMateriales(8))
''    '/////////////////////////////////////////////////////////////////////////////////
''
''    If !Estado_Reserva = "R" Then
''        FillRepuestosReservados gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion, "T"  'tempario
''        FillRepuestosFaltantes gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
''    Else
''        '//// Si no encuentra reserva de repuestos busca los repuestos de los servicios
''        Dim i As Integer
''        lvwRepuestosMantencion.ListItems.Clear
''        For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
''            mstrAgregaPresupuesto = False
''            Repuestos_de_la_Mantencion Me.lblIdMarca, Me.lblIdModelo, lvwServiciosMecanica.ListItems(i), IIf(Me.lvwServiciosMecanica.ListItems(i).SubItems(12) = "S", True, False)
''        Next
''        FillRepuestosReservados gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion, "Q"  'presupuesto
''    End If
''
''    TotalFinal
''    '/////////////////////////////////////////////////////////////////////////////////
''    If ValorNulo(!estado) = "B" Or ValorNulo(!estado) = "F" Then
''        tlbBarraHerramientas.Buttons.item(15).Enabled = mblnOtFacturada  'LIQUIDAR
''        tlbBarraHerramientas.Buttons.item(2).Enabled = mblnOtFacturada   'GUARDAR
''
''    End If
''
''    If !estado = "B" Or !estado = "F" Or !estado = "L" Then
''        Me.stbSeguroTaller.Panels(2).Text = Retorna_Valor_General("Select sum(SeguroTaller) as Seguro from Tllr_Facturacion where id_ot='" & !Id_OT & "' And Seccion_OT='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
''        If Me.stbSeguroTaller.Panels(2).Text = "" Then
''            gcurSeguroTaller = 0
''            Me.stbSeguroTaller.Panels(2).Text = "0"
''        Else
''            gcurSeguroTaller = CDbl(Me.stbSeguroTaller.Panels(2).Text)
''        End If
''    Else
''        Me.stbSeguroTaller.Panels(2).Text = "0"
''    End If
''
''    gstrEstado = ValorNulo(!estado)
''
''    Bloqueo ValorNulo(!estado)
''
''End With
''End Sub
''
''Function VerificaServicioCarroceria(strIdConcepto As String, strIdParte As String) As Boolean
''VerificaServicioCarroceria = True
''For intIndice = 1 To lvwServiciosCarroceria.ListItems.Count
''    Set lvwServiciosCarroceria.SelectedItem = lvwServiciosCarroceria.ListItems(intIndice)
''    If lvwServiciosCarroceria.SelectedItem.SubItems(1) = strIdConcepto Then
''        If lvwServiciosCarroceria.SelectedItem.SubItems(4) = strIdParte Then
''            VerificaServicioCarroceria = False
''            Exit Function
''        Else
''            VerificaServicioCarroceria = True
''        End If
''    Else
''        VerificaServicioCarroceria = True
''    End If
''Next intIndice
''End Function
''
''Private Sub cmdAnularReserva_Click()
''Dim EstadoReserva As String
''Dim AdoAnular As New ADODB.Recordset
''If Me.lvwRepuestosMantencion.ListItems.Count > 0 Then
''
''    '/// valida que la reserva no haya pasado a Consumo
''    EstadoReserva = Retorna_Valor_General("Select Estado_Reserva from Stck_Regularizacion Where Id_OT='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
''    If EstadoReserva = "L" Then
''        MsgBox "Esta Reserva ya paso a ser un Consumo...", vbInformation, "Anular Reserva de Repuestos"
''        Exit Sub
''    End If
''    'Levanta listview con los repuestos de la mantencion
''    If MsgBox(" Esta Seguro de Anular esta esta Reserva de Repuestos ", vbQuestion + vbYesNo, "Confirma Anulación") = vbYes Then
''
''        mstrSql = "Select Id_Regularizacion as Numero from Stck_Regularizacion where id_ot='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(mstrSql, AdoAnular, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''            With AdoAnular
''                If Not .BOF And Not .EOF Then
''                    .MoveFirst
''                    While Not .EOF
''                        NroRegularizacion = !NUMERO
''                        Call Actualiza_Saldos_VS_Detalle("S", "Select Canrtidad, Id_Empresa, Id_sucursal, Id_Bodega,Id_Ubicacion,Id_Item From Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & NroRegularizacion & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Empresa = '" & gstrIdEmpresa & "'")
''
''                        EliminaReservaRepuestos NroRegularizacion, lblNroRecepcion
''
''                        .MoveNext
''                    Wend
''                    '/// Actualiza estado de reserva
''                    mstrSql = "UPDATE TLLR_OT SET Estado_Reserva='N' "
''                    mstrSql = mstrSql & "Where Id_OT='" & frmRecepcion.lblNroRecepcion & "' "
''                    mstrSql = mstrSql & "And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Seccion_OT='" & gstrSeccion & "'"
''                    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''                    DesactivaBotonAnularReserva
''                End If
''            End With
''        End If
''    Else
''        Exit Sub
''    End If
''End If
''
''End Sub
''Sub DesactivaBotonAnularReserva()
''    cmdAnularReserva.Enabled = False
''    cmdReserva.Enabled = True
''End Sub
''
''Private Sub cmdConsultaSaldo_Click()
''If Me.lvwRepuestosMantencion.ListItems.Count > 0 Then
''    'Levanta listview con los repuestos de la mantencion
''    gstrProcedencia = "Consulta"  'para que solo consulte y no reserve
''    frmRepuestosReservados.Show vbModal
''    gstrProcedencia = "Movimientos"  'vuelve al estado original
''End If
''End Sub
''
''Private Sub cmdConsultaStock_Click()
''    If Me.lvwRepuestos.ListItems.Count > 0 Then
''        'Levanta listview con los repuestos del presupuesto
''        frmRepuestosReservados.Show vbModal
''        ActualizarSaldoRepuestos lblNroRecepcion, gstrSeccion
''    End If
''End Sub
''
''Private Sub cmdReserva_Click()
''If Me.lvwRepuestosMantencion.ListItems.Count > 0 Then
''    'Levanta listview con los repuestos de la mantencion
''    If MsgBox(" Las Cantidades ya estan Confirmadas ? ", vbQuestion + vbYesNo, "Verifica Cantidades") = vbYes Then
''        GrabaReservaRepuestosRecepcion
''        frmRepuestosReservados.Show vbModal
''    Else
''        Exit Sub
''    End If
''End If
''End Sub
''
''Private Sub cmdTemparios_Click()
''frmTemparios.Show
''End Sub
''
''Private Sub dtcConceptos_Change()
''txtSeccion = TipoConcepto(dtcConceptos.BoundText)
''End Sub
''Private Sub dtcGarantia_Change()
''mstrCargo = TraeCargo(dtcGarantia.BoundText)
''TipoOt dtcGarantia.BoundText
''gstrIdCargo = mstrCargo
''End Sub
''Private Sub dtcPartePieza_Change()
''txtHorasCar = TraeHorasDefinidas(lblCompañia.Tag, dtcConceptos.BoundText, dtcPartePieza.BoundText)
''txtValorDefCar = TraeValorDefinido(lblCompañia.Tag, dtcConceptos.BoundText, dtcPartePieza.BoundText)
''txtValorFinCar = TraeValorDefinido(lblCompañia.Tag, dtcConceptos.BoundText, dtcPartePieza.BoundText)
''End Sub
''
''Private Sub dtcTipoCono_Click(Area As Integer)
''If Area > 0 Then
''    txtNroCono.SetFocus
''End If
''End Sub
''
''Private Sub Form_Load()
''    mblnSW = True
''    gstrSeccion = "M"
''    stbServicios.tab = 0
''    gstrKmsAutoNuevo = ""
''    mstrLiquidaPresupuesto = False
''
''    'gcurInsumoDef = gcurInsumo
''End Sub
''
''Private Sub Form_Resize()
''''Dim ldblAncho As Double
''''Dim ldblAnchoCol As Double
''''Dim ldblAnchoBtnSmall As Double
''''
''''Screen.MousePointer = vbHourglass
'''''kjcv 20-01-12
''''ldblAncho = 120
''''ldblAnchoBtnSmall = 240
'''''
''''Me.Frame8.Left = ldblAncho
''''Me.Frame8.Width = Me.Frame8.Width
''''
''''Me.stbServicios.Left = ldblAncho
''''Me.stbServicios.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0)
''''
''''Me.fmePat.Left = ldblAncho
''''Me.fmePat.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''''
''''Me.fmeCia.Left = ldblAncho
''''Me.fmeCia.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''''
''''Me.fmeInv.Width = IIf(Me.ScaleWidth / 2 - (ldblAncho * 2) >= 0, Me.ScaleWidth / 2 - (ldblAncho * 2), 0) - ldblAncho
''''Me.lvwInventario.Width = Me.lvwInventario.Width
''''
''''Me.txtComentario.Width = IIf(Me.ScaleWidth / 2 - (ldblAncho * 2) >= 0, Me.ScaleWidth / 2 - (ldblAncho * 2), 0) - 4 * ldblAncho
''''
''''Me.fmeCom.Left = Me.fmeInv.Left + Me.fmeInv.Width + ldblAncho
''''Me.fmeCom.Width = IIf(Me.ScaleWidth / 2 - (ldblAncho * 2) >= 0, Me.ScaleWidth / 2 - (ldblAncho * 2), 0)
''''
''''Me.fmeMec.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''''
''''Me.lvwServiciosMecanica.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''''
''''Me.lvwRepuestosMantencion.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''''
''''Me.fmeCar.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''''Me.lvwServiciosCarroceria.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''''
''''Me.fmeOtr.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''''Me.lvwOtrosServicios.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''''
''''Me.fmeTer.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''''Me.lvwServiciosTerceros.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''''
''''Me.fmeRep.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''''
''''Me.lvwRepuestos.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''
''End Sub
''
''
''Private Sub lblIdCliente_Change()
''If DatosCliente(lblIdCliente) Then DoEvents
''End Sub
''
''Private Sub lblNroRecepcion_DblClick()
''If gstrImpresion = "O" And Me.lblNroRecepcion <> "" Then
''    gstrBusca = InputBox("Ingrese El Numero de O/T Deseado :", "Ir a....", CStr(Val(Mid(lblNroRecepcion, 6, Len(lblNroRecepcion) - 5))))
''    gstrBusca = FormatOT(gstrBusca)
''    If gstrBusca <> "" Then
'''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
'''kjcv 02.01.13
''mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT like  '%" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
''        gstrSql = letSql(mstrWhere, mstrOrderBy)
''        If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''                LeerCampos
''            End If
''        End If
''        Conexion.CloseHost adoPrincipal
''    End If
''    Screen.MousePointer = vbDefault
''    Me.SetFocus
''End If
''End Sub
''
''Private Sub lblVin_Change()
''VerificaCampañas
''End Sub
''
''Private Sub lvwOtrosServicios_DblClick()
''If mblnBloqueo = False And Me.lvwOtrosServicios.ListItems.Count > 0 Then
''    With lvwOtrosServicios
''        If .SelectedItem.SubItems(11) <> "S" Then
''            If Not .SelectedItem Is Nothing Then
''                frmEditaOtroServicio.Show vbModal
''                AsignaTotal mcFichaOtros, stbTotalOtros
''                TotalFinal
''            End If
''        Else
''            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
''        End If
''    End With
''End If
''End Sub
''Private Sub lvwOtrosServicios_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''If mblnBloqueo = False Then
''    If Me.lvwOtrosServicios.ListItems.Count > 0 Then
''        Select Case Button
''            Case vbRightButton  '//BOTON DERECHO
''                gstrProcedenciaBotonDerecho = "Otros"
''                frmMain.popup(5).Enabled = True
''                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
''        End Select
''    End If
''End If
''
''
'''    Dim i As Integer
'''    Dim gstrBusca As String
'''
'''    Select Case Button
'''        Case vbRightButton  '//BOTON DERECHO
'''            gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
'''            If IsNumeric(gstrBusca) Then
'''                If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
'''                    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
'''                        If Me.lvwOtrosServicios.ListItems(i).Selected Then
'''                            dblTotalInicial = Round(CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(2)) * CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(3)), 2)
'''                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
'''                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
'''                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
'''                        End If
'''
'''                    Next
'''                    AsignaTotal mcFichaOtros, stbTotalOtros
'''                    TotalFinal
'''                Else
'''                    MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
'''                End If
'''            Else
'''                MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
'''            End If
'''    End Select
''
''End Sub
''
''Private Sub lvwRepuestos_DblClick()
''If mblnBloqueo = False And Me.lvwRepuestos.ListItems.Count > 0 Then
''    With lvwRepuestos
''        If .SelectedItem.SubItems(10) <> "S" Then
''            If Not .SelectedItem Is Nothing Then
''                frmEditaServicioRepuesto.Show vbModal
''                gitmActual = .SelectedItem.Index
''                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''                TotalFinal
''                Set .SelectedItem = .ListItems(gitmActual)
''            End If
''        Else
''            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
''        End If
''    End With
''End If
''End Sub
''Private Sub lvwRepuestos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''
''If mblnBloqueo = False Then
''    If Me.lvwRepuestos.ListItems.Count > 0 Then
''        Select Case Button
''            Case vbRightButton  '//BOTON DERECHO
''                gstrProcedenciaBotonDerecho = "Repuestos"
''                frmMain.popup(5).Enabled = False
''                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
''        End Select
''    End If
''End If
''End Sub
''
''Private Sub lvwRepuestosMantencion_DblClick()
''If lvwRepuestosMantencion.ListItems.Count > 0 And Me.cmdReserva.Enabled = True Then
''strMode = "Edit"
''Set lsiItem = lvwRepuestosMantencion.SelectedItem
''With frmEditaTempRepuesto
''    .Caption = "Editar Repuesto"
''    .txtMarca = frmRecepcion.lblMarca
''    .txtModelo = frmRecepcion.lblModelo
''    '.txtServicio = frmRecepcion.lvwServiciosMecanica.ListItems(1).SubItems(1)
''    .txtCodigo = lsiItem
''    .txtDescripcion = lsiItem.SubItems(1)
''    .txtValor = SacarFormatoValor(lsiItem.SubItems(3), "")
''    .txtCantidad = SacarFormatoValor(lsiItem.SubItems(2), "")
''    .Show 1
''End With
''End If
''
''End Sub
''
''Private Sub lvwServiciosCarroceria_DblClick()
''If mblnBloqueo = False And Me.lvwServiciosCarroceria.ListItems.Count > 0 Then
''    With lvwServiciosCarroceria
''        If .SelectedItem.SubItems(17) <> "S" Then
''            If Not .SelectedItem Is Nothing Then
''                gitmActual = .SelectedItem.Index
''                frmEditaTrabajoCarroceria.Show vbModal
''                AsignaTotal mcFichaCarroceria, stbTotalCarroceria
''                TotalFinal
''                Set .SelectedItem = .ListItems(gitmActual)
''            End If
''        Else
''            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
''        End If
''    End With
''End If
''End Sub
''
''Private Sub lvwServiciosCarroceria_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''If mblnBloqueo = False Then
''    If Me.lvwServiciosCarroceria.ListItems.Count > 0 Then
''        Select Case Button
''            Case vbRightButton  '//BOTON DERECHO
''                gstrProcedenciaBotonDerecho = "Carroceria"
''                frmMain.popup(5).Enabled = False
''                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
''        End Select
''    End If
''End If
''
''End Sub
''
''Private Sub lvwServiciosMecanica_DblClick()
''If mblnBloqueo = False And Me.lvwServiciosMecanica.ListItems.Count > 0 Then
''    With lvwServiciosMecanica
''        If .SelectedItem.SubItems(11) <> "S" Then
''            If Not .SelectedItem Is Nothing Then
''                gitmActual = .SelectedItem.Index
''                frmEditaServicioMecanica.Show vbModal
''                AsignaTotal mcFichaMecanica, stbTotalMec
''                TotalFinal
''                Set .SelectedItem = .ListItems(gitmActual)
''            End If
''        Else
''            MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
''        End If
''    End With
''End If
''End Sub
''Private Sub lvwServiciosMecanica_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''If mblnBloqueo = False Then
''    If Me.lvwServiciosMecanica.ListItems.Count > 0 Then
''        Select Case Button
''            Case vbRightButton  '//BOTON DERECHO
''                gstrProcedenciaBotonDerecho = "Mecanica"
''                frmMain.popup(5).Enabled = True
''                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
''        End Select
''    End If
''End If
''End Sub
''
''Private Sub lvwServiciosTerceros_DblClick()
''If mblnBloqueo = False And Me.lvwServiciosTerceros.ListItems.Count > 0 Then
''With lvwServiciosTerceros
''    If .SelectedItem.SubItems(15) <> "S" Then
''        If Not .SelectedItem Is Nothing Then
''            frmEditaServicioTercero.Show 1
''            gitmActual = .SelectedItem.Index
''            AsignaTotal mcFichaTerceros, stbTotalTerceros
''            TotalFinal
''            Set .SelectedItem = .ListItems(gitmActual)
''        End If
''    Else
''        MsgBox "Este Cargo ya fue FACTURADO", vbInformation, "Modificación de Item"
''    End If
''End With
''End If
''
''
''End Sub
''
''Private Sub lvwServiciosTerceros_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''If mblnBloqueo = False Then
''    If Me.lvwServiciosTerceros.ListItems.Count > 0 Then
''        Select Case Button
''            Case vbRightButton  '//BOTON DERECHO
''                gstrProcedenciaBotonDerecho = "Terceros"
''                frmMain.popup(5).Enabled = False
''                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
''        End Select
''    End If
''End If
''End Sub
''
''Private Sub optRecepcion_Click(Index As Integer)
''Select Case Index
''Case 0
''    stbServicios.tab = 0
''    gstrSeccion = "M"
''    If Me.Tag = "" Then
''        Renovar
''    End If
''    'stbServicios.TabEnabled(3) = False
''    Screen.MousePointer = vbDefault
''Case 1
''    stbServicios.tab = 0
''    gstrSeccion = "C"
''    If Me.Tag = "" Then
''        Renovar
''    End If
''    'stbServicios.TabEnabled(3) = True
''    Screen.MousePointer = vbDefault
''End Select
''End Sub
''
''
''
''Private Sub tlbAddRep_ButtonClick(ByVal Button As MSComctlLib.Button)
''Select Case Button.Key
''Case "Agregar" ' ////////////////AGREGAR
''        If Trim(txtPatente.Text) <> "" Then
''            mstrProcedenciaAux = gstrProcedencia
''            gstrProcedencia = "Presupuestos"
''            frmSelTempRepuestos.Show vbModal
''            AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''            TotalFinal
''            gstrProcedencia = mstrProcedenciaAux
''        End If
''    Case "Quitar" ' ////////////////QUITAR
''        If Me.lvwRepuestos.ListItems.Count > 0 Then
''            If Me.lvwRepuestos.SelectedItem.SubItems(11) = "PRESUPUESTO" Then
''                If Not lvwRepuestos.SelectedItem Is Nothing Then
''                    If AccesoEliminar(lvwRepuestos.SelectedItem) = True Then
''                        lvwRepuestos.ListItems.Remove (lvwRepuestos.SelectedItem.Index)
''                        AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''                        TotalFinal
''                    Else
''                        MsgBox ""
''                    End If
''                End If
''            End If
''        End If
''    End Select
''End Sub
''
''Private Sub tlbAddServicioCar_ButtonClick(ByVal Button As MSComctlLib.Button)
''Select Case Button.Key
''Case Is = "Agregar"
''    If Trim(txtPatente) <> "" Then
''        frmAddTrabajosCarroceria.Show vbModal
''        AsignaTotal mcFichaCarroceria, stbTotalCarroceria
''        TotalFinal
''    Else
''        MsgBox LoadResString(301), vbOKOnly, LoadResString(4)
''    End If
''Case Is = "Quitar"
''    'If MsgBox(LoadResString(801), vbYesNo, LoadResString(4)) = 6 Then
''        Call ServicioCarroceria(mDelItem)
''        AsignaTotal mcFichaCarroceria, stbTotalCarroceria
''        TotalFinal
''    'End If
''Case Else
''    DoEvents
''End Select
''End Sub
''Private Sub tlbAddServicioMec_ButtonClick(ByVal Button As MSComctlLib.Button)
''Dim i As Integer
''Dim j As Integer
''Dim lstrServicioMecanica As String
''
''Select Case Button.Key
''Case Is = "Agregar"
''    If Trim(txtPatente) <> "" Then
''        mstrProcedenciaAux = gstrProcedencia
''        gstrProcedencia = "Movimientos"
''        frmAddServiciosMarMod.Show 1
''        lvwRepuestosMantencion.ListItems.Clear
''        mstrAgregaPresupuesto = True
''        For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
''            Repuestos_de_la_Mantencion Me.lblIdMarca, Me.lblIdModelo, lvwServiciosMecanica.ListItems(i), IIf(Me.lvwServiciosMecanica.ListItems(i).SubItems(12) = "S", True, False)
''        Next
''        AsignaTotal mcFichaMecanica, stbTotalMec
''        TotalFinal
''        If lvwServiciosMecanica.ListItems.Count > 0 Then
''            lvwServiciosMecanica.ListItems(lvwServiciosMecanica.ListItems.Count).SubItems(12) = "N"
''        End If
''        gstrProcedencia = mstrProcedenciaAux
''    Else
''        MsgBox LoadResString(301), vbOKOnly, LoadResString(4)
''    End If
''Case Is = "Quitar"
''    If (lvwServiciosMecanica.ListItems.Count > 0 And Me.cmdReserva.Enabled = True) Or Me.dtcGarantia.BoundText = "PRE" Then
''        lstrServicioMecanica = lvwServiciosMecanica.SelectedItem
''        'kjcv 11.09.12 Cambio de subitems 12 a subitems(11), no se podia quitar Servicio de Mecanica
''        If Me.lvwServiciosMecanica.SelectedItem.SubItems(11) = "N" Then
''       ' If MsgBox(LoadResString(801), vbYesNo, LoadResString(4)) = 6 Then
''            If Not lvwServiciosMecanica.SelectedItem Is Nothing Then
''                If Me.dtcGarantia.BoundText = "PRE" Then
''                    '//// quita los repuestos que se agregaron a la ficha de repuestos
''                    Quita_Repuestos_Mantencion Me.lblIdMarca, Me.lblIdModelo, lstrServicioMecanica
''                End If
''                lvwServiciosMecanica.ListItems.Remove (lvwServiciosMecanica.SelectedItem.Index)
''
''                lvwRepuestosMantencion.ListItems.Clear
''                For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
''                    Repuestos_de_la_Mantencion Me.lblIdMarca, Me.lblIdModelo, lvwServiciosMecanica.ListItems(i), IIf(Me.lvwServiciosMecanica.ListItems(i).SubItems(12) = "S", True, False)
''                Next
''                AsignaTotal mcFichaMecanica, stbTotalMec
''                TotalFinal
''            Else
''                MsgBox LoadResString(802), vbOKOnly, LoadResString(4)
''            End If
''        End If
''    Else
''        MsgBox "Si Tiene Una reserva de Repuestos no puede quitar el Servicio", vbExclamation, "Quitar Servicio de Mecanica"
''    End If
''Case Else
''    DoEvents
''End Select
''End Sub
''
''Private Sub tlbAddServicioOtr_ButtonClick(ByVal Button As MSComctlLib.Button)
''Select Case Button.Key
''Case "Agregar"
''    If Trim(txtPatente.Text) <> "" Then
''        frmAddOtrosServicios.Show vbModal
''        AsignaTotal mcFichaOtros, stbTotalOtros
''        TotalFinal
''    End If
''Case "Quitar"
''    If lvwOtrosServicios.ListItems.Count > 0 Then
''        If Not lvwOtrosServicios.SelectedItem Is Nothing Then
''            If Me.lvwOtrosServicios.SelectedItem.SubItems(11) = "N" Then
''                lvwOtrosServicios.ListItems.Remove lvwOtrosServicios.SelectedItem.Index
''                AsignaTotal mcFichaOtros, stbTotalOtros
''                TotalFinal
''            End If
''        End If
''    End If
''End Select
''End Sub
''
''Private Sub tlbAddServicioTer_ButtonClick(ByVal Button As MSComctlLib.Button)
''
''If Not Atributos("Glbl", "Tllr_20_0180", False, False, False, False) Then
''        MsgBox "Ud. No cuenta con Acceso para realizar esta operación...", vbInformation, "Advertencia"
''        Exit Sub
''End If
''
''Select Case Button.Key
''Case "Agregar" ' ////////////////AGREGAR
''    If Trim(txtPatente.Text) <> "" Then
''        frmAddTrabajosTercero.Show vbModal
''        AsignaTotal mcFichaTerceros, stbTotalTerceros
''        TotalFinal
''    End If
''Case "Quitar" ' ////////////////QUITAR
''    If Not lvwServiciosTerceros.SelectedItem Is Nothing Then
''        If Me.lvwServiciosTerceros.SelectedItem.SubItems(15) = "N" Then
''            If Mid(Me.lvwServiciosTerceros.SelectedItem, 1, 2) = "OC" Then
''                MsgBox "No puede Eliminar este Item, porque fue registrado desde una Orden De Compra", vbInformation, "Advertencia"
''            Else
''                lvwServiciosTerceros.ListItems.Remove (lvwServiciosTerceros.SelectedItem.Index)
''                AsignaTotal mcFichaTerceros, stbTotalTerceros
''                TotalFinal
''            End If
''        End If
''    End If
''End Select
''
''End Sub
''
''Private Sub tlbAgregarRepuestos_ButtonClick(ByVal Button As MSComctlLib.Button)
''Select Case Button.Key
''Case "Agregar" ' ////////////////AGREGAR
''        If Trim(txtPatente.Text) <> "" Then
''            gstrProcedencia = "Movimientos"
''            gstrProcedenciaRptos = "Mantencion"
''            frmSelTempRepuestos.Show vbModal
''            gstrProcedenciaRptos = ""
''            'AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''            'TotalFinal
''        End If
''    Case "Quitar" ' ////////////////QUITAR
''        If Me.cmdReserva.Enabled = True Then
''            If Not Me.lvwRepuestosMantencion.SelectedItem Is Nothing Then
''                If AccesoEliminar(Me.lvwRepuestosMantencion.SelectedItem) = True Then
''                    Me.lvwRepuestosMantencion.ListItems.Remove (Me.lvwRepuestosMantencion.SelectedItem.Index)
''                    'AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''                    'TotalFinal
''                Else
''                    MsgBox ""
''                End If
''            End If
''        Else
''            MsgBox "Si tiene una Reserva no puede Quitar Repuestos", vbExclamation, "Reserva de Repuestos"
''        End If
''    End Select
''
''End Sub
''
''Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
''    Screen.MousePointer = vbHourglass
''    Select Case Button.Key
''        Case "Crear"
''            AgregarRegistro
''        Case "Grabar"
''            GrabarRegistro
''        Case "Cancelar"
''            CancelarAgregaRegistro
''        Case "Borrar"
''            BorrarRegistro
''        Case "Buscar"
''            BuscarRegistro
''        Case "Imprimir"
''            PrintOT
''        Case "Primero"
''            PrimerRegistro
''        Case "Anterior"
''            RegistroAnterior
''        Case "Siguiente"
''            RegistroSiguiente
''        Case "Ultimo"
''            UltimoRegistro
''        Case "Activar"
''            EstadosOT gOTActivar
''        Case "Anular"
''            EstadosOT gOTAnular
''        Case "Liquidar"
''            EstadosOT gOTLiquidar
''        Case "Renovar"
''            Renovar
''        Case "Cerrar"
''            CerrarSalir
''        Case "Confirmar"
''            ConfirmarReserva
''        Case "Vaciar"
''            CancelaReserva
''        Case "LiquidarPres"
''            LiquidarPresupuesto
''        Case "AnularPres"
''            AnularPresupuesto
''        Case "Editar"
''            frmHistoricoOT.Show
''        Case "ValoresCargo"
''            If Me.lblEstadoOTValor.Caption <> "VIGENTE" Then
''                frmValoresPorCargo.Show vbModal
''            Else
''                MsgBox "La OT aún está Vigente"
''            End If
''    End Select
''    Screen.MousePointer = vbDefault
''End Sub
''Private Sub Form_Activate()
''    If mblnSW Then
''        mstrProcedencia = gstrProcedencia
''        mblnSW = False
''        If mstrProcedencia = "Movimientos" Then
''            If Not Atributos("Glbl", "Tllr_20_0020", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
''                MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
''                Unload Me
''                Exit Sub
''            End If '/////////ojo
''        ElseIf mstrProcedencia = "Recepcion" Then
''            If Not Atributos("Glbl", "Tllr_20_0010", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
''                MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
''                Unload Me
''                Exit Sub
''            End If '/////////ojo
''        Else
''            If Not Atributos("Glbl", "Tllr_20_0030", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
''                MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
''                Unload Me
''                Exit Sub
''            End If '/////////ojo
''        End If
''
''        tlbAgregarRepuestos.Visible = True
''
''        FillConceptosInventario
''        FillGarantia dtcGarantia, datGarantia, IIf(gstrProcedencia = "Presupuestos", True, False)
''        FillRecepcionista dtcRecepcionista, datRecepcionista
''
''
''        FillTipoCono dtcTipoCono, datTipoCono
''        FillTime gintHoraInicio, gintHoratermino, cboHora
''
''        'FillTipoCargo dtcCargoCar, datCargoCar
''        'FillMecanicos dtcMecanicoCar, datMecanico
''        'FillPartePieza dtcPartePieza, datPartesPiezas
''
''        '//Crear registro por defecto...
''        If gapAccion = apcrear Then
''           AgregarRegistro
''           lblNroRecepcion = gstrBusca
''           Screen.MousePointer = vbDefault
''           Exit Sub
''        End If
''        '//Editar registro por defecto...
''        If gapAccion = apeditar Then
''            If gstrBusca <> "" Then
''                mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT='" & gstrBusca & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''                mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
''                gstrSql = letSql(mstrWhere, mstrOrderBy)
''                If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''                        LeerCampos
''                        ActivaBotones
''                    End If
''                End If
''                Conexion.CloseHost adoPrincipal
''            End If
''            Me.SetFocus
''            Screen.MousePointer = vbDefault
''            Exit Sub
''        End If
''
''        If gapAccion = apninguno Then
''           Renovar
''        End If
''
''        optRecepcion(0).Value = True
''    End If
''    gapAccion = apninguno
''    Screen.MousePointer = vbDefault
''    '//AgregarRegistro
''End Sub
''Private Sub Form_KeyPress(KeyAscii As Integer)
''    Select Case KeyAscii
''        Case vbKeyReturn
''            KeyAscii = 0
''            SendKeys "{tab}"
''        Case vbKeyEscape
''            KeyAscii = 0
''            CancelarAgregaRegistro
''        Case 14 And tlbBarraHerramientas.Buttons.item("Crear").Enabled
''            KeyAscii = 0
''            AgregarRegistro
''        Case 7 And tlbBarraHerramientas.Buttons.item("Grabar").Enabled
''            KeyAscii = 0
''            GrabarRegistro
''        Case 4 And tlbBarraHerramientas.Buttons.item("Borrar").Enabled
''            KeyAscii = 0
''            BorrarRegistro
''        Case 2 And tlbBarraHerramientas.Buttons.item("Buscar").Enabled
''            KeyAscii = 0
''            BuscarRegistro
''        Case 9 And tlbBarraHerramientas.Buttons.item("Imprimir").Enabled
''            KeyAscii = 0
''            PrintOT
''        Case 16 And tlbBarraHerramientas.Buttons.item("Primero").Enabled
''            KeyAscii = 0
''            PrimerRegistro
''        Case 1 And tlbBarraHerramientas.Buttons.item("Anterior").Enabled
''            KeyAscii = 0
''            RegistroAnterior
''        Case 19 And tlbBarraHerramientas.Buttons.item("Siguiente").Enabled
''            KeyAscii = 0
''            RegistroSiguiente
''        Case 21 And tlbBarraHerramientas.Buttons.item("Ultimo").Enabled
''            KeyAscii = 0
''            UltimoRegistro
''        Case 18 And tlbBarraHerramientas.Buttons.item("Renovar").Enabled
''            KeyAscii = 0
''            Renovar
''        Case 17 And tlbBarraHerramientas.Buttons.item("Cerrar").Enabled
''            KeyAscii = 0
''            CerrarSalir
''    End Select
''End Sub
''Private Sub AgregarRegistro()
''    Me.Tag = "Crear"
''    Bloqueo "V"
''    ParametrosDefecto gstrIdEmpresa, gstrIdSucursal
''    lblEstadoOTValor = ""
''    txtTipo = ""
''    DesactivaBotones
''    LimpiaCampos
''    ValoresporDefecto
''    'dtcGarantia.BoundText = gstrIdTipoOtDefecto
''    SetCheckOff lvwInventario
''    lvwServiciosMecanica.ListItems.Clear
''    lvwRepuestosMantencion.ListItems.Clear
''    lvwServiciosCarroceria.ListItems.Clear
''    lvwOtrosServicios.ListItems.Clear
''    lvwServiciosTerceros.ListItems.Clear
''    lvwRepuestos.ListItems.Clear
''    LimpiaTotales
''    stbServicios.tab = 0
''    txtPatente.Enabled = True
''    If fmePat.Enabled = True Then
''        txtPatente.SetFocus
''    End If
''    'kjcv 12.11.13 para DesBloquear el Buscar Placa
''    tlbPatente.Buttons(2).Enabled = True
''    '//// que obligatoriamente elija un tipo de OT
''    If InStr(gstrEmpresa, "AUTO SUMMIT") = 1 Then
''        If mstrProcedencia <> "Presupuestos" Then
''            frmElegirTipoOT.Show vbModal
''            dtcGarantia.Enabled = False
''        End If
''    End If
''
''    '////si es nuevo muestra la ot PRESUPUESTO
''    If mstrProcedencia = "Presupuestos" Then
''        dtcGarantia.BoundText = "PRE"
''        dtcGarantia.Enabled = False
''        lblEstadoOTValor = "PRESUPUESTO"
''    Else
''        dtcGarantia.BoundText = gstrIdTipoOtDefecto
''    End If
''    Me.Tag = "Crear"
''    mstrIdPresupuestoOrigen = ""
'''    gcurInsumoDef = gcurInsumo
''End Sub
''Private Sub CancelarAgregaRegistro()
''    Me.Tag = ""
''    ActivaBotones                                                                       'AND Tllr_OT.ID_OT = Tllr_OT.ID_OT >'" & Trim(lblNroRecepcion) & "'
''    If mstrProcedencia = "Presupuestos" Then
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
''    Else
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
''    End If
''    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
''    gstrSql = letSql(mstrWhere, mstrOrderBy)
''    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''            LeerCampos
''        Else
''            mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''            mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
''            gstrSql = letSql(mstrWhere, mstrOrderBy)
''            If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''                    LeerCampos
''                Else
''                    mblnTablaVacia = True
''                    LimpiaCampos
''                End If
''            End If
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End Sub
''Private Sub GrabarRegistro()
''Dim lstrIdTipoCono As String
''
''    If Not Validacion() Then
''        Exit Sub
''    End If
''
''    If Me.Tag = "Crear" Then
''        If Me.dtcGarantia.BoundText <> "PRE" Then  '  And mstrLiquidaPresupuesto = True Then
''            lblNroRecepcion = TraeCorrelativo(gcOrdenTrabajo, gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
''        Else
''            lblNroRecepcion = "P-" & TraeCorrelativoPresupuesto(gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
''            mstrIdPresupuestoOrigen = lblNroRecepcion
''            If Me.dtcTipoCono = "" Then
''                lstrIdTipoCono = Retorna_Valor_General("Select Top 1 Id_Tipo_Cono from Tllr_Tipo_Cono", gcdynamic)
''                dtcTipoCono.BoundText = lstrIdTipoCono
''            Else
''                lstrIdTipoCono = dtcTipoCono.BoundText
''            End If
''        End If
''
''
''
''        gstrBusca = lblNroRecepcion
''        mstrSql = "INSERT INTO Tllr_OT "
''        mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal, "
''        mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''        mstrSql = mstrSql & " Id_Garantia, Folio_Garantia, "
''        mstrSql = mstrSql & " Id_Tipo_Cono, Nro_Cono, "
''        mstrSql = mstrSql & " Patente, RealizadoPor,"
''        mstrSql = mstrSql & " Kilometros_Recepcion, Id_Compañia_seguro,"
''        mstrSql = mstrSql & " Fecha_Proxima_Visita, "                           'Fecha_Liquidacion,"
''        mstrSql = mstrSql & " Estado,Fecha_Emision, "
''        mstrSql = mstrSql & " Entrega_Estimada, Hora_Entrega, "
''        mstrSql = mstrSql & " Nro_Factura_Emitida,Nro_Presupuesto_Origen,"
''        mstrSql = mstrSql & " Nro_Siniestro, Nro_Poliza, Liquidador, "
''        mstrSql = mstrSql & " Comentario, Solicitado_Por,"
''        mstrSql = mstrSql & " Deducible_UF , Deducible_Pesos, "
''        mstrSql = mstrSql & " Total_Mecanica,Total_Carroceria,"
''        mstrSql = mstrSql & " Total_Desabolladura,Total_Pintura,"
''        mstrSql = mstrSql & " Total_Terceros,Total_Repuestos,"
''        mstrSql = mstrSql & " Total_Materiales,Total_Insumos, "
''        mstrSql = mstrSql & " Total_Otros,Total_Ot,"
''        mstrSql = mstrSql & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor,"
'''        mstrSql = mstrSql & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto,OrdenReparacion,NroReferencia,Bencina ) "
''        'kjcv 19.09.13 se incluyo usuario y fecha de quien genera OT
'''        mstrSQL = mstrSQL & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto,OrdenReparacion,NroReferencia,Bencina,Usr_Id,Usr_Fecha ) "
''        'kjcv 24.10.13 se incluye campo de PDI
''        mstrSql = mstrSql & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto,OrdenReparacion,NroReferencia,Bencina,Usr_Id,Usr_Fecha,PDI ) "
''        mstrSql = mstrSql & " VALUES ("
''        mstrSql = mstrSql & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
''        mstrSql = mstrSql & " '" & lblNroRecepcion & "', '" & gstrSeccion & "',"
''        mstrSql = mstrSql & " '" & Trim(dtcGarantia.BoundText) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), "S/F") & "',"
''        mstrSql = mstrSql & " '" & IIf(dtcGarantia.BoundText = "PRE", lstrIdTipoCono, dtcTipoCono.BoundText) & "', " & CLng(txtNroCono.Text) & ","
''        mstrSql = mstrSql & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
''        mstrSql = mstrSql & " " & CLng(txtKilAct) & ", '" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"   'OJO
''        mstrSql = mstrSql & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
''        mstrSql = mstrSql & " '" & IIf(Me.dtcGarantia.BoundText = "PRE", "P", "V") & "','" & CDate(pckFechaAtencion.Value) & "', "
''        mstrSql = mstrSql & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
''        mstrSql = mstrSql & " '" & "S/N" & "', '" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "',"
''        mstrSql = mstrSql & " '" & IIf(txtNroSiniestro <> " ", UCase(Trim(txtNroSiniestro)), "S/N") & " ','" & IIf(txtNroPoliza <> " ", UCase(Trim(txtNroPoliza)), "S/N") & "','" & IIf(txtLiquidador <> " ", UCase(Trim(txtLiquidador)), "S/L") & "' , "
''        mstrSql = mstrSql & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), "S/S") & "' ,"
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & " ," & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", " & IIf(Me.dtcGarantia.BoundText = "PRE", 0, gcurInsumo) & ", "
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & ", " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & " ,"
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & " ," & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ","
''        mstrSql = mstrSql & " '" & lblIdCliente & "',"
''        mstrSql = mstrSql & " '" & IIf(optMantencion.Value = True, "M", "R") & "',"
''        mstrSql = mstrSql & " '" & IIf(cmdReserva.Enabled = False, "R", "N") & "',"
''        mstrSql = mstrSql & " '" & mstrIdPresupuestoOrigen & "',"
'''        mstrSql = mstrSql & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & ")"
''        'kjcv 19.09.13 se agrego usuario y fecha de generacion de OT
'''        mstrSQL = mstrSQL & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & ",'" & gstrIdUsuario & "','" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "')"
''         mstrSql = mstrSql & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & ",'" & gstrIdUsuario & "','" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "', '" & IIf(Len(txtPatente) > 16, "S", "N") & "' )"
''
''
''        mstrIdPresupuestoOrigen = ""
''    Else
''        mstrSql = "UPDATE Tllr_OT "
''        mstrSql = mstrSql & " SET Id_Garantia='" & Trim(dtcGarantia.BoundText) & "', "
''        mstrSql = mstrSql & " Folio_Garantia='" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), ".") & "', "
''        mstrSql = mstrSql & " Id_Tipo_Cono='" & dtcTipoCono.BoundText & "', "
''        mstrSql = mstrSql & " Nro_Cono=" & CLng(txtNroCono.Text) & ", "
''        mstrSql = mstrSql & " Patente='" & txtPatente.Text & "', "
''        mstrSql = mstrSql & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
''        mstrSql = mstrSql & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
''        mstrSql = mstrSql & " Entrega_Estimada='" & CDate(pckFechaEntrega) & "', "
''        mstrSql = mstrSql & " Hora_Entrega='" & cboHora.Text & "', "
''        mstrSql = mstrSql & " Nro_Siniestro='" & IIf(txtNroSiniestro <> " ", UCase(Trim(txtNroSiniestro)), "S/N") & " ', "
''        mstrSql = mstrSql & " Nro_Poliza='" & IIf(txtNroPoliza <> " ", UCase(Trim(txtNroPoliza)), "S/N") & "', "
''        mstrSql = mstrSql & " Liquidador='" & IIf(txtLiquidador <> " ", UCase(Trim(txtLiquidador)), "S/L") & "', "
''        mstrSql = mstrSql & " Comentario='" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), ".") & "', "
''        mstrSql = mstrSql & " Solicitado_Por='" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), ".") & "',"
''        mstrSql = mstrSql & " Total_Mecanica=" & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Carroceria=" & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " Total_Desabolladura=" & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Pintura=" & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " Total_Terceros=" & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Repuestos=" & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " Total_Otros=" & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & "  ,"
''        mstrSql = mstrSql & " Total_Materiales=" & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Insumos=" & IIf(Me.dtcGarantia.BoundText = "PRE", 0, gcurInsumo) & ", "
'''        mstrSQL = mstrSQL & " Total_Ot=" & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo & "  ,"
''        'kjcv 14.10.15 se quito el valor de insumos
''        mstrSql = mstrSql & " Total_Ot=" & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & "  ,"
''        mstrSql = mstrSql & " Total_OT_Iva=" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & "  ,"
''        mstrSql = mstrSql & " Total_IVA =" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & "  ,"
''        mstrSql = mstrSql & " Deducible_UF = " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , "
''        mstrSql = mstrSql & " Deducible_Pesos = " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
''        mstrSql = mstrSql & " Nro_Presupuesto_Origen='" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "', "
''        mstrSql = mstrSql & " Kilometros_Recepcion=" & CLng(txtKilAct) & ","
''        mstrSql = mstrSql & " Id_Compañia_Seguro='" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"
''        mstrSql = mstrSql & " Fecha_Proxima_Visita = '" & DateAdd("d", 365, pckFechaAtencion.Value) & "',"
''        mstrSql = mstrSql & " Id_Cliente_Proveedor='" & lblIdCliente & "',"
''        mstrSql = mstrSql & " ReparacionMantencion='" & IIf(Me.optMantencion.Value = True, "M", "R") & "',"
''        mstrSql = mstrSql & " Estado_Reserva='" & IIf(Me.cmdReserva.Enabled = False, "R", "N") & "',"
''        mstrSql = mstrSql & " OrdenReparacion='" & txtOrdenReparacion & "',"
''        mstrSql = mstrSql & " NroReferencia='" & txtNReferencia & "',"
''        'kjcv 19.09.19
''        mstrSql = mstrSql & " Usr_Id='" & gstrIdUsuario & "',"
'''        mstrSQL = mstrSQL & " Usr_Fecha='" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "'"
''        'kjcv 24.10.13
''        mstrSql = mstrSql & " Usr_Fecha='" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "',"
''        mstrSql = mstrSql & " PDI='" & IIf(Len(txtPatente) > 16, "S", "N") & "'"
''        'mstrSql = mstrSql & " Id_Presupuesto='" & mstrIdPresupuestoOrigen & "'"
''        mstrSql = mstrSql & " WHERE Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal ='" & gstrIdSucursal & "' And Id_OT ='" & Trim(Trim(lblNroRecepcion)) & "' AND Seccion_OT ='" & gstrSeccion & "' "
''    End If                                                                                                                                                                                                                                                                              ''" & pckFechaVenta.Value & "'
''
''    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''        '/////////////////////////////// AQUI GUARDAR DATOS DEL VEHICULO
''            mstrSql = " Update Tllr_Vehiculo_Cliente "
''            mstrSql = mstrSql & " Set Kilometros_Actuales = " & IIf(Trim(txtKilAct) <> "", CLng(txtKilAct), 0) & " , "
''            mstrSql = mstrSql & " Concesionario='" & IIf(Trim(txtConcesionario) <> "", UCase(Trim(txtConcesionario)), "S/C") & "' ,"
''            mstrSql = mstrSql & " Fecha_Venta='" & pckFecVta.Value & "'"
''            mstrSql = mstrSql & " Where Patente='" & txtPatente & "'"
''        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''            MsgBox LoadResString(323)
''        End If
''
''            If GuardaInventario(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
''                MsgBox LoadResString(322)
''            End If
''
''            If GuardaMecanica(lblNroRecepcion, gcOrdenTrabajo) = False Then
''                MsgBox LoadResString(321)
''            End If
''
''            If GuardaCarroceria(lblNroRecepcion, gstrSeccion, lblCompañia.Tag, gcOrdenTrabajo) = False Then
''                MsgBox LoadResString(320)
''            End If
''
''            If GuardaOtros(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
''                MsgBox LoadResString(328)
''            End If
''
''            If GuardaTerceros(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
''                MsgBox LoadResString(319)
''            End If
''
''
''
''
''        'traspasa los repuestos de un presupuesto a una ot segun parametro
''        If mstrLiquidaPresupuesto = True Then
''            If gblnTraspasaRepuestos = True Then
''                If GuardaRepuestosPresupuesto(lblNroRecepcion, gstrSeccion) = False Then
''                    MsgBox LoadResString(318)
''                End If
''            End If
''        Else
''            If GuardaRepuestos(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
''                MsgBox LoadResString(318)
''            End If
''        End If
'''//////////////////////////////////
''
''        'actualiza datos de rent a car
''        If Me.dtcGarantia.BoundText = "REN" And Me.optMantencion.Value = True Then
''            gstrEstadoMantencion = Retorna_Valor_General("Select EstadoMantencion from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
''            gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoMantencion & "'"
''            gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
''            If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''            End If
''        End If
''        If Me.dtcGarantia.BoundText = "REN" And Me.optReparacion.Value = True Then
''            gstrEstadoReparacion = Retorna_Valor_General("Select EstadoReparacion from Rent_Parametros_Globales where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
''            gstrSql = "UPDATE Auto_Stock SET Id_ESTADO_Vehiculo = '" & gstrEstadoReparacion & "'"
''            gstrSql = gstrSql & " Where Patente = '" & Me.txtPatente & "'"
''            If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''            End If
''        End If
''
'''//////////////////////////////////
''        mblnTablaVacia = False
''        ActivaBotones
''        Me.Tag = ""
'''//////////////////////////////////
''        If lblEstadoOT.Visible = False Then
''            If MsgBox("Imprimirá la OT Nº " & lblNroRecepcion & ", Confirma el Documento", 4 + 32, "Imprime OT(Recepción)") = vbYes Then
''                PrintOT
''                If mstrLiquidaPresupuesto = False Then 'cuando liquida presupuesto no borre la pantalla
''                    AgregarRegistro
''                End If
''            Else
''                If mstrLiquidaPresupuesto = False Then
''                    AgregarRegistro
''                End If
''            End If
''        End If
''    End If '//////////////
''End Sub
''Sub GrabarPresupuesto(NumeroPresupuesto As String, NumeroOT As String, EstadoPresupuesto As String, MotivoAnula As String)
''
''    mstrSql = "INSERT INTO Tllr_Presupuesto "
''    mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal, "
''    mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''    mstrSql = mstrSql & " Id_Garantia, Folio_Garantia, "
''    mstrSql = mstrSql & " Id_Tipo_Cono, Nro_Cono, "
''    mstrSql = mstrSql & " Patente, RealizadoPor,"
''    mstrSql = mstrSql & " Kilometros_Recepcion, Id_Compañia_seguro,"
''    mstrSql = mstrSql & " Fecha_Proxima_Visita, "                           'Fecha_Liquidacion,"
''    mstrSql = mstrSql & " Estado,Fecha_Emision, "
''    mstrSql = mstrSql & " Entrega_Estimada, Hora_Entrega, "
''    mstrSql = mstrSql & " Nro_Factura_Emitida,Nro_Presupuesto_Origen,"
''    mstrSql = mstrSql & " Nro_Siniestro, Nro_Poliza, Liquidador, "
''    mstrSql = mstrSql & " Comentario, Solicitado_Por,"
''    mstrSql = mstrSql & " Deducible_UF , Deducible_Pesos, "
''    mstrSql = mstrSql & " Total_Mecanica,Total_Carroceria,"
''    mstrSql = mstrSql & " Total_Desabolladura,Total_Pintura,"
''    mstrSql = mstrSql & " Total_Terceros,Total_Repuestos,"
''    mstrSql = mstrSql & " Total_Materiales,Total_Insumos, "
''    mstrSql = mstrSql & " Total_Otros,Total_Ot,"
''    mstrSql = mstrSql & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor,"
''    mstrSql = mstrSql & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto, Descripcion_Anula, Fecha_Liquidacion ) "
''    mstrSql = mstrSql & " VALUES ("
''    mstrSql = mstrSql & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
''    mstrSql = mstrSql & " '" & NumeroOT & "', '" & gstrSeccion & "',"
''    mstrSql = mstrSql & " '" & Trim(dtcGarantia.BoundText) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), "S/F") & "',"
''    mstrSql = mstrSql & " '" & dtcTipoCono.BoundText & "', " & CLng(txtNroCono.Text) & ","
''    mstrSql = mstrSql & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
''    mstrSql = mstrSql & " " & CLng(txtKilAct) & ", '" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"   'OJO
''    mstrSql = mstrSql & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
''    mstrSql = mstrSql & " '" & EstadoPresupuesto & "','" & CDate(pckFechaAtencion.Value) & "', "
''    mstrSql = mstrSql & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
''    mstrSql = mstrSql & " '" & "S/N" & "', '" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "',"
''    mstrSql = mstrSql & " '" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ','" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "','" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "' , "
''    mstrSql = mstrSql & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), "S/S") & "' ,"
''    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
''    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & " ," & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
''    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
''    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
''    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", " & gcurInsumo & ", "
''    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & ", " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & " ,"
''    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & " ," & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ","
''    mstrSql = mstrSql & " '" & lblIdCliente & "',"
''    mstrSql = mstrSql & " '" & "M" & "',"
''    mstrSql = mstrSql & " '" & "N" & "',"
''    mstrSql = mstrSql & " '" & NumeroPresupuesto & "',"
''    mstrSql = mstrSql & " '" & MotivoAnula & "',"
''    mstrSql = mstrSql & " '" & Format(Date, "DD/MM/YYYY") & "')"
''
''    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''        If GuardaInventario(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
''            MsgBox LoadResString(322)
''        End If
''        If GuardaMecanica(NumeroPresupuesto, gcPresupuesto) = False Then
''            MsgBox LoadResString(321)
''        End If
''        If GuardaCarroceria(NumeroPresupuesto, gstrSeccion, lblCompañia.Tag, gcPresupuesto) = False Then
''            MsgBox LoadResString(320)
''        End If
''        If GuardaOtros(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
''            MsgBox LoadResString(328)
''        End If
''        If GuardaTerceros(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
''            MsgBox LoadResString(319)
''        End If
''        If GuardaRepuestos(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
''            MsgBox LoadResString(318)
''        End If
''
'''//////////////////////////////////
''        mblnTablaVacia = False
''    End If '//////////////
''
''End Sub
''Private Sub BorrarRegistro()
''    Screen.MousePointer = vbDefault
''    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
''        '////////////////////////////////ELIMINAR SERVICIOS DE MECANICA///////////////////////////////////
''        mstrSql = "DELETE FROM Tllr_Mecanica_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        Conexion.SendHost mstrSql, , , , gcTiempoEspera
''        '////////////////////////////////ELIMINAR SERVICIOS DE CARRPCERIA///////////////////////////////////
''        mstrSql = "DELETE FROM Tllr_Carroceria_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        Conexion.SendHost mstrSql, , , , gcTiempoEspera
''        '////////////////////////////////////ELIMINAR INENTARIO///////////////////////////////
''        mstrSql = "DELETE FROM Tllr_Inventario_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        Conexion.SendHost mstrSql, , , , gcTiempoEspera
''        '//////////////////////////////////////ENCABEZADO/////////////////////////////
''        mstrSql = "DELETE FROM Tllr_OT WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''            mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''            mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
''            gstrSql = letSql(mstrWhere, mstrOrderBy)
''            If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''                    LeerCampos
''                Else
''                    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''                    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
''                    gstrSql = letSql(mstrWhere, mstrOrderBy)
''
''                    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''                        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''                            LeerCampos
''                        Else
''                            mblnTablaVacia = True
''                            LimpiaCampos
''                        End If
''                    End If
''                End If
''            End If
''        End If
''        Conexion.CloseHost adoPrincipal
''    End If
''End Sub
''Private Sub BuscarRegistro()
''Screen.MousePointer = 1
''frmBuscaOT.Show vbModal
''Screen.MousePointer = 1
''If gstrBusca <> "" Then
''    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
''    gstrSql = letSql(mstrWhere, mstrOrderBy)
''    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''            LeerCampos
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End If
''Me.SetFocus
''
''End Sub
''Private Sub PrimerRegistro()
''    If mstrProcedencia = "Presupuestos" Then
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
''    Else
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
''    End If
''    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
''    gstrSql = letSql(mstrWhere, mstrOrderBy)
''    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''            LeerCampos
''        Else
''            Beep
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End Sub
''Private Sub RegistroAnterior()
''    If mstrProcedencia = "Presupuestos" Then
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
''    Else
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
''    End If
''    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
''    gstrSql = letSql(mstrWhere, mstrOrderBy)
''    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''            LeerCampos
''        Else
''            Beep
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End Sub
''Private Sub RegistroSiguiente()
''    If mstrProcedencia = "Presupuestos" Then
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
''    Else
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
''    End If
''    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT "
''    gstrSql = letSql(mstrWhere, mstrOrderBy)
''    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''            LeerCampos
''        Else
''            Beep
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End Sub
''Private Sub UltimoRegistro()
''    If mstrProcedencia = "Presupuestos" Then
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
''    Else
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
''    End If
''    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
''    gstrSql = letSql(mstrWhere, mstrOrderBy)
''    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
''            LeerCampos
''        Else
''            Beep
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End Sub
''Private Sub Renovar()
''
''    If mstrProcedencia = "Presupuestos" Then
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
''    Else
''        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
''    End If
''    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT "
''    gstrSql = letSql(mstrWhere, mstrOrderBy)
''    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''        VerificaTablaVacia
''        ActivaBotones
''        If Not mblnTablaVacia Then
''            PrimerRegistro
''        End If
''    End If
''    Conexion.CloseHost adoPrincipal
''End Sub
''Private Sub CerrarSalir()
''    Unload Me
''End Sub
''Private Sub ActivaBotones()
''    With tlbBarraHerramientas.Buttons
''        .item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
''        .item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(gstrProcedencia = "Recepcion", False, IIf(mblnAccesoEditar, True, False)))
''        .item("Cancelar").Enabled = False
''        .item("Buscar").Enabled = IIf(mblnTablaVacia, False, True)
''        .item("Imprimir").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoImprimir, True, False))
''        .item("Primero").Enabled = IIf(mblnTablaVacia, False, True)
''        .item("Anterior").Enabled = IIf(mblnTablaVacia, False, True)
''        .item("Siguiente").Enabled = IIf(mblnTablaVacia, False, True)
''        .item("Ultimo").Enabled = IIf(mblnTablaVacia, False, True)
''        .item("Renovar").Enabled = True
''        .item("Cerrar").Enabled = True
''    End With
''End Sub
''Private Sub DesactivaBotones()
''With tlbBarraHerramientas.Buttons
''    .item("Crear").Enabled = False
''    .item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
''    .item("Cancelar").Enabled = True
''    .item("Buscar").Enabled = False
''    .item("Imprimir").Enabled = False
''    .item("Primero").Enabled = False
''    .item("Anterior").Enabled = False
''    .item("Siguiente").Enabled = False
''    .item("Ultimo").Enabled = False
''    .item("Renovar").Enabled = False
''    .item("Cerrar").Enabled = True
''    .item("Activar").Enabled = False
''    .item("Liquidar").Enabled = False
''    .item("Anular").Enabled = False
''End With
''End Sub
''Private Sub VerificaTablaVacia()
''    If (Not adoPrincipal.BOF And Not adoPrincipal.EOF) And adoPrincipal.RecordCount > 0 Then
''        mblnTablaVacia = False
''    Else
''        mblnTablaVacia = True
''        LimpiaCampos
''    End If
''End Sub
''
''Private Sub LimpiaCampos()
''With Me
''    SetCheckOff .lvwInventario
''    .lvwServiciosCarroceria.ListItems.Clear
''    .lvwServiciosMecanica.ListItems.Clear
''    .lvwServiciosTerceros.ListItems.Clear
''    .lvwRepuestos.ListItems.Clear
''    .lblNroRecepcion.Text = ""
''    .dtcGarantia.BoundText = ""
''    .dtcGarantia.Enabled = True
''    .pckFechaAtencion.Value = Now
''    .txtPatente.Text = ""
''    .lblMarca.Caption = "": .lblIdMarca = ""
''    .lblModelo.Caption = "": .lblIdModelo = ""
''    .txtAño.Text = ""
''    .lblColorE.Caption = ""
''    .lblChasis.Caption = ""
''    .lblMotor.Caption = ""
''    .lblCliente.Caption = ""
''    .txtKilAct.Text = ""
''    .txtConcesionario.Text = ""
''    .pckFecVta.Value = Now
''    .dtcTipoCono.BoundText = ""
''    .txtNroCono.Text = ""
''    .dtcRecepcionista.BoundText = ""
''    .pckFechaEntrega.Value = Now
''    .cboHora.Text = ""
''    .lblCompañia.Caption = ""
''    .lblCompañia.Tag = ""
''    .txtDeducibleUF.Text = "0"
''    .txtDeduciblePesos.Text = "0"
''    .txtNroSiniestro.Text = ""
''    .txtNroPoliza.Text = ""
''    .txtLiquidador.Text = ""
''    .lblFono.Caption = ""
''    .lblVin.Caption = ""
''    .txtSolicita.Text = ""
''    .txtFolioGarantia.Text = ""
''    .txtRut.Text = ""
''    .txtComuna.Text = ""
''    .txtDir.Text = ""
''    .lblIdCliente.Caption = ""
''    .txtComentario = ""
''    .cmdAnularReserva.Enabled = False
''    .cmdReserva.Enabled = True
''    .lblPresupuesto = ""
''    .lblFechaLiquidacion = ""
''    .txtOrdenReparacion = ""
''    .lblDocumentos = ""
''    .txtNReferencia = ""
''    .txtTipo = ""
''End With
''End Sub
''Private Sub ValoresporDefecto()
''    txtAño.Text = Year(Now)
''    txtDeducibleUF.Text = "0"
''    txtNroCono.Text = "0"
''    txtDeduciblePesos.Text = "0"
''    txtNroSiniestro.Text = " "
''    txtNroPoliza.Text = " "
''    txtLiquidador.Text = " "
''    txtKilAct.Text = "0"
''    lblEstadoOTValor = "VIGENTE"
''    lblEstadoOTValor.Tag = "V"
''End Sub
''Private Function Validacion() As Boolean
''    Validacion = True
''With Me
''    If .dtcGarantia.BoundText = "" Then
''        MsgBox LoadResString(317), vbInformation, "Advertencia"
''        dtcGarantia.Enabled = True
''        dtcGarantia.SetFocus
''        Validacion = False
''        Exit Function
''    End If
''    If .txtPatente = "" Then
''        MsgBox LoadResString(316), vbInformation, "Advertencia"
''        txtPatente.SetFocus
''        Validacion = False
''        Exit Function
''    Else
''        If ExistePatente(txtPatente) = False Then
''            MsgBox LoadResString(329), vbInformation, "Advertencia"
''            txtPatente.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''    End If
''    If gstrSeccion = "C" Then
''        If .txtFolioGarantia = "" Then
''            MsgBox LoadResString(315), vbInformation, "Advertencia"
''            txtFolioGarantia.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''    End If
''    If .txtSolicita = "" Then
''        MsgBox LoadResString(314), vbInformation, "Advertencia"
''        txtSolicita.SetFocus
''        Validacion = False
''        Exit Function
''    End If
'''    If .txtKilAct = "" Then
'''        MsgBox LoadResString(313), vbInformation, "Advertencia"
'''        txtKilAct.SetFocus
'''        Validacion = False
'''        Exit Function
'''    End If
''    If (CDbl(.txtKilAct) = 0 Or .txtKilAct = "") Or (CDbl(.txtKilAct) <= KilometrajeEntrada) Then
''        If UCase(Me.Tag) = "CREAR" And Me.lblEstadoOTValor <> "RESERVA" And gstrKmsAutoNuevo <> "Nuevo" And Me.dtcGarantia.BoundText <> "PRE" And mstrLiquidaPresupuesto = False Then
''            'MsgBox LoadResString(313), vbInformation, "Advertencia"
''            MsgBox "El Kilometraje de la última visita fué de " & CDbl(.txtKilAct) & Chr(13) & "Verifique el kilometraje ingresado...", vbInformation, "Advertencia"
''            Me.Frame3.Enabled = True
''            Me.Frame4.Enabled = True
''            Me.Frame8.Enabled = True
''            txtKilAct.Enabled = True
''            'txtKilAct.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''    End If
''
''    If .dtcTipoCono.BoundText = "" Then
''        If Me.dtcGarantia.BoundText <> "PRE" Then
''            MsgBox LoadResString(312), vbInformation, "Advertencia"
''            dtcTipoCono.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''    End If
''    If .txtNroCono = "" Then
''        MsgBox LoadResString(311), vbInformation, "Advertencia"
''        txtNroCono.SetFocus
''        Validacion = False
''        Exit Function
''    End If
''    If .dtcRecepcionista.BoundText = "" Then
''        MsgBox LoadResString(310), vbInformation, "Advertencia"
''        dtcRecepcionista.SetFocus
''        Validacion = False
''        Exit Function
''    End If
''    If dtcGarantia.BoundText = "REN" Then
''        If Me.optMantencion.Value = False And Me.optReparacion.Value = False Then
''            MsgBox "Para Rent a Car Debe elegir Reparación o Mantención", vbInformation, "Advertencia"
''            dtcGarantia.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''    End If
''    If gstrProcedencia = "Recepcion" Then
''        If Me.cmbBencina.Text = "" Then
''            MsgBox "El estado del Estanque de Gasolina debe contener un valor", vbExclamation, "Recepción"
''            stbServicios.tab = 1
''            Me.cmbBencina.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''    End If
''    If .optRecepcion(1).Value = True Then
''        If .txtDeducibleUF.Text = "" Then
''            MsgBox LoadResString(308), vbInformation, "Advertencia"
''            txtDeducibleUF.SetFocus
''
''            Validacion = False
''            Exit Function
''        End If
''        If .txtDeduciblePesos.Text = "" Then
''            MsgBox LoadResString(307), vbInformation, "Advertencia"
''            txtDeduciblePesos.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''        If .txtNroSiniestro.Text = "" Then
''            MsgBox LoadResString(306), vbInformation, "Advertencia"
''            txtNroSiniestro.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''        If .txtNroPoliza.Text = "" Then
''            MsgBox LoadResString(305), vbInformation, "Advertencia"
''            txtNroPoliza.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''        If .txtLiquidador.Text = "" Then
''            MsgBox LoadResString(304), vbInformation, "Advertencia"
''            txtLiquidador.SetFocus
''            Validacion = False
''            Exit Function
''        End If
''    End If
''    '//////////////////////////////////CARROCERIA
''End With
''    '//Verifica si existe un registro...
''    If Me.Tag = "Crear" And Me.lblEstadoOTValor <> "RESERVA" And Me.lblEstadoOTValor <> "PRESUPUESTO" Then
''        Dim adoTemp As New ADODB.Recordset
''        mstrSql = "select ID_OT from TLLR_OT where SECCION_OT = '" & gstrSeccion & "' AND ID_OT ='" & lblNroRecepcion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''            If Not adoTemp.BOF And Not adoTemp.EOF Then
''                MsgBox "Este código ya esta registrado con la descripción "
''                Validacion = False
''            End If
''        End If
''        Conexion.CloseHost adoTemp
''    End If
''End Function
''Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
''    Set frmRecepcion = Nothing
''    gstrBusca = lblNroRecepcion.Text
''End Sub
''Private Sub RevizaAtributos()
''    mblnAccesoCrear = True
''    mblnAccesoEditar = True
''    mblnAccesoBorrar = True
''    mblnAccesoImprimir = True
''End Sub
''
''Private Sub tlbBusca_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
''Dim lstrNombre As String
''Dim lstrSQL As String
''
''Select Case Button.Key
''    Case "Nuevo"
''        gstrBusca = ""
''        lstrNombre = ""
'''        gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, gstrBusca, lstrNombre, apcrear, "Cliente - Proveedor", gstrIdSucursal)
''
''
''        lblIdCliente = gstrBusca
''        'ACTUALIZA PATENTE V/S CLIENTE
''        lstrSQL = "Update Tllr_Vehiculo_Cliente set Id_Cliente_Proveedor='" & lblIdCliente & "' Where Patente='" & txtPatente & "'"
''        Conexion.SendHost lstrSQL, , , , gcTiempoEspera
''    Case "Buscar"
'''        gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, lblIdCliente, lstrNombre, apeditar, "Cliente - Proveedor", gstrIdSucursal)
'''        lblIdCliente = gstrBusca
'''       Me.lblCliente.Caption = lblIdCliente
'''       DatosCliente (lblIdCliente)
'''kjcv 02-02-2012
''        gstrRutCliente = ""
''        gstrNombreCliente = ""
''        Libreria.ClienteBuscar Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
''         If gstrRutCliente <> "" Then
''            Me.lblCliente.Caption = gstrNombreCliente
''            Me.lblCliente.Tag = gstrRutCliente
''        End If
''    End Select
''End Sub
''
''Private Sub tlbCiaSeg_ButtonClick(ByVal Button As MSComctlLib.Button)
''Select Case Button.Key
''Case Is = "Nueva"
''    gstrProcedencia = "Movimientos"
''    frmMantenedorCompañiaSeguro.Show 1
''
''Case Is = "Buscar"
''    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Compañia_Seguro", "Id_Compañia_Seguro", "Nombre", "Busca Compañia de Seguro")
''    lblCompañia = NombreCiaSeg(gstrBusca)
''    lblCompañia.Tag = gstrBusca
''    FillConceptosVsCiaSeguro dtcConceptos, datConceptos, lblCompañia.Tag
''    txtDeduciblePesos.SetFocus
''End Select
''
''End Sub
''
''Private Sub tlbPatente_ButtonClick(ByVal Button As MSComctlLib.Button)
''
''If Me.Tag = "Crear" Then
''    Select Case Button.Key
''    Case "Nuevo"
''        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apcrear)
''        DatosVehiculo txtPatente
''    Case "Buscar"
''        gstrProcedencia = "Movimientos"
''        frmBuscaVehiculo.Show vbModal
''        'kjcv 30.10.15
''    Case "Historial"
''        frmHistorialPlaca.Show vbModal
''    End Select
''Else
''    Select Case Button.Key
''    Case "Nuevo"
''        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apeditar)
''        DatosVehiculo txtPatente
''    Case "Buscar"
''        gstrProcedencia = "Movimientos"
''        frmBuscaVehiculo.Show vbModal
''        'kjcv 30.10.15
''    Case "Historial"
''        frmHistorialPlaca.Show vbModal
''    End Select
''End If
''
''End Sub
''
''
''Private Sub tlbPatente_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
''If tlbPatente.Buttons(1).Key = "Nuevo" Then
''    tlbPatente.Buttons(1).ToolTipText = IIf(Me.Tag = "Crear", "Nuevo Vehiculo", "Editar Vehiculo")
''Else
''    tlbPatente.Buttons(2).ToolTipText = IIf(Me.Tag = "Crear", "Buscar Vehiculo", "Buscar Vehiculo")
''End If
''End Sub
''
''Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''
''End Sub
''
''Private Sub tlbTemparioCarroceria_ButtonClick(ByVal Button As MSComctlLib.Button)
''Select Case Button.Key
''    Case Is = "Temparios"
''        frmTemparios.Show
''
''End Select
''End Sub
''
''Private Sub txtConcesionario_GotFocus()
''MarcaTexto txtConcesionario
''End Sub
''
''Private Sub txtDeduciblePesos_GotFocus()
''MarcaTexto txtDeduciblePesos
''End Sub
''
''
''Private Sub txtDeducibleUF_GotFocus()
''MarcaTexto txtDeducibleUF
''
''End Sub
''
''
''Private Sub txtFolioGarantia_GotFocus()
''MarcaTexto txtFolioGarantia
''End Sub
''
''Private Sub txtKilAct_GotFocus()
''MarcaTexto txtKilAct
''End Sub
''
''Private Sub txtLiquidador_KeyPress(KeyAscii As Integer)
''KeyAscii = UpCaseLetter(KeyAscii)
''End Sub
''
''Private Sub txtNroCono_GotFocus()
''MarcaTexto txtNroCono
''End Sub
''
''Private Sub txtPatente_GotFocus()
''MarcaTexto txtPatente
''End Sub
''
''Private Sub txtPatente_KeyDown(KeyCode As Integer, Shift As Integer)
''Dim MyRecordset As New ADODB.Recordset
'''If Me.Tag = "Crear" Then
''Dim str1 As String
''Dim str2 As String
''    If KeyCode = 13 Then
''''        kjcv 24 - 01 - 12
''''        CheckPatente txtPatente, str1, str2  '/// devuelve el rut de la patente
'''        txtFolioGarantia = str2
''        If txtPatente <> "" Then
''            'If Len(txtPatente) = 6 And lblPat.Caption = gstrNombrePatente Or Me.dtcGarantia.BoundText = "PEX" Then
''                If dtcGarantia.BoundText = "VHP" Then  '/// valida patente vehiculos propios
''                    If ConsultaVehiculoPropio(txtPatente) = False Then
''                        MsgBox gstrNombrePatente & " no EXISTE en Vehiculos Propios", vbInformation, "Ingreso de " & gstrNombrePatente
''                        Exit Sub
''                    End If
''                End If
''                If dtcGarantia.BoundText = "REN" Then
''                    Set MyRecordset = cnnAux.Execute("EXEC RENT_ACTUALIZA_VEHICULO_CLIENTE '" & Me.txtPatente & "', '" & gstrIdUsuario & "', '" & Date & "'")
''                End If
''
''                If ConsultaVehiculo(txtPatente) = True Then
''                    'kjcv 15.11.13
''                    If ConsultaPatente(txtPatente) = True Then
''                        MsgBox "No hay Cupo en el Taller...", vbCritical, "Elisa"
''                        Call DatosVehiculo(txtPatente)
''                    Else
''                        Call DatosVehiculo(txtPatente)
''                    End If
'''                    Call DatosVehiculo(txtPatente)
''                Else
''                    gstrProcedencia = "Movimientos"
''                    gapAccion = apcrear
''                    gstrKmsAutoNuevo = "Nuevo"
''                    frmMantenedorVehiculoCliente.Show vbModal
''                End If
''
'''            ElseIf dtcGarantia.BoundText = "INW" Or dtcGarantia.BoundText = "INC" Then
'''                If ConsultaVinExistencia(txtPatente) = True Then
'''                    If ConsultaVehiculo(txtPatente) = True Then
'''                        If MsgBox("La " & gstrNombrePatente & " " & txtPatente & " Ya Existe, Desea Desplegar los Datos", 4 + 32, "Patente Existente") = vbYes Then
'''                            Call DatosVehiculo(txtPatente)
'''                        Else
'''                            LimpiaCampos
'''                        End If
'''                    Else
'''                        gstrProcedencia = "Movimientos"
'''                        gapAccion = apcrear
'''                        frmMantenedorVehiculoCliente.Show vbModal
'''                    End If
'''                End If
'''
'''            ElseIf gstrValidaPatente = "N" Then
'''                If ConsultaVehiculo(txtPatente) = True Then
'''                    Call DatosVehiculo(txtPatente)
'''                Else
'''                    gstrProcedencia = "Movimientos"
'''                    gapAccion = apcrear
'''                    gstrKmsAutoNuevo = "Nuevo"
'''                    frmMantenedorVehiculoCliente.Show vbModal
'''                End If
'''            Else
'''                'MsgBox LoadResString(326)
'''            End If
'''
'''        Else
'''            MsgBox LoadResString(327)
''        End If
''
''    End If
'''End If
''End Sub
''
''Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'''If Trim(lblPat.Caption) = gstrNombrePatente And dtcGarantia.BoundText <> "PEX" Then
'''    If gstrValidaPatente = "S" Then
'''        KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'''    End If
'''End If
''
'''kjcv 24-01-12
''If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
''    KeyAscii = 0: Beep
''Else
''    KeyAscii = UpCaseLetter(KeyAscii)
''End If
''End Sub
''
''Private Sub txtSolicita_GotFocus()
''MarcaTexto txtSolicita
''End Sub
''Private Sub ConfirmarReserva()
''    Screen.MousePointer = vbHourglass
''    gstrBuscaReserva = lblNroRecepcion
''    Me.Tag = "Crear"
''    If Validacion() = True Then
''        GrabarRegistro    '/// graba la ot de reserva en una ot definitiva
''
''        '/// Asigna numero de ot a la reserva de hora
''        mstrSql = "Update Tllr_ReservaHora "
''        mstrSql = mstrSql & " Set Id_OT = '" & gstrBusca & "',"
''        mstrSql = mstrSql & " Estado='R'"
''        mstrSql = mstrSql & " Where Id_Reserva='" & Mid(gstrBuscaReserva, 3, 5) & "'"
''        mstrSql = mstrSql & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''            MsgBox "Error Al Actualizar Los Datos De La Reserva de Hora"
''        End If
''
''        '/// actualiza la reserva de repuestos
''        mstrSql = "Update Stck_Regularizacion "
''        mstrSql = mstrSql & " Set Id_OT = '" & gstrSeccion & gstrBusca & "'"
''        mstrSql = mstrSql & " Where Id_Ot='" & gstrSeccion & gstrBuscaReserva & "'"
''        mstrSql = mstrSql & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''            MsgBox "Error Al Actualizar Los Datos De La Reserva de Repuestos"
''        End If
''
''        '/// actualiza repuestos reservados
''        mstrSql = "Update Tllr_Repuestos_Reservados "
''        mstrSql = mstrSql & " Set Id_OT = '" & gstrBusca & "'"
''        mstrSql = mstrSql & " Where Id_Ot='" & gstrBuscaReserva & "' And Seccion_OT='" & gstrSeccion & "'"
''        mstrSql = mstrSql & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''            MsgBox "Error Al Actualizar Los Datos de los Repuestos Reservados"
''        End If
''
''        '/// actualiza repuestos faltantes
''        mstrSql = "Update Tllr_Repuestos_Faltantes "
''        mstrSql = mstrSql & " Set Id_OT = '" & gstrBusca & "'"
''        mstrSql = mstrSql & " Where Id_Ot='" & gstrBuscaReserva & "' And Seccion_OT='" & gstrSeccion & "'"
''        mstrSql = mstrSql & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''            MsgBox "Error Al Actualizar Los Datos De Los Repuestos Faltantes"
''        End If
''    Else
''        Exit Sub
''    End If
''
''    EliminaReserva gstrBuscaReserva          '/// elimina la ot de reserva que fue grabada anteriormente
''    Screen.MousePointer = vbDefault
''End Sub
''Private Sub EliminaReserva(pstrNroReserva)
''    '////////////////////////////////ELIMINAR SERVICIOS DE MECANICA///////////////////////////////////
''    mstrSql = "DELETE FROM Tllr_Mecanica_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '////////////////////////////////ELIMINAR OTROS SERVICIOS ///////////////////////////////////
''    mstrSql = "DELETE FROM Tllr_Otro_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '////////////////////////////////ELIMINAR SERVICIOS DE CARRPCERIA///////////////////////////////////
''    mstrSql = "DELETE FROM Tllr_Carroceria_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '////////////////////////////////ELIMINAR SERVICIOS DE TERCEROS///////////////////////////////////
''    mstrSql = "DELETE FROM Tllr_Terceros_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '////////////////////////////////////ELIMINAR INENTARIO///////////////////////////////
''    mstrSql = "DELETE FROM Tllr_Inventario_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    mstrSql = "DELETE FROM Tllr_Repuestos_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '//////////////////////////////////////ENCABEZADO/////////////////////////////
''    mstrSql = "DELETE FROM Tllr_OT WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''
''End Sub
''Private Sub EliminaReservaRepuestos(pstrNroRegularizacion As String, pstrNroOT As String)
''
''    '//// Elimina Detalle de la Reserva de Repuestos
''    mstrSql = "DELETE FROM Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & pstrNroRegularizacion & "' AND Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '//// Elimina Cabezera de la Reserva de Repuestos
''    mstrSql = "DELETE FROM Stck_Regularizacion Where Id_Regularizacion = '" & pstrNroRegularizacion & "' AND Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '//// Elimina Repuestos Reservados
''    mstrSql = "DELETE FROM Tllr_Repuestos_Reservados WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & pstrNroOT & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''    '//// Elimina Repuestos Faltantes
''    mstrSql = "DELETE FROM Tllr_Repuestos_Faltantes WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & pstrNroOT & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''
''End Sub
''
''Private Sub CancelaReserva()
''Dim mstrMotivoCancela As String
''Dim AdoAnular As New ADODB.Recordset
''
''    If MsgBox("¿ Realmente Desea eliminar esta Reserva de Hora?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
''        If TieneReservadeRepuestos Then
''
''            mstrSql = "Select Id_Regularizacion as Numero from Stck_Regularizacion where id_ot='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''            If Conexion.SendHost(mstrSql, AdoAnular, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
''                With AdoAnular
''                    If Not .BOF And Not .EOF Then
''                        .MoveFirst
''                        While Not .EOF
''                            NroRegularizacion = !NUMERO
''                            Call Actualiza_Saldos_VS_Detalle("S", "Select Canrtidad, Id_Empresa, Id_sucursal, Id_Bodega,Id_Ubicacion,Id_Item From Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & NroRegularizacion & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Empresa = '" & gstrIdEmpresa & "'")
''
''                            EliminaReservaRepuestos NroRegularizacion, lblNroRecepcion
''
''                            .MoveNext
''                        Wend
''                    End If
''                End With
''            End If
''
''        End If
''        EliminaReserva lblNroRecepcion   'OT
''    Else
''        Exit Sub
''    End If
''
''    '/// ingresa motivo de cancelación
''    mstrMotivoCancela = InputBox("Ingrese el motivo de Cancelacion de la Reserva", "Por que Cancela...")
''
''    '/// Actualiza Estado Reserva de Hora
''    mstrSql = " Update Tllr_ReservaHora "
''    mstrSql = mstrSql & " Set Estado = 'E',"
''    mstrSql = mstrSql & " Fecha_Cancelacion='" & Date & "',"
''    mstrSql = mstrSql & " Quien_Cancela='" & gstrUsuario & "',"
''    mstrSql = mstrSql & " MotivoCancela='" & mstrMotivoCancela & "'"
''    mstrSql = mstrSql & " Where Id_Reserva='" & Mid(lblNroRecepcion, 3, 5) & "'"
''    mstrSql = mstrSql & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''        MsgBox "Error Al Actualizar Los Datos De La Reserva de Hora"
''    End If
''
''    CancelarAgregaRegistro
''
''End Sub
''Sub Repuestos_de_la_Mantencion(stridMarca As String, stridModelo As String, stridServicio As String, blnLlenaRepuestos As Boolean)
''Dim adoTemp As New ADODB.Recordset
''
'''lvwRepuestos.ListItems.Clear
''mstrSql = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
''mstrSql = mstrSql & " Stck_Item.Descripcion AS NOMBRE, "
''mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
''mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Valor AS VLR, "
''mstrSql = mstrSql & " Stck_Item.Id_Familia AS IDFAM, "
''mstrSql = mstrSql & " Stck_Item.Precio_Venta as Precio, "
''mstrSql = mstrSql & " Glbl_Familia.Descripcion AS FAMILIA "
''mstrSql = mstrSql & " FROM Glbl_Familia RIGHT OUTER JOIN Stck_Item ON  Glbl_Familia.Id_Familia = Stck_Item.Id_Familia RIGHT OUTER JOIN Tllr_Actividad_Repuesto ON Stck_Item.Id_Item = Tllr_Actividad_Repuesto.Id_Item"
''mstrSql = mstrSql & " WHERE Tllr_Actividad_Repuesto.Id_Marca = '" & stridMarca & "' AND Tllr_Actividad_Repuesto.Id_Modelo = '" & stridModelo & "' AND Tllr_Actividad_Repuesto.Id_Servicio = '" & stridServicio & "' "
''
''
''    If Conexion.SendHost(mstrSql, adoTemp, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
''        With adoTemp
''            If Not .BOF And Not .EOF Then
''                .MoveFirst
''                While Not .EOF
''                    Set lsiItem = lvwRepuestosMantencion.ListItems.Add(, , ValorNulo(!Codigo))
''                    lsiItem.SubItems(1) = ValorNulo(!Nombre)
''                    lsiItem.SubItems(2) = FormatoValor(!CANTY, "", 2)
''                    lsiItem.SubItems(3) = FormatoValor(ValorNulo(!Precio), "", gintDecimalesMoneda)
''                    lsiItem.SubItems(4) = ValorNulo(!Familia)
''                    lsiItem.SubItems(5) = Me.lvwServiciosMecanica.SelectedItem.SubItems(6)
''
''                    If Me.dtcGarantia.BoundText = "PRE" And mstrAgregaPresupuesto = True And blnLlenaRepuestos = True Then
''                        Set itmAux = lvwRepuestos.ListItems.Add(, , ValorNulo(!Codigo))
''                        itmAux.SubItems(1) = ValorNulo(!Nombre)
''                        itmAux.SubItems(2) = FormatoValor(!CANTY, "", 2)
''                        itmAux.SubItems(3) = FormatoValor(!Precio, "", gintDecimalesMoneda)
''                        itmAux.SubItems(4) = FormatoValor(0, "", 2)
''                        itmAux.SubItems(5) = FormatoValor(0, "", gintDecimalesMoneda)
''                        itmAux.SubItems(6) = "" 'TraeCargoDes(gstrIdCargo)
''                        itmAux.SubItems(7) = gstrIdCargo
''                        itmAux.SubItems(8) = Format(Val(SacarFormatoValor(itmAux.SubItems(2), "")) * Val(SacarFormatoValor(itmAux.SubItems(3), "")), "###,##0.00")
''                        itmAux.SubItems(9) = ValorNulo(!IDFAM)
''                        itmAux.SubItems(10) = "N"
''                        itmAux.SubItems(11) = "PRESUPUESTO"
''                    End If
''
''                    .MoveNext
''                Wend
''                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''                TotalFinal
''            End If
''        End With
''    End If
''    Conexion.CloseHost adoTemp
''End Sub
''Sub Quita_Repuestos_Mantencion(stridMarca As String, stridModelo As String, stridServicio As String)
''Dim i As Integer
''
'''lvwRepuestos.ListItems.Clear
''mstrSql = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
''mstrSql = mstrSql & " Stck_Item.Descripcion AS NOMBRE, "
''mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
''mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Valor AS VLR, "
''mstrSql = mstrSql & " Stck_Item.Id_Familia AS IDFAM, "
''mstrSql = mstrSql & " Glbl_Familia.Descripcion AS FAMILIA "
''mstrSql = mstrSql & " FROM Glbl_Familia RIGHT OUTER JOIN Stck_Item ON  Glbl_Familia.Id_Familia = Stck_Item.Id_Familia RIGHT OUTER JOIN Tllr_Actividad_Repuesto ON Stck_Item.Id_Item = Tllr_Actividad_Repuesto.Id_Item"
''mstrSql = mstrSql & " WHERE Tllr_Actividad_Repuesto.Id_Marca = '" & stridMarca & "' AND Tllr_Actividad_Repuesto.Id_Modelo = '" & stridModelo & "' AND Tllr_Actividad_Repuesto.Id_Servicio = '" & stridServicio & "' "
''
''
''    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
''        With adoPrincipal
''            If Not .BOF And Not .EOF Then
''                .MoveFirst
''                While Not .EOF
''                    Set itmAux = lvwRepuestos.FindItem(!Codigo, lvwText, , 0)
''                    If Not itmAux Is Nothing Then   ' Si no hay coincidencia
''                        lvwRepuestos.ListItems.Remove (lvwRepuestos.FindItem(!Codigo).Index)
''
''                    End If
''                    .MoveNext
''                Wend
''                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
''                TotalFinal
''            End If
''        End With
''    End If
''
''End Sub
''
''Function TieneReservadeRepuestos() As Boolean
''Dim lstrEstadoReserva As String
''
''    TieneReservadeRepuestos = False
''
''    lstrEstadoReserva = Retorna_Valor_General("Select estado_reserva As Codigo From Tllr_Ot where id_ot='" & Me.lblNroRecepcion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Seccion_OT='" & gstrSeccion & "'")
''    If lstrEstadoReserva = "R" Then
''        TieneReservadeRepuestos = True
''    End If
''End Function
''Sub GrabaReservaRepuestosRecepcion()
''    If Me.Tag = "Crear" Then
''        lblNroRecepcion = TraeCorrelativo(gcOrdenTrabajo, gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
''        gstrBusca = lblNroRecepcion
''        mstrSql = "INSERT INTO Tllr_OT "
''        mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal, "
''        mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''        mstrSql = mstrSql & " Id_Garantia, Folio_Garantia, "
''        mstrSql = mstrSql & " Id_Tipo_Cono, Nro_Cono, "
''        mstrSql = mstrSql & " Patente, RealizadoPor,"
''        mstrSql = mstrSql & " Kilometros_Recepcion, Id_Compañia_seguro,"
''        mstrSql = mstrSql & " Fecha_Proxima_Visita, "                           'Fecha_Liquidacion,"
''        mstrSql = mstrSql & " Estado,Fecha_Emision, "
''        mstrSql = mstrSql & " Entrega_Estimada, Hora_Entrega, "
''        mstrSql = mstrSql & " Nro_Factura_Emitida,Nro_Presupuesto_Origen,"
''        mstrSql = mstrSql & " Nro_Siniestro, Nro_Poliza, Liquidador, "
''        mstrSql = mstrSql & " Comentario, Solicitado_Por,"
''        mstrSql = mstrSql & " Deducible_UF , Deducible_Pesos, "
''        mstrSql = mstrSql & " Total_Mecanica,Total_Carroceria,"
''        mstrSql = mstrSql & " Total_Desabolladura,Total_Pintura,"
''        mstrSql = mstrSql & " Total_Terceros,Total_Repuestos,"
''        mstrSql = mstrSql & " Total_Materiales,Total_Insumos, "
''        mstrSql = mstrSql & " Total_Otros,Total_Ot,"
''        mstrSql = mstrSql & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor, ReparacionMantencion, Estado_Reserva ) "
''        mstrSql = mstrSql & " VALUES ("
''        mstrSql = mstrSql & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
''        mstrSql = mstrSql & " '" & lblNroRecepcion & "', '" & gstrSeccion & "',"
''        mstrSql = mstrSql & " '" & Trim(dtcGarantia.BoundText) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), "S/F") & "',"
''        mstrSql = mstrSql & " '" & dtcTipoCono.BoundText & "', " & CLng(txtNroCono.Text) & ","
''        mstrSql = mstrSql & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
''        mstrSql = mstrSql & " " & CLng(txtKilAct) & ", '" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"   'OJO
''        mstrSql = mstrSql & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
''        mstrSql = mstrSql & " 'V','" & CDate(pckFechaAtencion.Value) & "', "
''        mstrSql = mstrSql & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
''        mstrSql = mstrSql & " '" & "S/N" & "', '" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "',"
''        mstrSql = mstrSql & " '" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ','" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "','" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "' , "
''        mstrSql = mstrSql & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), "S/S") & "' ,"
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & " ," & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", " & gcurInsumo & ", "
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & ", " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & " ,"
''        mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & " ," & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ","
''        mstrSql = mstrSql & " '" & lblIdCliente & "',"
''        mstrSql = mstrSql & " '" & IIf(optMantencion.Value = True, "M", "R") & "',"
''        mstrSql = mstrSql & " '" & IIf(cmdReserva.Enabled = False, "R", "N") & "')"
''    Else
''        mstrSql = "UPDATE Tllr_OT "
''        mstrSql = mstrSql & " SET Id_Garantia='" & Trim(dtcGarantia.BoundText) & "', "
''        mstrSql = mstrSql & " Folio_Garantia='" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), ".") & "', "
''        mstrSql = mstrSql & " Id_Tipo_Cono='" & dtcTipoCono.BoundText & "', "
''        mstrSql = mstrSql & " Nro_Cono=" & CLng(txtNroCono.Text) & ", "
''        mstrSql = mstrSql & " Patente='" & txtPatente.Text & "', "
''        mstrSql = mstrSql & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
''        mstrSql = mstrSql & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
''        mstrSql = mstrSql & " Entrega_Estimada='" & CDate(pckFechaEntrega) & "', "
''        mstrSql = mstrSql & " Hora_Entrega='" & cboHora.Text & "', "
''        mstrSql = mstrSql & " Nro_Siniestro='" & IIf(txtNroSiniestro <> "", UCase(Trim(txtNroSiniestro)), "S/N") & " ', "
''        mstrSql = mstrSql & " Nro_Poliza='" & IIf(txtNroPoliza <> "", UCase(Trim(txtNroPoliza)), "S/N") & "', "
''        mstrSql = mstrSql & " Liquidador='" & IIf(txtLiquidador <> "", UCase(Trim(txtLiquidador)), "S/L") & "', "
''        mstrSql = mstrSql & " Comentario='" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), ".") & "', "
''        mstrSql = mstrSql & " Solicitado_Por='" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), ".") & "',"
''        mstrSql = mstrSql & " Total_Mecanica=" & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Carroceria=" & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " Total_Desabolladura=" & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Pintura=" & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " Total_Terceros=" & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Repuestos=" & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
''        mstrSql = mstrSql & " Total_Otros=" & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & "  ,"
''        mstrSql = mstrSql & " Total_Materiales=" & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", "
''        mstrSql = mstrSql & " Total_Insumos=" & gcurInsumo & ", "
''        mstrSql = mstrSql & " Total_Ot=" & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo & "  ,"
''        mstrSql = mstrSql & " Total_OT_Iva=" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & "  ,"
''        mstrSql = mstrSql & " Total_IVA =" & (CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & "  ,"
''        mstrSql = mstrSql & " Deducible_UF = " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , "
''        mstrSql = mstrSql & " Deducible_Pesos = " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
''        mstrSql = mstrSql & " Nro_Presupuesto_Origen='" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "', "
''        mstrSql = mstrSql & " Kilometros_Recepcion=" & CLng(txtKilAct) & ","
''        mstrSql = mstrSql & " Id_Compañia_Seguro='" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"
''        mstrSql = mstrSql & " Fecha_Proxima_Visita = '" & DateAdd("d", 365, pckFechaAtencion.Value) & "',"
''        mstrSql = mstrSql & " Id_Cliente_Proveedor='" & lblIdCliente & "',"
''        mstrSql = mstrSql & " ReparacionMantencion='" & IIf(Me.optMantencion.Value = True, "M", "R") & "',"
''        mstrSql = mstrSql & " Estado_Reserva='" & IIf(Me.cmdReserva = False, "R", "N") & "'"
''        mstrSql = mstrSql & " WHERE Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal ='" & gstrIdSucursal & "' And Id_OT ='" & Trim(Trim(lblNroRecepcion)) & "' AND Seccion_OT ='" & gstrSeccion & "' "
''    End If                                                                                                                                                                                                                                                                              ''" & pckFechaVenta.Value & "'
''
''    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
''        If GuardaMecanica(lblNroRecepcion, gcOrdenTrabajo) = False Then
''            MsgBox LoadResString(321)
''        End If
''
'''//////////////////////////////////
''        mblnTablaVacia = False
''        ActivaBotones
''        Me.Tag = ""
''    End If
''End Sub
''Sub LiquidarPresupuesto()
''
''    Screen.MousePointer = vbDefault
''    'Pregunto si el presupuesto lo va agregar a una ot existente
''    gstrProcedencia = "Movimientos"
''    frmPresupuestoAdicional.Show vbModal
''    gstrProcedencia = "Presupuestos"
''    mstrLiquidaPresupuesto = True
''    Screen.MousePointer = vbHourglass
''    If gintOtExistente = 2 Then         'ot nueva
''        gstrBuscaReserva = lblNroRecepcion
''        mstrIdPresupuestoOrigen = lblNroRecepcion
''        Me.Tag = "Crear"
''        dtcGarantia.BoundText = gstrIdTipoOtDefecto
''        GrabarRegistro                                              '/// graba el presupuesto en una ot definitiva
''        GrabarPresupuesto gstrBuscaReserva, gstrBusca, "L", ""      '/// Graba presupuesto en tablas de presupuesto
''        EliminaReserva gstrBuscaReserva                             '/// elimina el presupuesto que fue grabado anteriormente como OT
''
''        MsgBox "Fue Creada la OT Numero : " & gstrBusca, vbInformation, "OT"
''
''        UltimoRegistro
''
''    ElseIf gintOtExistente = 1 Then         'ot existente
''
''        Dim lstrNumeroPresupuesto As String
''        Dim lstrSQL As String
''
''        If GuardaMecanicaPresupuesto(gstrBusca, gstrSeccion) = False Then
''            MsgBox LoadResString(321)
''        End If
''        If GuardaCarroceriaPresupuesto(gstrBusca, gstrSeccion) = False Then
''            MsgBox LoadResString(320)
''        End If
''        If GuardaOtrosPresupuesto(gstrBusca, gstrSeccion) = False Then
''            MsgBox LoadResString(328)
''        End If
''        If GuardaTercerosPresupuesto(gstrBusca, gstrSeccion) = False Then
''            MsgBox LoadResString(319)
''        End If
''        If gblnTraspasaRepuestos = True Then
''            If GuardaRepuestosPresupuesto(gstrBusca, gstrSeccion) = False Then
''                MsgBox LoadResString(318)
''            End If
''        End If
''        mstrIdPresupuestoOrigen = lblNroRecepcion
''        GrabarPresupuesto lblNroRecepcion, gstrBusca, "L", ""       '/// Graba presupuesto en tablas de presupuesto
''        EliminaReserva lblNroRecepcion                              '/// elimina el presupuesto que fue grabado anteriormente como OT
''        'actualiza id_presupuesto de tllr_ot
''        lstrNumeroPresupuesto = Retorna_Valor_General("Select Nro_Presupuesto_Origen from Tllr_OT Where id_ot='" & gstrBusca & "' And Seccion_OT='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
''
''        lstrSQL = "Update Tllr_OT Set Nro_Presupuesto_Origen='" & lstrNumeroPresupuesto & "/" & lblNroRecepcion & "' "
''        lstrSQL = lstrSQL & "Where Id_ot='" & gstrBusca & "' And Seccion_OT='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''        If Conexion.SendHost(lstrSQL, , , , gcTiempoEspera) = apAbort Then
''            MsgBox "Problemas para actualizar el numero de presupuesto", vbInformation, "Actualización"
''        End If
''
''        UltimoRegistro
''    End If
''    mstrLiquidaPresupuesto = False
''End Sub
''Sub AnularPresupuesto()
''Dim mstrMotivoAnula As String
''
''    Screen.MousePointer = vbHourglass
''    mstrMotivoAnula = InputBox("Ingrese El Motivo por que Anula :", "Por que Anula Presupuesto....")
''    If mstrMotivoAnula <> "" Then
''        gstrBuscaReserva = lblNroRecepcion
''        GrabarPresupuesto gstrBuscaReserva, "S/N", "N", mstrMotivoAnula    '/// Graba presupuesto en tablas de presupuesto
''        EliminaReserva gstrBuscaReserva          '/// elimina el presupuesto que fue grabado anteriormente como OT
''        Renovar
''    End If
''End Sub
''Function NumerosDocumentos(IdOT As String, SeccionOT As String) As String
''Dim adoTemp As New ADODB.Recordset
''Dim lstrSQL As String
''
''    NumerosDocumentos = ""
''    lstrSQL = "Select Nro_Factura_Emitida from Tllr_Facturacion where id_Ot='" & IdOT & "' And Seccion_OT='" & SeccionOT & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
''    If Conexion.SendHost(lstrSQL, adoTemp, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''        With adoTemp
''            While Not .EOF
''                NumerosDocumentos = NumerosDocumentos & ValorNulo(!Nro_Factura_Emitida) & "/"
''                adoTemp.MoveNext
''            Wend
''        End With
''    End If
''End Function
''Private Function GuardaMecanicaPresupuesto(strIdOt As String, strSeccion As String) As Boolean
''
'''valida que no exista ya el servicio
''If Me.lvwServiciosMecanica.ListItems.Count > 0 Then
''    If ValidaServicioMecanica(strIdOt, strSeccion, Trim(lblIdMarca), Trim(lblIdModelo), IIf(Me.lvwServiciosMecanica.ListItems.Count > 0, Trim(Me.lvwServiciosMecanica.SelectedItem), "")) = True Then
''        Exit Function
''    End If
''End If
''GuardaMecanicaPresupuesto = True
''With lvwServiciosMecanica
''    If .ListItems.Count > 0 Then
''        For intIndice = 1 To .ListItems.Count
''        Set .SelectedItem = .ListItems(intIndice)
''        mstrSql = "Insert Into Tllr_Mecanica_OT "
''        mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''        mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''        mstrSql = mstrSql & " Id_Marca, Id_Modelo, "
''        mstrSql = mstrSql & " Id_Servicio, "
''        mstrSql = mstrSql & " Id_Tipo_Cargo,Mecanico_Designado,"
''        mstrSql = mstrSql & " Horas,Valor,"
''        mstrSql = mstrSql & " Porcentaje_Descuento, Monto_Descuento, "
''        mstrSql = mstrSql & " SubTotal, Facturado)"
''        mstrSql = mstrSql & " Values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
''        mstrSql = mstrSql & " '" & strIdOt & "', '" & strSeccion & "',"
''        mstrSql = mstrSql & " '" & Trim(lblIdMarca) & "','" & Trim(lblIdModelo) & "',"
''        mstrSql = mstrSql & " '" & Trim(.SelectedItem) & "',"
''        mstrSql = mstrSql & " '" & .SelectedItem.SubItems(6) & "'," & IIf(.SelectedItem.SubItems(8) = "", "NULL", " '" & .SelectedItem.SubItems(8) & "' ") & ", "
''        mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & " , " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & " , "
''        mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & " ," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''        mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & .SelectedItem.SubItems(11) & "' )"
''        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''            GuardaMecanicaPresupuesto = False
''            Exit Function
''        End If
''        Next
''    Else
''        GuardaMecanicaPresupuesto = True
''    End If
''End With
''End Function
''
''Private Function GuardaCarroceriaPresupuesto(strIdOt As String, strSeccion As String) As Boolean
''
''GuardaCarroceriaPresupuesto = True
''With lvwServiciosCarroceria
''    If .ListItems.Count > 0 Then
''        For intIndice = 1 To .ListItems.Count
''            Set .SelectedItem = .ListItems(intIndice)
''            '/////////////////////////////////////////////////VALIDAR SI EXISTE EN PARENT
''            'If ExisteRegistro(strCiaSeguro, .SelectedItem.SubItems(1), .SelectedItem.SubItems(4)) = True Then
''                mstrSql = "INSERT INTO Tllr_Carroceria_OT"
''                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''                mstrSql = mstrSql & " Id_Compañia_Seguro, "
''                mstrSql = mstrSql & " Id_Concepto, "
''                mstrSql = mstrSql & " D_P,"
''                mstrSql = mstrSql & " Id_Parte_Pieza, "
''                mstrSql = mstrSql & " Id_Tipo_Cargo, Mecanico_Designado,"
''                mstrSql = mstrSql & " Horas, Valor,Valor_Definido ,"
''                mstrSql = mstrSql & " Porcentaje_Descuento,Monto_Descuento,"
''                mstrSql = mstrSql & " SubTotal,Facturado,Porcentaje_Recargo,Monto_Recargo,Id_Proveedor,Descripcion,Id_Servicio_Carroceria)"
''                mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "       '///empresa, sucursal
''                mstrSql = mstrSql & " '" & strIdOt & "', '" & strSeccion & "',"                  '///nro ot, seccion
''                mstrSql = mstrSql & " '" & frmRecepcion.lblCompañia.Tag & "', "                                         '///cia seguro
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(1)) & "', "                      '///concepto
''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(3) & "',"                                                   'Trim(.SelectedItem.SubItems(2)) ///d_p
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(4)) & "', "                      '///parte y pieza
''                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(13) & "','" & gstrMecanicoDefectoSecCar & "',"            '///mecanico designado
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(5), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(6), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(9), "######.00"))) & " ,"
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "######.00"))) & ","
''                mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(16), "######.00"))) & ",'" & .SelectedItem.SubItems(17) & "',"
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "######.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "######.00"))) & ","
''                mstrSql = mstrSql & " " & IIf(.SelectedItem.SubItems(15) = "", "NULL" & ",", " '" & .SelectedItem.SubItems(15) & "',")
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(2)) & "',"
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(18)) & "')"
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaCarroceriaPresupuesto = False
''                    Exit Function
''                End If
''            'End If
''        Next
''    Else
''        GuardaCarroceriaPresupuesto = True
''    End If
''End With
''End Function
''Private Function GuardaOtrosPresupuesto(strIdOt As String, strSeccion As String) As Boolean
''
''GuardaOtrosPresupuesto = True
''With lvwOtrosServicios
''    If .ListItems.Count > 0 Then
''        For intIndice = 1 To .ListItems.Count
''            Set .SelectedItem = .ListItems(intIndice)
''            mstrSql = "INSERT INTO Tllr_Otro_OT"
''            mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''            mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''            mstrSql = mstrSql & " Id_Otro_Servicio, "
''            mstrSql = mstrSql & " Id_Tipo_Cargo,"
''            mstrSql = mstrSql & " Mecanico_Asignado, "
''            mstrSql = mstrSql & " Horas,Valor,"
''            mstrSql = mstrSql & " Porcentaje_Descuento,Monto_Descuento,"
''            mstrSql = mstrSql & " SubTotal,Descripcion_Otro,Facturado)"
''            mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
''            mstrSql = mstrSql & " '" & strIdOt & "', '" & strSeccion & "',"
''            mstrSql = mstrSql & " '" & .SelectedItem & "', "
''            mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(6)) & "', "
''            mstrSql = mstrSql & " '" & IIf(Trim(.SelectedItem.SubItems(8)) = "", "SIN", Trim(.SelectedItem.SubItems(8))) & "', "
''            mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
''            mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''            mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & UCase(Trim(.SelectedItem.SubItems(1))) & "','" & UCase(Trim(.SelectedItem.SubItems(11))) & "')"
''            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                GuardaOtrosPresupuesto = False
''                Exit Function
''            End If
''        Next
''    Else
''        GuardaOtrosPresupuesto = True
''    End If
''End With
''End Function
''Private Function GuardaTercerosPresupuesto(strIdOt As String, strSeccion As String) As Boolean
''
''GuardaTercerosPresupuesto = True
''With lvwServiciosTerceros
''    If .ListItems.Count > 0 Then
''        For intIndice = 1 To .ListItems.Count
''            Set .SelectedItem = .ListItems(intIndice)
''            mstrSql = "INSERT INTO Tllr_Terceros_OT"
''            mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''            mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''            mstrSql = mstrSql & " Id_Proveedor, "
''            mstrSql = mstrSql & " Id_Servicio_Tercero,"
''            mstrSql = mstrSql & " Id_Tipo_Cargo, "
''            mstrSql = mstrSql & " Cantidad,Valor,"
''            mstrSql = mstrSql & " Porcentaje_Recargo,Monto_Recargo,"
''            mstrSql = mstrSql & " Precio_Final,"
''            mstrSql = mstrSql & " Descripcion , NroFarctura, "
''            mstrSql = mstrSql & " SubTotal, Facturado, "
''            mstrSql = mstrSql & " Porcentaje_Dscto, Monto_Dscto)"
''            mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
''            mstrSql = mstrSql & " '" & strIdOt & "', '" & strSeccion & "',"
''            mstrSql = mstrSql & " '" & .SelectedItem.SubItems(2) & "', "
''            mstrSql = mstrSql & " '" & Trim(.SelectedItem) & "', "
''            mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(14)) & "', "
''            mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(6), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''            mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
''            mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(9), "#####0.00"))) & ","
''            mstrSql = mstrSql & " '" & .SelectedItem.SubItems(3) & "', "
''            mstrSql = mstrSql & " '" & .SelectedItem.SubItems(4) & "', "
''            mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(12), "#####0.00"))) & ",'" & .SelectedItem.SubItems(15) & "',"
''            mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "#####0.00"))) & ")"
''            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                GuardaTercerosPresupuesto = False
''                Exit Function
''            End If
''        Next
''    Else
''        GuardaTercerosPresupuesto = True
''    End If
''End With
''End Function
''
''Private Function GuardaRepuestosPresupuesto(strIdOt As String, strSeccion As String) As Boolean
''
'''primero actualiza tllr_repuestos_ot
''
''GuardaRepuestosPresupuesto = True
''
''With lvwRepuestos
''    If .ListItems.Count > 0 Then
''        For intIndice = 1 To .ListItems.Count
''            Set .SelectedItem = .ListItems(intIndice)
''            If VerificaRepuesto(.SelectedItem, strIdOt, strSeccion, "Tllr_Repuestos_OT") = True Then
''                mstrSql = "UPDATE Tllr_Repuestos_OT"
''                mstrSql = mstrSql & " SET Id_Tipo_Cargo='" & Trim(.SelectedItem.SubItems(7)) & "',"
''                mstrSql = mstrSql & " Cantidad = " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & ", "
''                mstrSql = mstrSql & " Valor = " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
''                mstrSql = mstrSql & " Porcentaje_Descuento = " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & ","
''                mstrSql = mstrSql & " Monto_Descuento = " & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''                mstrSql = mstrSql & " SubTotal = " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
''                mstrSql = mstrSql & " Facturado = " & UCase(Trim(IIf(.SelectedItem.SubItems(10) = "", "'N'", "'" & .SelectedItem.SubItems(10) & "'"))) & ","
''           '     mstrSql = mstrSql & " cantidad_real = " & CDbl(Val(Format(.SelectedItem.SubItems(13), "#####0.0"))) & ", "
''                mstrSql = mstrSql & " Consumo = '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "'"
''                mstrSql = mstrSql & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
''                mstrSql = mstrSql & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
''                mstrSql = mstrSql & " Id_OT = '" & strIdOt & "' AND  "
''                mstrSql = mstrSql & " Seccion_OT = '" & strSeccion & "' AND "
''                mstrSql = mstrSql & " Id_Item = '" & .SelectedItem & "' "
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaRepuestosPresupuesto = False
''                    Exit Function
''                End If
''            Else
''                '///////////////////////////////////VALIDAR SI EXISTE EN PARENT
''                mstrSql = "INSERT INTO Tllr_Repuestos_OT"
''                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''                mstrSql = mstrSql & " Id_Item, "
''                mstrSql = mstrSql & " Id_Tipo_Cargo, "
''                mstrSql = mstrSql & " Cantidad, Valor,"
''                mstrSql = mstrSql & " Porcentaje_Descuento,Monto_Descuento,"
''                mstrSql = mstrSql & " SubTotal,Facturado,Consumo)"
''                mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
''                mstrSql = mstrSql & " '" & strIdOt & "', '" & strSeccion & "',"
''                mstrSql = mstrSql & " '" & .SelectedItem & "', "
''                mstrSql = mstrSql & " '" & Trim(.SelectedItem.SubItems(7)) & "', "
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ",'" & .SelectedItem.SubItems(10) & "',"
''                mstrSql = mstrSql & " '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "')"
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaRepuestosPresupuesto = False
''                    Exit Function
''                End If
''            End If
''        Next
''    Else
''        GuardaRepuestosPresupuesto = True
''    End If
''End With
''
''
'''Ahora actualiza repuestos reservados
''GuardaRepuestosPresupuesto = True
''
''With lvwRepuestos
''    If .ListItems.Count > 0 Then
''        For intIndice = 1 To .ListItems.Count
''            Set .SelectedItem = .ListItems(intIndice)
''            If VerificaRepuesto(.SelectedItem, strIdOt, strSeccion, "Tllr_Repuestos_Reservados") = True Then
''                mstrSql = "UPDATE Tllr_Repuestos_Reservados"
''                mstrSql = mstrSql & " SET Solicitado = " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & ", "
''                mstrSql = mstrSql & " Precio_Unitario = " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
''                mstrSql = mstrSql & " Reservado= " & 0 & ","
''                mstrSql = mstrSql & " Estado = 'V'" & ","
''                mstrSql = mstrSql & " Tipo = 'Q'"
''                mstrSql = mstrSql & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
''                mstrSql = mstrSql & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
''                mstrSql = mstrSql & " Id_OT = '" & strIdOt & "' AND  "
''                mstrSql = mstrSql & " Seccion_OT = '" & strSeccion & "' AND "
''                mstrSql = mstrSql & " Id_Item = '" & .SelectedItem & "' "
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaRepuestosPresupuesto = False
''                    Exit Function
''                End If
''            Else
''                '///////////////////////////////////VALIDAR SI EXISTE EN PARENT
''                mstrSql = "INSERT INTO Tllr_Repuestos_Reservados"
''                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
''                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
''                mstrSql = mstrSql & " Id_Item, "
''                mstrSql = mstrSql & " Precio_Unitario,Solicitado,"
''                mstrSql = mstrSql & " Reservado,Estado,Tipo)"
''                mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
''                mstrSql = mstrSql & " '" & strIdOt & "', '" & strSeccion & "',"
''                mstrSql = mstrSql & " '" & .SelectedItem & "', "
''                mstrSql = mstrSql & " " & CDbl(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & ","
''                mstrSql = mstrSql & " " & 0 & ",'V','Q')"
''                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
''                    GuardaRepuestosPresupuesto = False
''                    Exit Function
''                End If
''            End If
''        Next
''    Else
''        GuardaRepuestosPresupuesto = True
''    End If
''End With
''
''
''End Function
''
''Sub ActualizarSaldoRepuestos(strIdDocumento, strSeccion)
''Dim i As Integer
''For intIndice = 1 To lvwRepuestos.ListItems.Count
''    mstrSql = "UPDATE Tllr_Repuestos_OT SET Saldo='" & lvwRepuestos.ListItems(intIndice).SubItems(12) & "'"
''    mstrSql = mstrSql & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
''    mstrSql = mstrSql & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
''    mstrSql = mstrSql & " Id_OT = '" & strIdDocumento & "' AND  "
''    mstrSql = mstrSql & " Seccion_OT = '" & strSeccion & "' AND "
''    mstrSql = mstrSql & " Id_Item = '" & lvwRepuestos.ListItems(intIndice) & "' "
''    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''Next
''End Sub
''Sub VerificaCampañas()
''Dim adoTemp As New ADODB.Recordset
''Dim lstrSQL As String
''
''    lstrSQL = "Select Vin,Id_Item,Servicio from Tllr_Campañas where Vin='" & Me.lblVin & "' And Estado='V' And Fecha_Inicio <='" & Format(Date, "DD/MM/YYYY") & "' And Fecha_Termino>='" & Format(Date, "DD/MM/YYYY") & "'"
''    If Conexion.SendHost(lstrSQL, adoTemp, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
''        With adoTemp
''            While Not .EOF
''                If MsgBox("Campaña:" & Chr(13) & adoTemp!servicio & Chr(13) & "esta VIGENTE la Realiza ahora ? ", vbInformation + vbYesNo, "Advertencia") = vbYes Then
''                    txtComentario = Me.txtComentario & "Campaña :  " & adoTemp!servicio
''                    mstrSql = "Update Tllr_Campañas Set Estado='T' Where Vin='" & Me.lblVin & "' And Id_Item='" & adoTemp!Id_Item & "'"
''                    Conexion.SendHost mstrSql, , , , gcTiempoEspera
''                End If
''                adoTemp.MoveNext
''            Wend
''        End With
''    End If
''
''End Sub
''
''Private Sub txtSolicita_KeyPress(KeyAscii As Integer)
'''kjcv  08-02-12
''KeyAscii = UpCaseLetter(KeyAscii)
''
''End Sub
''
''
''

