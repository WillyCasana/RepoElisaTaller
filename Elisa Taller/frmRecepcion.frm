VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecepcion 
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   13830
   Icon            =   "frmRecepcion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   13830
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Width           =   12210
      Begin VB.TextBox txtCorreSpiga 
         Height          =   315
         Left            =   10560
         TabIndex        =   161
         Top             =   165
         Width           =   1575
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
         Left            =   12720
         TabIndex        =   156
         Top             =   165
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox lblNroRecepcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   180
         Width           =   2100
      End
      Begin MSComCtl2.DTPicker pckFechaAtencion 
         Height          =   315
         Left            =   5070
         TabIndex        =   67
         Top             =   165
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95551489
         CurrentDate     =   36776
      End
      Begin VB.Label Label9 
         Caption         =   "Num Spiga"
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
         Left            =   9480
         TabIndex        =   162
         Top             =   165
         Width           =   975
      End
      Begin VB.Label lblTipo 
         Caption         =   "TIPO"
         Height          =   255
         Left            =   12240
         TabIndex        =   155
         Top             =   240
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   72
         Top             =   165
         Width           =   1815
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
         TabIndex        =   69
         Top             =   240
         Width           =   1035
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
         TabIndex        =   27
         Top             =   225
         Width           =   1290
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
         TabIndex        =   26
         Top             =   240
         Width           =   1275
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
      TabIndex        =   119
      Top             =   0
      Width           =   13830
      _ExtentX        =   24395
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
   Begin VB.Frame Frame8 
      Caption         =   "Sección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9960
      TabIndex        =   58
      Top             =   480
      Visible         =   0   'False
      Width           =   2790
      Begin VB.OptionButton optRecepcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Carrocería"
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
         Height          =   300
         Index           =   1
         Left            =   1500
         TabIndex        =   60
         Tag             =   "Carrocería"
         Top             =   195
         Width           =   1230
      End
      Begin VB.OptionButton optRecepcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Mecánica"
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
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   59
         Tag             =   "Mecánica"
         Top             =   195
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin TabDlg.SSTab stbServicios 
      Height          =   6495
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   11456
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
      TabPicture(0)   =   "frmRecepcion.frx":038A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fmeCia"
      Tab(0).Control(1)=   "fmePat"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Inventario Recepción - Comentario"
      TabPicture(1)   =   "frmRecepcion.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeCom"
      Tab(1).Control(1)=   "fmeInv"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Mecánica"
      TabPicture(2)   =   "frmRecepcion.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeMec"
      Tab(2).Control(1)=   "stbTotalMec"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Carroceria"
      TabPicture(3)   =   "frmRecepcion.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "stbTotalPintura"
      Tab(3).Control(1)=   "stbTotalCarroceria"
      Tab(3).Control(2)=   "stbTotalArmeyDesarme"
      Tab(3).Control(3)=   "stbTotalDesabolladura"
      Tab(3).Control(4)=   "fmeCar"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Trabajos Adicionales"
      TabPicture(4)   =   "frmRecepcion.frx":03FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeOtr"
      Tab(4).Control(1)=   "stbTotalOtros"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Trabajos de Terceros"
      TabPicture(5)   =   "frmRecepcion.frx":0416
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "stbTotalTerceros"
      Tab(5).Control(1)=   "fmeTer"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Repuestos"
      TabPicture(6)   =   "frmRecepcion.frx":0432
      Tab(6).ControlEnabled=   -1  'True
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
         Left            =   1665
         TabIndex        =   125
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
      Begin VB.Frame fmePat 
         Height          =   4515
         Left            =   -75000
         TabIndex        =   28
         Top             =   360
         Width           =   12900
         Begin VB.TextBox txtCorreo 
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
            Left            =   10560
            TabIndex        =   177
            Top             =   4080
            Width           =   2130
         End
         Begin VB.TextBox txtTelefono 
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
            Left            =   8400
            TabIndex        =   176
            Top             =   4080
            Width           =   1410
         End
         Begin VB.TextBox txtChasis 
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
            TabIndex        =   163
            Top             =   2100
            Width           =   2820
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
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   148
            Text            =   "0"
            Top             =   1020
            Width           =   1095
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
            TabIndex        =   140
            Top             =   990
            Width           =   1200
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
            Left            =   8640
            TabIndex        =   134
            Top             =   975
            Visible         =   0   'False
            Width           =   1290
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
            Left            =   7320
            TabIndex        =   133
            Top             =   975
            Visible         =   0   'False
            Width           =   1305
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
            Left            =   8040
            MaxLength       =   30
            TabIndex        =   120
            Top             =   240
            Width           =   1515
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
            Left            =   10560
            MaxLength       =   3
            TabIndex        =   111
            Top             =   2955
            Visible         =   0   'False
            Width           =   180
         End
         Begin MSComCtl2.DTPicker pckFecVta 
            Height          =   315
            Left            =   9360
            TabIndex        =   2
            Top             =   2505
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   95551489
            CurrentDate     =   36796
         End
         Begin VB.TextBox txtRut 
            Height          =   315
            Left            =   8955
            MaxLength       =   50
            TabIndex        =   109
            Top             =   4665
            Width           =   2085
         End
         Begin VB.TextBox txtComuna 
            Height          =   315
            Left            =   4710
            MaxLength       =   50
            TabIndex        =   108
            Top             =   4650
            Width           =   4185
         End
         Begin VB.TextBox txtDir 
            Height          =   315
            Left            =   435
            MaxLength       =   50
            TabIndex        =   107
            Top             =   4635
            Width           =   4185
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
            Left            =   5175
            TabIndex        =   1
            Top             =   2520
            Width           =   2850
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
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   8
            Top             =   4080
            Width           =   2865
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
            TabIndex        =   4
            Top             =   3555
            Width           =   930
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
            Left            =   8235
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   29
            Top             =   1695
            Width           =   600
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
            TabIndex        =   0
            Top             =   2535
            Width           =   1380
         End
         Begin VB.ComboBox cboHora 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   11400
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   2880
            Visible         =   0   'False
            Width           =   1170
         End
         Begin MSComCtl2.DTPicker pckFechaEntrega 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   4005
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   95551489
            CurrentDate     =   36733
         End
         Begin MSDataListLib.DataCombo dtcTipoCono 
            Bindings        =   "frmRecepcion.frx":044E
            Height          =   315
            Left            =   1080
            TabIndex        =   3
            Top             =   3525
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
            Top             =   3645
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
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImgBarraHerramienta"
            DisabledImageList=   "ImgBarraHerramienta"
            HotImageList    =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
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
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Presupuesto"
                  ImageIndex      =   28
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcRecepcionista 
            Bindings        =   "frmRecepcion.frx":0468
            Height          =   315
            Left            =   8025
            TabIndex        =   5
            Top             =   3540
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
            Top             =   3660
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
            Left            =   12360
            Top             =   600
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
                  Picture         =   "frmRecepcion.frx":0487
                  Key             =   "Crear"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":0599
                  Key             =   "Menos"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":09F1
                  Key             =   "Mas"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":0E49
                  Key             =   "Persona"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":12A1
                  Key             =   "Editar"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":13B3
                  Key             =   "Grabar"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":14C5
                  Key             =   "Cancelar"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":15D7
                  Key             =   "Borrar"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":16E9
                  Key             =   "Buscar"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":17FB
                  Key             =   "Imprimir"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":190D
                  Key             =   "Cerrar"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":1A1F
                  Key             =   "Ayuda"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":1B31
                  Key             =   "Primero"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":1C43
                  Key             =   "Anterior"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":1D55
                  Key             =   "Siguiente"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":1E67
                  Key             =   "Ultimo"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":1F79
                  Key             =   "Renovar"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":208B
                  Key             =   "SortAsc"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":219D
                  Key             =   "SortDesc"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":22AF
                  Key             =   "Seleccion"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":2701
                  Key             =   "Seleccion1"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":2B53
                  Key             =   "Copiar"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":2C65
                  Key             =   "Vaciar"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":30B9
                  Key             =   "Confirmar"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":33D5
                  Key             =   "LiquidarPres"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":382D
                  Key             =   "AnularPres"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":3C81
                  Key             =   "Salir"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRecepcion.frx":3FD3
                  Key             =   "list"
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcGarantia 
            Bindings        =   "frmRecepcion.frx":40E5
            Height          =   315
            Left            =   1080
            TabIndex        =   121
            Top             =   285
            Width           =   2175
            _ExtentX        =   3836
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
            TabIndex        =   143
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
         Begin MSDataListLib.DataCombo dtcTrabajo 
            Bindings        =   "frmRecepcion.frx":40FF
            Height          =   315
            Left            =   4800
            TabIndex        =   166
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
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
         Begin MSAdodcLib.Adodc datTrabajo 
            Height          =   330
            Left            =   4920
            Top             =   240
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
         Begin MSDataListLib.DataCombo dbcboTipoVenta 
            Bindings        =   "frmRecepcion.frx":4118
            Height          =   315
            Left            =   13920
            TabIndex        =   171
            Top             =   2505
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   "Descripcion"
            BoundColumn     =   "id_Tipo_Venta"
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
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Correo"
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
            Index           =   35
            Left            =   9840
            TabIndex        =   175
            Top             =   4080
            Width           =   660
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono"
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
            Index           =   34
            Left            =   7440
            TabIndex        =   174
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label lblHoraAtencion 
            Height          =   135
            Left            =   1560
            TabIndex        =   173
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Forma Pago"
            Height          =   255
            Left            =   12960
            TabIndex        =   172
            Top             =   2610
            Width           =   855
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Trabajo"
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
            Index           =   12
            Left            =   3480
            TabIndex        =   167
            Top             =   285
            Width           =   1245
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
            Left            =   6120
            TabIndex        =   149
            Top             =   765
            Width           =   1215
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
            Left            =   11040
            TabIndex        =   145
            Top             =   960
            Width           =   1575
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
            Left            =   9960
            TabIndex        =   144
            Top             =   960
            Width           =   1095
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
            Left            =   3840
            TabIndex        =   142
            Top             =   765
            Width           =   1215
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
            Left            =   3720
            TabIndex        =   141
            Top             =   1020
            Width           =   2175
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Presupuesto"
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
            Left            =   9720
            TabIndex        =   137
            Top             =   285
            Width           =   1215
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
            Left            =   11040
            TabIndex        =   136
            Top             =   240
            Width           =   1815
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
            TabIndex        =   124
            Top             =   2610
            Width           =   900
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
            TabIndex        =   123
            Top             =   285
            Width           =   735
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
            Left            =   6960
            TabIndex        =   122
            Top             =   285
            Width           =   990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   195
            X2              =   12840
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   4
            X1              =   135
            X2              =   12600
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   120
            X2              =   12840
            Y1              =   720
            Y2              =   720
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
            Left            =   8160
            TabIndex        =   110
            Top             =   2610
            Width           =   1050
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
            Left            =   5175
            TabIndex        =   106
            Top             =   2100
            Width           =   2790
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   210
            X2              =   12720
            Y1              =   3345
            Y2              =   3345
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
            Left            =   7800
            TabIndex        =   66
            Top             =   2985
            Width           =   495
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
            Left            =   8400
            TabIndex        =   65
            Top             =   2955
            Width           =   2010
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario/Conductor"
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
            Left            =   2640
            TabIndex        =   64
            Top             =   4080
            Width           =   1860
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
            Left            =   8880
            TabIndex        =   63
            Top             =   2115
            Width           =   315
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
            Left            =   9360
            TabIndex        =   62
            Top             =   2085
            Width           =   3420
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
            TabIndex        =   51
            Top             =   2130
            Width           =   570
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
            TabIndex        =   47
            Top             =   3585
            Width           =   1350
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
            TabIndex        =   46
            Top             =   3555
            Width           =   780
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
            TabIndex        =   45
            Top             =   2955
            Width           =   5880
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
            Left            =   9420
            TabIndex        =   44
            Top             =   1695
            Width           =   2880
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
            Left            =   4035
            TabIndex        =   43
            Top             =   1695
            Width           =   3540
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
            TabIndex        =   42
            Top             =   1695
            Width           =   1980
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
            TabIndex        =   41
            Top             =   1020
            Width           =   525
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
            TabIndex        =   40
            Top             =   1695
            Width           =   510
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
            Left            =   3405
            TabIndex        =   39
            Top             =   1695
            Width           =   600
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
            Left            =   7890
            TabIndex        =   38
            Top             =   1695
            Width           =   330
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
            Left            =   8895
            TabIndex        =   37
            Top             =   1725
            Width           =   480
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
            TabIndex        =   36
            Top             =   2955
            Width           =   600
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
            Left            =   3555
            TabIndex        =   35
            Top             =   2610
            Width           =   1215
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
            TabIndex        =   34
            Top             =   3540
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F. Entrega"
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
            TabIndex        =   33
            Top             =   4050
            Width           =   870
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
            Left            =   11520
            TabIndex        =   32
            Top             =   2640
            Visible         =   0   'False
            Width           =   1125
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
            Left            =   4260
            TabIndex        =   31
            Top             =   2100
            Width           =   840
         End
         Begin VB.Label lblIdMarca 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1170
            TabIndex        =   49
            Top             =   1695
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblIdModelo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5640
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
            X2              =   12600
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   3
            X1              =   135
            X2              =   12840
            Y1              =   720
            Y2              =   720
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
            TabIndex        =   50
            Top             =   4605
            Visible         =   0   'False
            Width           =   450
         End
      End
      Begin VB.Frame fmeCar 
         Height          =   4905
         Left            =   -74950
         TabIndex        =   84
         Top             =   350
         Width           =   11700
         Begin VB.TextBox txtValorDefCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4920
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   139
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtHorasCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   138
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtSeccion 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1995
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   90
            Top             =   405
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox txtValorFinCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7455
            MaxLength       =   8
            TabIndex        =   87
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtPorcDesCar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5955
            TabIndex        =   85
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
            TabIndex        =   86
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin MSDataListLib.DataCombo dtcCargoCar 
            Bindings        =   "frmRecepcion.frx":4133
            Height          =   315
            Left            =   8460
            TabIndex        =   88
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
            Bindings        =   "frmRecepcion.frx":414D
            Height          =   315
            Left            =   9720
            TabIndex        =   89
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
            TabIndex        =   91
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
            Bindings        =   "frmRecepcion.frx":4167
            Height          =   315
            Left            =   2370
            TabIndex        =   92
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
            Bindings        =   "frmRecepcion.frx":4185
            Height          =   315
            Left            =   60
            TabIndex        =   93
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
            TabIndex        =   115
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
            TabIndex        =   150
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
            Caption         =   "Concepto"
            Height          =   195
            Index           =   24
            Left            =   720
            TabIndex        =   103
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
            TabIndex        =   102
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
            TabIndex        =   101
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
            TabIndex        =   100
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
            TabIndex        =   99
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
            TabIndex        =   98
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   18
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
               Object.Width           =   1499
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
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "FechaAsigna"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   17
               Text            =   "Horas Asignadas"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbAddServicioOtr 
            Height          =   330
            Left            =   105
            TabIndex        =   116
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
            TabIndex        =   117
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
      Begin VB.Frame fmeRep 
         Height          =   4800
         Left            =   50
         TabIndex        =   73
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
            TabIndex        =   147
            Top             =   4400
            Width           =   1695
         End
         Begin MSComctlLib.ListView lvwRepuestos 
            Height          =   4065
            Left            =   0
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
            TabIndex        =   118
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
            TabIndex        =   152
            Top             =   4800
            Visible         =   0   'False
            Width           =   1815
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
            TabIndex        =   135
            Top             =   4800
            Visible         =   0   'False
            Width           =   1650
         End
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
            TabIndex        =   132
            Top             =   4785
            Visible         =   0   'False
            Width           =   1890
         End
         Begin MSComctlLib.Toolbar tlbAgregarRepuestos 
            Height          =   330
            Left            =   120
            TabIndex        =   129
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
            TabIndex        =   114
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
            TabIndex        =   131
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
            TabIndex        =   130
            Top             =   2565
            Width           =   1980
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
         Height          =   6015
         Left            =   -70440
         TabIndex        =   54
         Top             =   350
         Width           =   6885
         Begin VB.TextBox txtNroCupon 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            MaxLength       =   4
            TabIndex        =   160
            Top             =   5400
            Width           =   1335
         End
         Begin VB.ComboBox cmbCuponera 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmRecepcion.frx":41A0
            Left            =   240
            List            =   "frmRecepcion.frx":41B0
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   5400
            Width           =   2055
         End
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
            Height          =   2415
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   6570
         End
         Begin MSDataListLib.DataCombo dtcPromocion 
            Bindings        =   "frmRecepcion.frx":41EE
            Height          =   315
            Left            =   5040
            TabIndex        =   165
            Top             =   6240
            Width           =   3615
            _ExtentX        =   6376
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
         Begin MSAdodcLib.Adodc datPromocion 
            Height          =   330
            Left            =   5520
            Top             =   6360
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
         Begin MSComctlLib.ListView lvwCampana 
            Height          =   2055
            Left            =   150
            TabIndex        =   168
            Top             =   3000
            Width           =   6570
            _ExtentX        =   11589
            _ExtentY        =   3625
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
               Object.Width           =   9790
            EndProperty
         End
         Begin VB.Label Label11 
            Caption         =   "Campañas:"
            Height          =   255
            Left            =   240
            TabIndex        =   170
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Campaña"
            Height          =   375
            Left            =   4200
            TabIndex        =   164
            Top             =   6240
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Nro Cupón"
            Height          =   255
            Left            =   2760
            TabIndex        =   159
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Dscto. Cuponera"
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
            TabIndex        =   158
            Top             =   5160
            Width           =   1455
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
         Height          =   6015
         Left            =   -74835
         TabIndex        =   52
         Top             =   350
         Width           =   4050
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir Inventario"
            Height          =   375
            Left            =   2160
            TabIndex        =   169
            Top             =   5520
            Width           =   1455
         End
         Begin VB.ComboBox cmbBencina 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmRecepcion.frx":4209
            Left            =   1200
            List            =   "frmRecepcion.frx":421C
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   5160
            Width           =   2535
         End
         Begin MSComctlLib.ListView lvwInventario 
            Height          =   4815
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3675
            _ExtentX        =   6482
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
            TabIndex        =   154
            Top             =   5160
            Width           =   855
         End
      End
      Begin VB.Frame fmeCia 
         Height          =   1545
         Left            =   -75000
         TabIndex        =   10
         Top             =   4785
         Width           =   12900
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
            Left            =   11160
            TabIndex        =   22
            Top             =   720
            Width           =   1620
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
            TabIndex        =   15
            Top             =   765
            Width           =   5400
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
               TabIndex        =   17
               Top             =   240
               Width           =   1920
            End
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
               TabIndex        =   16
               Top             =   240
               Width           =   1920
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Soles"
               Height          =   195
               Index           =   19
               Left            =   2730
               TabIndex        =   20
               Top             =   270
               Width           =   390
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dólares"
               Height          =   195
               Index           =   20
               Left            =   105
               TabIndex        =   18
               Top             =   270
               Width           =   540
            End
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
            TabIndex        =   24
            Top             =   1125
            Width           =   5460
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
            TabIndex        =   21
            Top             =   750
            Width           =   1500
         End
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
            TabIndex        =   19
            Top             =   330
            Width           =   2925
         End
         Begin MSComctlLib.Toolbar tlbCiaSeg 
            Height          =   330
            Left            =   5085
            TabIndex        =   105
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
            Left            =   9600
            TabIndex        =   146
            Top             =   765
            Width           =   1575
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
            TabIndex        =   23
            Top             =   420
            Width           =   4890
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
            TabIndex        =   14
            Top             =   405
            Width           =   1020
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
            TabIndex        =   13
            Top             =   825
            Width           =   765
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
            TabIndex        =   12
            Top             =   1230
            Width           =   885
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
            TabIndex        =   11
            Top             =   225
            Width           =   1815
         End
      End
      Begin MSComctlLib.StatusBar stbTotalRepuestos 
         Height          =   405
         Left            =   6720
         TabIndex        =   76
         Top             =   5640
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
         TabIndex        =   104
         Top             =   5160
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
               Text            =   "T. Mat. procesivos"
               TextSave        =   "T. Mat. procesivos"
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
         TabIndex        =   113
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
         TabIndex        =   126
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
         TabIndex        =   128
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
      TabIndex        =   70
      Top             =   7440
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
      TabIndex        =   151
      Top             =   7440
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
   Begin Crystal.CrystalReport rptOTS 
      Left            =   9120
      Top             =   7560
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
   Begin Crystal.CrystalReport crInventario 
      Left            =   10320
      Top             =   7680
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
   Begin Crystal.CrystalReport rptOTA 
      Left            =   10680
      Top             =   8160
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
   Begin MSAdodcLib.Adodc datTipoVenta 
      Height          =   330
      Left            =   11760
      Top             =   7080
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
   Begin VB.Image Image1 
      Height          =   1830
      Left            =   240
      Picture         =   "frmRecepcion.frx":4257
      Top             =   8400
      Visible         =   0   'False
      Width           =   8745
   End
End
Attribute VB_Name = "frmRecepcion"
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
Dim mstrProcedencia As String
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
Dim mblnOtFacturada As Boolean
Dim curSumaInsumos As Currency
Dim mstrProcedenciaAux As String
Public ConfirmarImprimirInventarioVehiculo As String

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
        
    End With
Case "CS"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INA"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INR"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INS"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INU"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INW"
    With Me
        .lblPat.Caption = "V.I.N."
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "NGN"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "INC"
    With Me
        .lblPat.Caption = "V.I.N."
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "PEX"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "REN"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = True
        .optReparacion.Visible = True
        
        .tlbAgregarRepuestos.Visible = True
    End With
Case "PRE"
    With Me
        .lblPat.Caption = gstrNombrePatente
        .txtFolioGarantia = "S/F"
        .txtFolioGarantia.Enabled = False
        .optMantencion.Visible = False
        .optReparacion.Visible = False
        .cmdAnularReserva.Visible = False
        .cmdReserva.Visible = False
        .tlbAgregarRepuestos.Visible = False
        mstrEstadoPresupuesto = "ON"
        mstrLiquidaPresupuesto = False
        gcurInsumo = 0
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
If pstrEstado = "V" Or pstrEstado = "R" Or pstrEstado = "P" Then
    fmePat.Enabled = True
    fmeCia.Enabled = True
    fmeInv.Enabled = True
    fmeCom.Enabled = True
    mblnBloqueo = False
ElseIf pstrEstado = "B" Or pstrEstado = "F" Then
    If mblnOtFacturada = True Then
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
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_MECANICA_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssOtr Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_OTRO_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssCar Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN  FROM TLLR_CARROCERIA_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssTer Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_TERCEROS_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssRep Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_REPUESTOS_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' "
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    End If
Else
    If Seccion = ssMec Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_MECANICA_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssOtr Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_OTRO_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssCar Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN  FROM TLLR_CARROCERIA_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssTer Then
        gstrSql = "SELECT SUM(SUBTOTAL) AS RESUMEN FROM TLLR_TERCEROS_OT"
        gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
        gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
        gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    ElseIf Seccion = ssRep Then
        gstrSql = "SELECT SUM(SUBTOTAL)  AS RESUMEN FROM TLLR_REPUESTOS_OT"
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
Dim SumaInsumos As Currency
                            
                            
If pstrIdTipoCargo = "" Then
    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "'"
    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoLubricantes & "'" '90'
Else
    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoLubricantes & "'" '90'
End If
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
If pstrIdTipoCargo = "" Then
    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "'"
    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoMateriales & "'" '85'
Else
    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoMateriales & "'" '85'

End If
SumaMateriales = 0

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


'///// Insumos
If pstrIdTipoCargo = "" Then
    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "'"
    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoInsumos & "'" '85'
Else
    gstrSql = "SELECT TLLR_REPUESTOS_OT.SUBTOTAL,"
    gstrSql = gstrSql & " Stck_Item.ID_FAMILIA  FROM TLLR_REPUESTOS_OT"
    gstrSql = gstrSql & " INNER JOIN STCK_ITEM ON STCK_ITEM.ID_ITEM = TLLR_REPUESTOS_OT.ID_ITEM"
    gstrSql = gstrSql & " WHERE  ID_OT = '" & pstrIdOT & "' AND ID_TIPO_CARGO = '" & pstrIdTipoCargo & "'"
    gstrSql = gstrSql & " AND ID_EMPRESA = '" & pstrIdEmpresa & "' AND ID_SUCURSAL='" & pstrIdSucursal & "'"
    gstrSql = gstrSql & " AND SECCION_OT = '" & pstrTipoOt & "'"
    gstrSql = gstrSql & " AND STCK_ITEM.ID_FAMILIA = '" & gstrCodigoInsumos & "'" '85'

End If
SumaInsumos = 0

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        While Not .EOF
            SumaInsumos = SumaInsumos + !SubTotal
            .MoveNext
        Wend
        .Close
    End With
End If
curSumaInsumos = SumaInsumos


End Function

Function AccesoEliminar(itmSeleccionado As ListItem) As Boolean
'If itmSeleccionado.SubItems(5) = "85" Then
    AccesoEliminar = True
'Else
'    AccesoEliminar = False
'End If
End Function


Public Function nroMovilXId(cod As String) As String

Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "Select  isnull(movil,'') as NroMovil From Tllr_Mecanicos Where Id_Mecanico =  '" & cod & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        nroMovilXId = recAux!NroMovil
    End If
End If
Conexion.CloseHost recAux


End Function
Sub ImprimirDocumentoASP()

Dim DbsnuevaOtA As Database
Dim TablaOtA As DAO.Recordset
Dim GcamBaseTemOtA As String

Dim rcOtA As Long
Dim WinPathOtA As String
    WinPathOtA = Space$(300)
    rcOtA = GetWindowsDirectory(WinPathOtA, 300)
    GcamBaseTemOtA = Trim$(WinPathOtA)
    GcamBaseTemOtA = Mid(GcamBaseTemOtA, 1, Len(GcamBaseTemOtA) - 1) & "\Temp"
    
    
Dim wrkPredeterminadoOtA As Workspace
Set wrkPredeterminadoOtA = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
If Dir(gstrPathReporte & "\BDNuevaOtA.mdb") <> "" Then Kill gstrPathReporte & "\BDNuevaOtA.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
Set DbsnuevaOtA = wrkPredeterminadoOtA.CreateDatabase(gstrPathReporte & "\BDNuevaOtA.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    
DbsnuevaOtA.Execute "CREATE TABLE Tllr_OT (Id_Empresa text,   Id_Sucursal text,   Id_OT text,   Seccion_OT text,   Patente text,   Fecha_Emision text,   Nro_Siniestro text,   Nro_Poliza text,   Liquidador text,   Comentario memo,   Fecha_Liquidacion text,   Kilometros_Recepcion text,   Id_Compañia_Seguro text,   Id_Presupuesto text, Entrega text, FormaPago text ,Solicitado text, Cono text,ComentarioAux memo, TelefonoOT text, CorreoOT text )"
DbsnuevaOtA.Execute "CREATE TABLE Tllr_Vehiculo_Cliente (   Id_Marca text,   Año text,   Nro_Motor text,   VIN text, Fecha_Venta text  )"
DbsnuevaOtA.Execute "CREATE TABLE Tllr_Mecanicos (   Nombre text, Movil text,E_Mail text )"
DbsnuevaOtA.Execute "CREATE TABLE Glbl_Marca (   Descripcion text  )"
DbsnuevaOtA.Execute "CREATE TABLE Glbl_Modelo (   Descripcion text ,CodigoModeloMarca text,CombDescripcion text  )"
DbsnuevaOtA.Execute "CREATE TABLE Glbl_Cliente_Proveedor (   Telefono text,   Rut text,Razon_Social text,Direccion text, Email text  )"
DbsnuevaOtA.Execute "CREATE TABLE Glbl_Color_Exterior (   Descripcion text  )"
DbsnuevaOtA.Execute "CREATE TABLE Tllr_Parametro (   NotaRecepcion text  )"
    
     
     Dim gadoPrincipalOtA As New ADODB.Recordset
     Dim gstrSqlOtA As String
'
'     gstrSqlOtA = " Select Id_Empresa ,Id_Sucursal ,Id_OT ,Seccion_OT ,Patente ,Fecha_Emision ,Nro_Siniestro , Nro_Poliza , Liquidador , Comentario ,Fecha_Liquidacion , Kilometros_Recepcion ,Id_Compañia_Seguro, "
'     gstrSqlOtA = gstrSqlOtA & "Id_Presupuesto  ,VC.Año,VC.Id_Marca,VC.Nro_Motor,VC.VIN,Marca.Descripcion as MarcaDescripcion ,Modelo.Descripcion as ModeloDescripcion, Modelo.CodigoModeloMarca, Modelo.CombDescripcion, Mec.Nombre as MecNombre, Mec.Movil as MecMovil, Mec.E_Mail as MecEmail"
'     gstrSqlOtA = gstrSqlOtA & " ,CP.Rut ,CP.Telefono,CP.Razon_Social,CP.Direccion,CE.Descripcion as CEDescripcion, Entrega_Estimada,Id_Tipo_Venta,Solicitado_Por From Tllr_OT  "
'     gstrSqlOtA = gstrSqlOtA & "Outer Apply(Select Id_Marca,Id_Modelo,Año,Nro_Motor,VIN, Id_Cliente_Proveedor,Id_Color_Exterior From Tllr_Vehiculo_Cliente Where Patente = Tllr_OT.Patente) VC "
'     gstrSqlOtA = gstrSqlOtA & "Outer Apply(Select Descripcion From Glbl_Marca Where Glbl_Marca.Id_Marca=Vc.Id_Marca) Marca "
'     gstrSqlOtA = gstrSqlOtA & " Outer Apply(Select m.Descripcion,m.CodigoModeloMarca,c.Descripcion as CombDescripcion From Glbl_Modelo m Left Join Glbl_Combustible c on (m.Id_Combustible=c.Id_Combustible)  Where Id_Marca =Vc.Id_Marca and Id_Modelo= Vc.Id_Modelo) Modelo "
'     gstrSqlOtA = gstrSqlOtA & "Outer Apply(Select Telefono,Rut,Razon_Social,Direccion,E_Mail From Glbl_Cliente_Proveedor Where Id_Cliente_Proveedor = Vc.Id_Cliente_Proveedor) CP "
'     gstrSqlOtA = gstrSqlOtA & "Outer Apply(Select Descripcion From Glbl_Color_Exterior Where Id_Color_Exterior = VC.Id_Color_Exterior) CE "
'     gstrSqlOtA = gstrSqlOtA & "Outer Apply(Select Nombre, Movil,E_Mail From Tllr_Mecanicos Where Id_Mecanico = Tllr_OT.RealizadoPor) Mec "
    gstrSqlOtA = " Select Id_Empresa ,Id_Sucursal ,Id_OT ,Seccion_OT ,Patente ,Fecha_Emision ,Nro_Siniestro , Nro_Poliza , Liquidador , Comentario ,Fecha_Liquidacion , "
    gstrSqlOtA = gstrSqlOtA & " Kilometros_Recepcion ,Id_Compañia_Seguro, Id_Presupuesto  ,VC.Año,VC.Id_Marca,VC.Nro_Motor,VC.VIN,Marca.Descripcion as MarcaDescripcion ,Modelo.Descripcion as ModeloDescripcion,"
    gstrSqlOtA = gstrSqlOtA & " Modelo.CodigoModeloMarca, Modelo.CombDescripcion, Mec.Nombre as MecNombre, Mec.Movil as MecMovil, Mec.E_Mail as MecEmail ,"
    gstrSqlOtA = gstrSqlOtA & " CP.Rut ,CP.Telefono,CP.Razon_Social,CP.Direccion,CP.E_Mail as CliEmail, CE.Descripcion as  CEDescripcion, Entrega_Estimada,FormaPago.Descripcion as FormaPago,Solicitado_Por ,VC.Fecha_Venta as FechaVenta , Cono.Color as Cono, Tllr_OT.Telefono as TelefonoOT, Tllr_OT.Correo as CorreoOT"
    gstrSqlOtA = gstrSqlOtA & " From Tllr_OT"
    gstrSqlOtA = gstrSqlOtA & " Outer Apply(Select Id_Marca,Id_Modelo,Año,Nro_Motor,VIN, Id_Cliente_Proveedor,Id_Color_Exterior, Fecha_Venta From Tllr_Vehiculo_Cliente Where Patente = Tllr_OT.Patente) VC"
    gstrSqlOtA = gstrSqlOtA & " Outer Apply(Select Descripcion From Glbl_Marca Where Glbl_Marca.Id_Marca=Vc.Id_Marca) Marca"
    gstrSqlOtA = gstrSqlOtA & " Outer Apply(Select m.Descripcion,m.CodigoModeloMarca,c.Descripcion as CombDescripcion From Glbl_Modelo m"
    gstrSqlOtA = gstrSqlOtA & " Left Join Glbl_Combustible c on (m.Id_Combustible=c.Id_Combustible)   Where Id_Marca =Vc.Id_Marca and Id_Modelo= Vc.Id_Modelo) Modelo"
    gstrSqlOtA = gstrSqlOtA & " Outer Apply(Select Telefono,Rut,Razon_Social,Direccion,E_Mail  From Glbl_Cliente_Proveedor Where Id_Cliente_Proveedor = Vc.Id_Cliente_Proveedor) CP"
    gstrSqlOtA = gstrSqlOtA & " Outer Apply(Select Descripcion From Glbl_Color_Exterior Where Id_Color_Exterior = VC.Id_Color_Exterior) CE"
    gstrSqlOtA = gstrSqlOtA & " Outer Apply(Select Nombre, Movil,E_Mail From Tllr_Mecanicos Where Id_Mecanico = Tllr_OT.RealizadoPor) Mec"
    gstrSqlOtA = gstrSqlOtA & " Outer apply (Select Descripcion from Glbl_Tipo_Venta where Glbl_Tipo_Venta.Id_Tipo_Venta=Tllr_OT.Id_Tipo_Venta) FormaPago"
    gstrSqlOtA = gstrSqlOtA & " Outer apply (Select Color,Descripcion from Tllr_Tipo_Cono where Tllr_Tipo_Cono.Id_Tipo_Cono =Tllr_OT.Id_Tipo_Cono  ) Cono"
     gstrSqlOtA = gstrSqlOtA & " Where Tllr_OT.Id_Empresa= '" & gstrIdEmpresa & "'"
     gstrSqlOtA = gstrSqlOtA & " And Tllr_OT.Id_Sucursal= '" & gstrIdSucursal & "'"
     gstrSqlOtA = gstrSqlOtA & " And Tllr_OT.Id_OT= '" & lblNroRecepcion & "'"
     gstrSqlOtA = gstrSqlOtA & " And Tllr_OT.Seccion_OT= '" & gstrSeccion & "'"
     
     
     Dim Id_EmpresaBdS As String
     Dim Id_SucursalBdS As String
     Dim Id_OTBdS As String
     Dim Seccion_OTBdS As String
     Dim PatenteBdS As String
     Dim Fecha_EmisionBdS As String
     Dim Nro_SiniestroBdS As String
     Dim Nro_PolizaBdS As String
     Dim LiquidadorBdS As String
     Dim ComentarioBdS As String
     Dim Fecha_LiquidacionBdS As String
     Dim Kilometros_RecepcionBdS As String
     Dim Id_Compañia_SeguroBdS As String
     Dim Id_PresupuestoBdS As String
     Dim AñoBdS As String
     Dim Id_MarcaBdS As String
     Dim Nro_MotorBdS As String
     Dim VINBdS As String
     Dim MarcaDescripcionBdS As String
     Dim ModeloDescripcionBdS As String
     Dim CodigoModeloMarcaBdS As String
     Dim CombDescripcionBdS As String
     Dim RutBdS As String
     Dim TelefonoBdS As String
     Dim Razon_SocialBdS As String
     Dim DireccionBdS As String
     Dim CEDescripcionBdS As String
     Dim EntregaEstimadaBdS As String
     Dim FormaPagoBdS As String
     Dim SolicitadoBdS As String
     Dim CliEmailBdS As String
     Dim FechaVentaBdS As String
     Dim ConoBdS As String
     
     Dim MecNombreBdS As String
     Dim MecMovilBdS As String
     Dim MecEmailBdS As String
     
     Dim TelefonoOT As String
     Dim CorreoOT As String
     
     
     
     If Conexion.SendHost(gstrSqlOtA, gadoPrincipalOtA, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With gadoPrincipalOtA
        If Not .BOF And Not .EOF Then
             .MoveFirst
             While Not .EOF
                 Id_EmpresaBdS = ValorNulo(!Id_Empresa)
                 Id_SucursalBdS = ValorNulo(!Id_Sucursal)
                 Id_OTBdS = ValorNulo(!Id_OT)
                 Seccion_OTBdS = ValorNulo(!Seccion_OT)
                 PatenteBdS = ValorNulo(!Patente)
                 Fecha_EmisionBdS = ValorNulo(!Fecha_Emision)
                 Nro_SiniestroBdS = IIf(IsNull(!Nro_Siniestro), "", !Nro_Siniestro)
                 Nro_PolizaBdS = IIf(IsNull(!Nro_Poliza), "", !Nro_Poliza)
                 
                 LiquidadorBdS = IIf(IsNull(!Liquidador), "", !Liquidador)
                 
                 ComentarioBdS = IIf(IsNull(!Comentario), "", !Comentario)
                 ComentarioBdS = LTrim(RTrim(ComentarioBdS))
                 Fecha_LiquidacionBdS = IIf(IsNull(!Fecha_Liquidacion), "", !Fecha_Liquidacion)
                 Kilometros_RecepcionBdS = IIf(IsNull(!Kilometros_Recepcion), "", !Kilometros_Recepcion)
                 Id_Compañia_SeguroBdS = IIf(IsNull(!Id_Compañia_Seguro), "", !Id_Compañia_Seguro)
                 Id_PresupuestoBdS = IIf(IsNull(!Id_Presupuesto), "", !Id_Presupuesto)
                 AñoBdS = IIf(IsNull(!Año), "", !Año)
                 Id_MarcaBdS = IIf(IsNull(!Id_Marca), "", !Id_Marca)
                 Nro_MotorBdS = IIf(IsNull(!Nro_Motor), "", !Nro_Motor)
                 VINBdS = IIf(IsNull(!VIN), "", !VIN)
                 MarcaDescripcionBdS = ValorNulo(!MarcaDescripcion)
                 ModeloDescripcionBdS = ValorNulo(!ModeloDescripcion)
                 CodigoModeloMarcaBdS = ValorNulo(!CodigoModeloMarca)
                 CombDescripcionBdS = IIf(IsNull(!CombDescripcion), "", !CombDescripcion)
                 RutBdS = ValorNulo(!rut)
                 TelefonoBdS = IIf(IsNull(!Telefono), "", !Telefono)
                 Razon_SocialBdS = IIf(IsNull(!Razon_Social), "", !Razon_Social)
                 DireccionBdS = IIf(IsNull(!Direccion), "", !Direccion)
                 CEDescripcionBdS = IIf(IsNull(!CEDescripcion), "", !CEDescripcion)
                 EntregaEstimadaBdS = IIf(IsNull(!Entrega_Estimada), "", !Entrega_Estimada)
                 FormaPagoBdS = IIf(IsNull(!FormaPago), "", !FormaPago)
                 SolicitadoBdS = IIf(IsNull(!Solicitado_Por), "", !Solicitado_Por)
                 FechaVentaBdS = IIf(IsNull(!FechaVenta), "", !FechaVenta)
                 ConoBdS = IIf(IsNull(!Cono), "", !Cono)
                 
                 CliEmailBdS = IIf(IsNull(!CliEmail), "", !CliEmail)
                 
                 MecNombreBdS = IIf(IsNull(!MecNombre), "", !MecNombre)
                 MecMovilBdS = IIf(IsNull(!MecMovil), "", !MecMovil)
                 MecEmailBdS = IIf(IsNull(!MecEmail), "", !MecEmail)
                 
                 TelefonoOT = ValorNulo(!TelefonoOT)
                 CorreoOT = ValorNulo(!CorreoOT)

                .MoveNext
              Wend

        End If
        End With
     End If
    Conexion.CloseHost gadoPrincipalOtA
     
    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Tllr_OT")
    TablaOtA.AddNew
    

    TablaOtA!Id_Empresa = Id_EmpresaBdS
    TablaOtA!Id_Sucursal = Id_SucursalBdS
    TablaOtA!Id_OT = Id_OTBdS
    TablaOtA!Seccion_OT = Seccion_OTBdS
    TablaOtA!Patente = PatenteBdS
    TablaOtA!Fecha_Emision = Fecha_EmisionBdS
    TablaOtA!Nro_Siniestro = Nro_SiniestroBdS
    TablaOtA!Nro_Poliza = Nro_PolizaBdS
    TablaOtA!Liquidador = LiquidadorBdS
    ComentarioBdS = LTrim(RTrim(ComentarioBdS))
    TablaOtA!Comentario = ComentarioBdS
    TablaOtA!ComentarioAux = ComentarioBdS
    TablaOtA!Fecha_Liquidacion = Fecha_LiquidacionBdS
    TablaOtA!Kilometros_Recepcion = Kilometros_RecepcionBdS
    TablaOtA!Id_Compañia_Seguro = Id_Compañia_SeguroBdS
    TablaOtA!Id_Presupuesto = Id_PresupuestoBdS
    TablaOtA!FormaPago = FormaPagoBdS
    TablaOtA!Entrega = EntregaEstimadaBdS
    TablaOtA!Solicitado = SolicitadoBdS
    TablaOtA!Cono = ConoBdS
    TablaOtA!TelefonoOT = TelefonoOT
    TablaOtA!CorreoOT = CorreoOT
  
    TablaOtA.Update
    TablaOtA.Close


    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Tllr_Vehiculo_Cliente")
    TablaOtA.AddNew
    TablaOtA!Id_Marca = Id_MarcaBdS
    TablaOtA!Año = AñoBdS
    TablaOtA!Nro_Motor = Nro_MotorBdS
    TablaOtA!VIN = VINBdS
    TablaOtA!Fecha_Venta = FechaVentaBdS
    TablaOtA.Update
    TablaOtA.Close
    
    
    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Tllr_Mecanicos")
    TablaOtA.AddNew
    TablaOtA!Nombre = MecNombreBdS
    TablaOtA!Movil = MecMovilBdS
    TablaOtA!E_Mail = MecEmailBdS
    TablaOtA.Update
    TablaOtA.Close
    
    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Glbl_Marca")
    TablaOtA.AddNew
    TablaOtA!Descripcion = MarcaDescripcionBdS
    TablaOtA.Update
    TablaOtA.Close
    
    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Glbl_Modelo")
    TablaOtA.AddNew
    TablaOtA!Descripcion = ModeloDescripcionBdS
    TablaOtA!CodigoModeloMarca = CodigoModeloMarcaBdS
    TablaOtA!CombDescripcion = CombDescripcionBdS
    TablaOtA.Update
    TablaOtA.Close
    
        
    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Glbl_Cliente_Proveedor")
    TablaOtA.AddNew
    TablaOtA!Telefono = TelefonoBdS
    TablaOtA!rut = RutBdS
    TablaOtA!Razon_Social = Razon_SocialBdS
    TablaOtA!Direccion = DireccionBdS
    TablaOtA!email = CliEmailBdS
    TablaOtA.Update
    TablaOtA.Close
    
    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Glbl_Color_Exterior")
    TablaOtA.AddNew
    TablaOtA!Descripcion = CEDescripcionBdS
    TablaOtA.Update
    TablaOtA.Close
    
    Set TablaOtA = DbsnuevaOtA.OpenRecordset("SELECT * FROM Tllr_Parametro")
    TablaOtA.AddNew
    TablaOtA!NotaRecepcion = "." 'no se muestra en el reporte
    TablaOtA.Update
    TablaOtA.Close
    
    
    DbsnuevaOtA.Close
      With rptOTA
      
                              
            Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
            Me.cdImpresora.CancelError = True
            Me.cdImpresora.Action = 5
                                             
            .CopiesToPrinter = cdImpresora.Copies
            .ReportFileName = gstrPathReporte & "\OT_NewVistaPrevia1.rpt"
'            .Destination = crptToWindow
            .Destination = crptToPrinter
            .WindowState = crptMaximized
            .DataFiles(0) = gstrPathReporte & "\BDNuevaOtA.mdb"
            '.Formulas(1) = "Comentario= ''"
            '.Formulas(0) = "comentario= ''"
            
            .Action = True
                
     End With
   

End Sub




Sub ImprimirDocumentoSMP()
 
    
    Dim DbsnuevaOtS As Database
    Dim TablaOtS As DAO.Recordset
    'Dim i As Integer
    Dim GcamBaseTemOtS As String
    
    Dim rcOtS As Long
    Dim WinPathOtS As String
    WinPathOtS = Space$(300)
    rcOtS = GetWindowsDirectory(WinPathOtS, 300)
    GcamBaseTemOtS = Trim$(WinPathOtS)
    GcamBaseTemOtS = Mid(GcamBaseTemOtS, 1, Len(GcamBaseTemOtS) - 1) & "\Temp"
    
    
    Dim wrkPredeterminadoOtS As Workspace
    Set wrkPredeterminadoOtS = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(gstrPathReporte & "\BDNuevaOtS.mdb") <> "" Then Kill gstrPathReporte & "\BDNuevaOtS.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set DbsnuevaOtS = wrkPredeterminadoOtS.CreateDatabase(gstrPathReporte & "\BDNuevaOtS.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    
    DbsnuevaOtS.Execute "CREATE TABLE Tllr_OT (Id_Empresa text,   Id_Sucursal text,   Id_OT text,   Seccion_OT text,   Patente text,   Fecha_Emision text,   Nro_Siniestro text,   Nro_Poliza text,   Liquidador text,   Comentario text,   Fecha_Liquidacion text,   Kilometros_Recepcion text,   Id_Compañia_Seguro text,   Id_Presupuesto text  )"
    DbsnuevaOtS.Execute "CREATE TABLE Tllr_Vehiculo_Cliente (   Id_Marca text,   Año text,   Nro_Motor text,   VIN text  )"
    DbsnuevaOtS.Execute "CREATE TABLE Tllr_Mecanicos (   Nombre text, Movil text,E_Mail text )"
    DbsnuevaOtS.Execute "CREATE TABLE Glbl_Marca (   Descripcion text  )"
    DbsnuevaOtS.Execute "CREATE TABLE Glbl_Modelo (   Descripcion text ,CodigoModeloMarca text,CombDescripcion text  )"
    DbsnuevaOtS.Execute "CREATE TABLE Glbl_Cliente_Proveedor (   Telefono text,   Rut text,Razon_Social text,Direccion text  )"
    DbsnuevaOtS.Execute "CREATE TABLE Glbl_Color_Exterior (   Descripcion text  )"
    DbsnuevaOtS.Execute "CREATE TABLE Tllr_Parametro (   NotaRecepcion text  )"
    
     
     Dim gadoPrincipalOtS As New ADODB.Recordset
     Dim gstrSqlOtS As String

     gstrSqlOtS = " Select Id_Empresa ,Id_Sucursal ,Id_OT ,Seccion_OT ,Patente ,Fecha_Emision ,Nro_Siniestro , Nro_Poliza , Liquidador , Comentario ,Fecha_Liquidacion , Kilometros_Recepcion ,Id_Compañia_Seguro, "
     gstrSqlOtS = gstrSqlOtS & "Id_Presupuesto  ,VC.Año,VC.Id_Marca,VC.Nro_Motor,VC.VIN,Marca.Descripcion as MarcaDescripcion ,Modelo.Descripcion as ModeloDescripcion, Modelo.CodigoModeloMarca, Modelo.CombDescripcion, Mec.Nombre as MecNombre, Mec.Movil as MecMovil, Mec.E_Mail as MecEmail"
     gstrSqlOtS = gstrSqlOtS & " ,CP.Rut ,CP.Telefono,CP.Razon_Social,CP.Direccion,CE.Descripcion as CEDescripcion From Tllr_OT  "
     gstrSqlOtS = gstrSqlOtS & "Outer Apply(Select Id_Marca,Id_Modelo,Año,Nro_Motor,VIN, Id_Cliente_Proveedor,Id_Color_Exterior From Tllr_Vehiculo_Cliente Where Patente = Tllr_OT.Patente) VC "
     gstrSqlOtS = gstrSqlOtS & "Outer Apply(Select Descripcion From Glbl_Marca Where Glbl_Marca.Id_Marca=Vc.Id_Marca) Marca "
     gstrSqlOtS = gstrSqlOtS & " Outer Apply(Select m.Descripcion,m.CodigoModeloMarca,c.Descripcion as CombDescripcion From Glbl_Modelo m Left Join Glbl_Combustible c on (m.Id_Combustible=c.Id_Combustible)  Where Id_Marca =Vc.Id_Marca and Id_Modelo= Vc.Id_Modelo) Modelo "
     gstrSqlOtS = gstrSqlOtS & "Outer Apply(Select Telefono,Rut,Razon_Social,Direccion From Glbl_Cliente_Proveedor Where Id_Cliente_Proveedor = Vc.Id_Cliente_Proveedor) CP "
     gstrSqlOtS = gstrSqlOtS & "Outer Apply(Select Descripcion From Glbl_Color_Exterior Where Id_Color_Exterior = VC.Id_Color_Exterior) CE "
     gstrSqlOtS = gstrSqlOtS & "Outer Apply(Select Nombre, Movil,E_Mail From Tllr_Mecanicos Where Id_Mecanico = Tllr_OT.RealizadoPor) Mec "
     gstrSqlOtS = gstrSqlOtS & " Where Tllr_OT.Id_Empresa= '" & gstrIdEmpresa & "'"
     gstrSqlOtS = gstrSqlOtS & " And Tllr_OT.Id_Sucursal= '" & gstrIdSucursal & "'"
     gstrSqlOtS = gstrSqlOtS & " And Tllr_OT.Id_OT= '" & lblNroRecepcion & "'"
     gstrSqlOtS = gstrSqlOtS & " And Tllr_OT.Seccion_OT= '" & gstrSeccion & "'"
     
     
     Dim Id_EmpresaBdS As String
     Dim Id_SucursalBdS As String
     Dim Id_OTBdS As String
     Dim Seccion_OTBdS As String
     Dim PatenteBdS As String
     Dim Fecha_EmisionBdS As String
     Dim Nro_SiniestroBdS As String
     Dim Nro_PolizaBdS As String
     Dim LiquidadorBdS As String
     Dim ComentarioBdS As String
     Dim Fecha_LiquidacionBdS As String
     Dim Kilometros_RecepcionBdS As String
     Dim Id_Compañia_SeguroBdS As String
     Dim Id_PresupuestoBdS As String
     Dim AñoBdS As String
     Dim Id_MarcaBdS As String
     Dim Nro_MotorBdS As String
     Dim VINBdS As String
     Dim MarcaDescripcionBdS As String
     Dim ModeloDescripcionBdS As String
     Dim CodigoModeloMarcaBdS As String
     Dim CombDescripcionBdS As String
     Dim RutBdS As String
     Dim TelefonoBdS As String
     Dim Razon_SocialBdS As String
     Dim DireccionBdS As String
     Dim CEDescripcionBdS As String
     
     Dim MecNombreBdS As String
     Dim MecMovilBdS As String
     Dim MecEmailBdS As String
     
     
     If Conexion.SendHost(gstrSqlOtS, gadoPrincipalOtS, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With gadoPrincipalOtS
        If Not .BOF And Not .EOF Then
             .MoveFirst
             While Not .EOF
                 Id_EmpresaBdS = ValorNulo(!Id_Empresa)
                 Id_SucursalBdS = ValorNulo(!Id_Sucursal)
                 Id_OTBdS = ValorNulo(!Id_OT)
                 Seccion_OTBdS = ValorNulo(!Seccion_OT)
                 PatenteBdS = ValorNulo(!Patente)
                 Fecha_EmisionBdS = ValorNulo(!Fecha_Emision)
                 Nro_SiniestroBdS = IIf(IsNull(!Nro_Siniestro), "", !Nro_Siniestro)
                 Nro_PolizaBdS = IIf(IsNull(!Nro_Poliza), "", !Nro_Poliza)
                 
                 LiquidadorBdS = IIf(IsNull(!Liquidador), "", !Liquidador)
                 
                 ComentarioBdS = IIf(IsNull(!Comentario), "", !Comentario)
                 Fecha_LiquidacionBdS = IIf(IsNull(!Fecha_Liquidacion), "", !Fecha_Liquidacion)
                 Kilometros_RecepcionBdS = IIf(IsNull(!Kilometros_Recepcion), "", !Kilometros_Recepcion)
                 Id_Compañia_SeguroBdS = IIf(IsNull(!Id_Compañia_Seguro), "", !Id_Compañia_Seguro)
                 Id_PresupuestoBdS = IIf(IsNull(!Id_Presupuesto), "", !Id_Presupuesto)
                 AñoBdS = IIf(IsNull(!Año), "", !Año)
                 Id_MarcaBdS = IIf(IsNull(!Id_Marca), "", !Id_Marca)
                 Nro_MotorBdS = IIf(IsNull(!Nro_Motor), "", !Nro_Motor)
                 VINBdS = IIf(IsNull(!VIN), "", !VIN)
                 MarcaDescripcionBdS = ValorNulo(!MarcaDescripcion)
                 ModeloDescripcionBdS = ValorNulo(!ModeloDescripcion)
                 CodigoModeloMarcaBdS = ValorNulo(!CodigoModeloMarca)
                 CombDescripcionBdS = IIf(IsNull(!CombDescripcion), "", !CombDescripcion)
                 RutBdS = ValorNulo(!rut)
                 TelefonoBdS = IIf(IsNull(!Telefono), "", !Telefono)
                 Razon_SocialBdS = IIf(IsNull(!Razon_Social), "", !Razon_Social)
                 DireccionBdS = IIf(IsNull(!Direccion), "", !Direccion)
                 CEDescripcionBdS = IIf(IsNull(!CEDescripcion), "", !CEDescripcion)
                 
                 MecNombreBdS = IIf(IsNull(!MecNombre), "", !MecNombre)
                 MecMovilBdS = IIf(IsNull(!MecMovil), "", !MecMovil)
                 MecEmailBdS = IIf(IsNull(!MecEmail), "", !MecEmail)
                 

                .MoveNext
              Wend

        End If
        End With
     End If
    Conexion.CloseHost gadoPrincipalOtS
     
    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Tllr_OT")
    TablaOtS.AddNew
    

    TablaOtS!Id_Empresa = Id_EmpresaBdS
    TablaOtS!Id_Sucursal = Id_SucursalBdS
    TablaOtS!Id_OT = Id_OTBdS
    TablaOtS!Seccion_OT = Seccion_OTBdS
    TablaOtS!Patente = PatenteBdS
    TablaOtS!Fecha_Emision = Fecha_EmisionBdS
    TablaOtS!Nro_Siniestro = Nro_SiniestroBdS
    TablaOtS!Nro_Poliza = Nro_PolizaBdS
    TablaOtS!Liquidador = LiquidadorBdS
    TablaOtS!Comentario = ComentarioBdS
    TablaOtS!Fecha_Liquidacion = Fecha_LiquidacionBdS
    TablaOtS!Kilometros_Recepcion = Kilometros_RecepcionBdS
    TablaOtS!Id_Compañia_Seguro = Id_Compañia_SeguroBdS
    TablaOtS!Id_Presupuesto = Id_PresupuestoBdS
  
    TablaOtS.Update
    TablaOtS.Close


    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Tllr_Vehiculo_Cliente")
    TablaOtS.AddNew
    TablaOtS!Id_Marca = Id_MarcaBdS
    TablaOtS!Año = AñoBdS
    TablaOtS!Nro_Motor = Nro_MotorBdS
    TablaOtS!VIN = VINBdS
    TablaOtS.Update
    TablaOtS.Close
    
    
    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Tllr_Mecanicos")
    TablaOtS.AddNew
    TablaOtS!Nombre = MecNombreBdS
    TablaOtS!Movil = MecMovilBdS
    TablaOtS!E_Mail = MecEmailBdS
    TablaOtS.Update
    TablaOtS.Close
    
    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Glbl_Marca")
    TablaOtS.AddNew
    TablaOtS!Descripcion = MarcaDescripcionBdS
    TablaOtS.Update
    TablaOtS.Close
    
    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Glbl_Modelo")
    TablaOtS.AddNew
    TablaOtS!Descripcion = ModeloDescripcionBdS
    TablaOtS!CodigoModeloMarca = CodigoModeloMarcaBdS
    TablaOtS!CombDescripcion = CombDescripcionBdS
    TablaOtS.Update
    TablaOtS.Close
    
        
    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Glbl_Cliente_Proveedor")
    TablaOtS.AddNew
    TablaOtS!Telefono = TelefonoBdS
    TablaOtS!rut = RutBdS
    TablaOtS!Razon_Social = Razon_SocialBdS
    TablaOtS!Direccion = DireccionBdS
    TablaOtS.Update
    TablaOtS.Close
    
    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Glbl_Color_Exterior")
    TablaOtS.AddNew
    TablaOtS!Descripcion = CEDescripcionBdS
    TablaOtS.Update
    TablaOtS.Close
    
    Set TablaOtS = DbsnuevaOtS.OpenRecordset("SELECT * FROM Tllr_Parametro")
    TablaOtS.AddNew
    TablaOtS!NotaRecepcion = "." 'no se muestra en el reporte
    TablaOtS.Update
    TablaOtS.Close
    
    
    DbsnuevaOtS.Close
       

      With rptOTS
                            
            Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
            Me.cdImpresora.CancelError = True
            Me.cdImpresora.Action = 5
                                             
            .CopiesToPrinter = cdImpresora.Copies
            .ReportFileName = gstrPathReporte & "\OT_VistaPrevia_S.rpt"
        
            '.Destination = crptToPrinter
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .DataFiles(0) = gstrPathReporte & "\BDNuevaOtS.mdb"
            
            'formulas
            
'            .Formulas(0) = "TManoObra=" & mcurTMec + mcurTOtr & ""
'            .Formulas(1) = "TRepuesto=" & mcurTRep & ""
'            .Formulas(2) = "TDyP=" & mcurTCar & ""
'            .Formulas(3) = "Terceros=" & mcurTTer & ""
'            .Formulas(4) = "TMateriales=" & gcurMateriales & "" '
'            .Formulas(5) = "TInsumos=" & curSumaInsumos & ""
'            .Formulas(6) = "tLubricantes=" & mcurTLub & ""
'            .Formulas(7) = "SeguroTaller=" & gcurSeguroTaller & ""
'
'            .Formulas(8) = "TNetoOT=" & mcurTNeto & ""
'            .Formulas(9) = "NombreIva='" & gstrNombreIva & "'"
'            .Formulas(10) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
'            .Formulas(11) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
            
            .Formulas(12) = "Comentario= ''"
            
            .Action = True
            
'            mcurTMec = 0
'            mcurTOtr = 0
'            mcurTRep = 0
'            mcurTLub = 0
'            gcurMateriales = 0
'            curSumaInsumos = 0
'            mcurTCar = 0
'            mcurTTer = 0
'            gcurInsumo = 0
'            mcurTNeto = 0
'            gstrNombreIva = ""
            
    
     End With
   

End Sub


Sub PrintOT()
Dim mstrIdCargo As String
Dim mcurTNeto As Currency
Dim mcurTMec As Currency
Dim mcurTOtr As Currency
Dim mcurTCar As Currency
'Dim mcurTCarAux As Currency 'wcs
Dim mcurTTer As Currency
Dim mcurTRep As Currency
Dim mcurTMat As Currency
Dim mcurTIns As Currency
Dim mcurTLub As Currency
Dim mcurDeducible As Currency
Dim lstrArchivoIni As String
lstrArchivoIni = Command()
gstrPathReporte = LetConnectionString("TLLR", "RPT", lstrArchivoIni, 256)

'/// MODIFICADO POR FDO DIAZ EL 11/12/2000
'/// PREGUNTA PRIMERO SI ES UNA RECEPCION Y DESPUES PREGUNTA DE QUE TIPO DE IMPRESION ES.
'/// SI ES PREIMPRESO COMO AUTOSUMMIT O UNA IMPRESION EN BLANCO

On Error GoTo Solucion

If gstrImpresion = "R" Then
    
    If TipoImpresion = "C" Then  'LA LETRA "C" ES PREIMPRESO AUTOSUMMIT
        If gstrIdEmpresa = "20604506078" Then 'summit motor
            ImprimirDocumentoSMP
        Else
'             ImprimirDocumento gRecepcion
            ImprimirDocumentoASP
        End If
       
    
       
        
    ElseIf TipoImpresion = "P" Then  'LA LETRA "P" ES PREIMPRESO PIAMONTE
        ImprimirDocumentoPiamonte gRecepcion
    ElseIf TipoImpresion = "K" Then  'LA LETRA "K" ES PREIMPRESO klassik car
        ImprimirDocumentoKlassik gRecepcion
    Else
       ImprimirDocumentoRecepcion gRecepcion  ' // FORMATO RECEPCION STANDARD(HOJA EN BLANCO)
    End If
  
ElseIf gstrImpresion = "O" Then


  
    
  If Me.dtcGarantia.BoundText <> "PRE" Then 'reporte OT se cambia a access
  
    If Val(txtDeduciblePesos) = 0 And Val(txtDeducibleUF) = 0 Then
        
        gstrSql = "SELECT ID_TIPO_CARGO FROM TLLR_TIPO_CARGO where Id_Empresa='" & gstrIdEmpresa & "'"
        If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            With gadoPrincipal
                If Not .BOF And Not .EOF Then
                    .MoveFirst
                    While Not .EOF
                        mstrIdCargo = !Id_Tipo_Cargo
                        mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssMec)
                        mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssOtr)
                        mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssCar)
                        mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssTer)
                        
                        mcurTLub = VerificaLubricantesTipoCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep)
                        mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep) '- IIf(mstrIdCargo = "01", gcurMateriales, 0)
'                        mcurTIns = CalculoInsumos(8) + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0)
 
                        mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0) + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0)
'kjcv 27.07.20
'                        mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + (mcurTRep - gcurMateriales) + IIf(mstrIdCargo = "01", gcurInsumo, 0) + IIf(mstrIdCargo = "01", gcurMateriales, 0) + IIf(mstrIdCargo = "01", gcurSeguroTaller, 0)
                        
                        
                        'wcs 05.03.2020 base formada en access
                        Dim DbsnuevaOt As Database
                        Dim TablaOt As DAO.Recordset
                        'Dim i As Integer
                        Dim GcamBaseTemOt As String
                        
                        gstrNombreRecepcionista = NombreRecepcionista(dtcRecepcionista.BoundText)
                        
'                        gstrNombreRecepcionista = NombreRecepcionista("113")
                        
                        
'                       gstrNombreRecepLlamado = NombreRecepcionista(dtcRecepcionista.BoundText)

                        Dim rcOt As Long
                        Dim WinPathOt As String
                        WinPathOt = Space$(300)
                        rcOt = GetWindowsDirectory(WinPathOt, 300)
                        GcamBaseTemOt = Trim$(WinPathOt)
                        GcamBaseTemOt = Mid(GcamBaseTemOt, 1, Len(GcamBaseTemOt) - 1) & "\Temp"
                        
                        
                        Dim wrkPredeterminadoOt As Workspace
                        'Dim prpBucle As Property
                        Set wrkPredeterminadoOt = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
                        If Dir(gstrPathReporte & "\BDNuevaOt.mdb") <> "" Then Kill gstrPathReporte & "\BDNuevaOt.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
                        Set DbsnuevaOt = wrkPredeterminadoOt.CreateDatabase(gstrPathReporte & "\BDNuevaOt.mdb", dbLangGeneral) ' Crea a una base de datos nueva
                        
                        DbsnuevaOt.Execute "CREATE TABLE Tllr_OT (Id_Empresa text,   Id_Sucursal text,   Id_OT text,   Seccion_OT text,   Patente text,   Fecha_Emision text,   Nro_Siniestro text,   Nro_Poliza text,   Liquidador text,   Comentario memo,   Fecha_Liquidacion text,   Kilometros_Recepcion text,   Id_Compañia_Seguro text,   Id_Presupuesto text  )"
                        DbsnuevaOt.Execute "CREATE TABLE Tllr_Vehiculo_Cliente (   Id_Marca text,   Año text,   Nro_Motor text,   VIN text  )"
                        DbsnuevaOt.Execute "CREATE TABLE Tllr_Mecanicos (   Nombre text  )"
                        DbsnuevaOt.Execute "CREATE TABLE Glbl_Marca (   Descripcion text  )"
                        DbsnuevaOt.Execute "CREATE TABLE Glbl_Modelo (   Descripcion text  )"
                        DbsnuevaOt.Execute "CREATE TABLE Glbl_Cliente_Proveedor (   Telefono text,   Rut text,Razon_Social text,Direccion text  )"
                        DbsnuevaOt.Execute "CREATE TABLE Glbl_Color_Exterior (   Descripcion text  )"
                        DbsnuevaOt.Execute "CREATE TABLE Tllr_Parametro (   NotaRecepcion text  )"
                        
                        
                       
                         
                         
                         
                         Dim gadoPrincipalOt As New ADODB.Recordset
                         Dim gstrSqlOt As String

                         gstrSqlOt = " Select Id_Empresa ,Id_Sucursal ,Id_OT ,Seccion_OT ,Patente ,Fecha_Emision ,Nro_Siniestro , Nro_Poliza , Liquidador , Comentario ,Fecha_Liquidacion , Kilometros_Recepcion ,Id_Compañia_Seguro, "
                         gstrSqlOt = gstrSqlOt & "Id_Presupuesto  ,VC.Año,VC.Id_Marca,VC.Nro_Motor,VC.VIN,Marca.Descripcion as MarcaDescripcion ,Modelo.Descripcion as ModeloDescripcion"
                         gstrSqlOt = gstrSqlOt & " ,CP.Rut ,CP.Telefono,CP.Razon_Social,CP.Direccion,CE.Descripcion as CEDescripcion From Tllr_OT  "
                         gstrSqlOt = gstrSqlOt & "Outer Apply(Select Id_Marca,Id_Modelo,Año,Nro_Motor,VIN, Id_Cliente_Proveedor,Id_Color_Exterior From Tllr_Vehiculo_Cliente Where Patente = Tllr_OT.Patente) VC "
                         gstrSqlOt = gstrSqlOt & "Outer Apply(Select Descripcion From Glbl_Marca Where Glbl_Marca.Id_Marca=Vc.Id_Marca) Marca "
                         gstrSqlOt = gstrSqlOt & " Outer Apply(Select Descripcion From Glbl_Modelo Where Id_Marca =Vc.Id_Marca and Id_Modelo= Vc.Id_Modelo) Modelo "
                         gstrSqlOt = gstrSqlOt & "Outer Apply(Select Telefono,Rut,Razon_Social,Direccion From Glbl_Cliente_Proveedor Where Id_Cliente_Proveedor = Vc.Id_Cliente_Proveedor) CP "
                         gstrSqlOt = gstrSqlOt & "Outer Apply(Select Descripcion From Glbl_Color_Exterior Where Id_Color_Exterior = VC.Id_Color_Exterior) CE "
                         gstrSqlOt = gstrSqlOt & " Where Tllr_OT.Id_Empresa= '" & gstrIdEmpresa & "'"
                         gstrSqlOt = gstrSqlOt & " And Tllr_OT.Id_Sucursal= '" & gstrIdSucursal & "'"
                         gstrSqlOt = gstrSqlOt & " And Tllr_OT.Id_OT= '" & lblNroRecepcion & "'"
                         gstrSqlOt = gstrSqlOt & " And Tllr_OT.Seccion_OT= '" & gstrSeccion & "'"
                         
                         
                         Dim Id_EmpresaBd As String
                         Dim Id_SucursalBd As String
                         Dim Id_OTBd As String
                         Dim Seccion_OTBd As String
                         Dim PatenteBd As String
                         Dim Fecha_EmisionBd As String
                         Dim Nro_SiniestroBd As String
                         Dim Nro_PolizaBd As String
                         Dim LiquidadorBd As String
                         Dim ComentarioBd As String
                         Dim Fecha_LiquidacionBd As String
                         Dim Kilometros_RecepcionBd As String
                         Dim Id_Compañia_SeguroBd As String
                         Dim Id_PresupuestoBd As String
                         Dim AñoBd As String
                         Dim Id_MarcaBd As String
                         Dim Nro_MotorBd As String
                         Dim VINBd As String
                         Dim MarcaDescripcionBd As String
                         Dim ModeloDescripcionBd As String
                        
                         Dim RutBd As String
                         Dim TelefonoBd As String
                         Dim Razon_SocialBd As String
                         Dim DireccionBd As String
                         Dim CEDescripcionBd As String
                         
                         
                         If Conexion.SendHost(gstrSqlOt, gadoPrincipalOt, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                            With gadoPrincipalOt
                            If Not .BOF And Not .EOF Then
                                 .MoveFirst
                                 While Not .EOF
                                     Id_EmpresaBd = ValorNulo(!Id_Empresa)
                                     Id_SucursalBd = ValorNulo(!Id_Sucursal)
                                     Id_OTBd = ValorNulo(!Id_OT)
                                     Seccion_OTBd = ValorNulo(!Seccion_OT)
                                     PatenteBd = ValorNulo(!Patente)
                                     Fecha_EmisionBd = ValorNulo(!Fecha_Emision)
                                     Nro_SiniestroBd = ValorNulo(!Nro_Siniestro)
                                     Nro_PolizaBd = ValorNulo(!Nro_Poliza)
                                     LiquidadorBd = ValorNulo(!Liquidador)
                                     ComentarioBd = ValorNulo(!Comentario)
                                     Fecha_LiquidacionBd = IIf(IsNull(!Fecha_Liquidacion), "", !Fecha_Liquidacion)
                                     Kilometros_RecepcionBd = ValorNulo(!Kilometros_Recepcion)
                                     Id_Compañia_SeguroBd = ValorNulo(!Id_Compañia_Seguro)
                                     Id_PresupuestoBd = ValorNulo(!Id_Presupuesto)
                                     AñoBd = ValorNulo(!Año)
                                     Id_MarcaBd = ValorNulo(!Id_Marca)
                                     Nro_MotorBd = ValorNulo(!Nro_Motor)
                                     VINBd = ValorNulo(!VIN)
                                     MarcaDescripcionBd = ValorNulo(!MarcaDescripcion)
                                     ModeloDescripcionBd = ValorNulo(!ModeloDescripcion)
                                   
                                     RutBd = ValorNulo(!rut)
                                     TelefonoBd = ValorNulo(!Telefono)
                                     Razon_SocialBd = ValorNulo(!Razon_Social)
                                     DireccionBd = ValorNulo(!Direccion)
                                     CEDescripcionBd = ValorNulo(!CEDescripcion)

                                    .MoveNext
                                  Wend

                            End If
                            End With
                         End If
                        Conexion.CloseHost gadoPrincipalOt
                         
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Tllr_OT")
                        TablaOt.AddNew
                        
        
                        TablaOt!Id_Empresa = Id_EmpresaBd
                        TablaOt!Id_Sucursal = Id_SucursalBd
                        TablaOt!Id_OT = Id_OTBd
                        TablaOt!Seccion_OT = Seccion_OTBd
                        TablaOt!Patente = PatenteBd
                        TablaOt!Fecha_Emision = Fecha_EmisionBd
                        TablaOt!Nro_Siniestro = Nro_SiniestroBd
                        TablaOt!Nro_Poliza = Nro_PolizaBd
                        TablaOt!Liquidador = LiquidadorBd
                        TablaOt!Comentario = ComentarioBd
                        TablaOt!Fecha_Liquidacion = Fecha_LiquidacionBd
                        TablaOt!Kilometros_Recepcion = Kilometros_RecepcionBd
                        TablaOt!Id_Compañia_Seguro = Id_Compañia_SeguroBd
                        TablaOt!Id_Presupuesto = Id_PresupuestoBd
                      
                        TablaOt.Update
                        TablaOt.Close
                    
                    
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Tllr_Vehiculo_Cliente")
                        TablaOt.AddNew
                        TablaOt!Id_Marca = Id_MarcaBd
                        TablaOt!Año = AñoBd
                        TablaOt!Nro_Motor = Nro_MotorBd
                        TablaOt!VIN = VINBd
                        TablaOt.Update
                        TablaOt.Close
                        
                        
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Tllr_Mecanicos")
                        TablaOt.AddNew
                        TablaOt!Nombre = "" 'se deja en blanco ya que no se imprime en el reporte /verificar como obtener valor
                        TablaOt.Update
                        TablaOt.Close
                        
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Glbl_Marca")
                        TablaOt.AddNew
                        TablaOt!Descripcion = MarcaDescripcionBd
                        TablaOt.Update
                        TablaOt.Close
                        
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Glbl_Modelo")
                        TablaOt.AddNew
                        TablaOt!Descripcion = ModeloDescripcionBd
                      
                        TablaOt.Update
                        TablaOt.Close
                        
                            
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Glbl_Cliente_Proveedor")
                        TablaOt.AddNew
                        TablaOt!Telefono = TelefonoBd
                        TablaOt!rut = RutBd
                        TablaOt!Razon_Social = Razon_SocialBd
                        TablaOt!Direccion = DireccionBd
                        TablaOt.Update
                        TablaOt.Close
                        
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Glbl_Color_Exterior")
                        TablaOt.AddNew
                        TablaOt!Descripcion = CEDescripcionBd
                        TablaOt.Update
                        TablaOt.Close
                        
                        Set TablaOt = DbsnuevaOt.OpenRecordset("SELECT * FROM Tllr_Parametro")
                        TablaOt.AddNew
                        TablaOt!NotaRecepcion = "." 'no se muestra en el reporte
                        TablaOt.Update
                        TablaOt.Close
                        
                        
                        DbsnuevaOt.Close
                        
                        
                        
                        If mcurTNeto > 0 Then
                          ' Antes
                            With rptOT
                            
                            Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
                            Me.cdImpresora.CancelError = True
                            Me.cdImpresora.Action = 5
                                                              
                            .CopiesToPrinter = cdImpresora.Copies
                            If gstrServiciosMarca = "S" Then
                                .ReportFileName = gstrPathReporte & "\OTMM.rpt"
                            Else
                                .ReportFileName = gstrPathReporte & "\OT_Original.rpt"
                            End If
                            .Destination = crptToPrinter
'                            .Destination = crptToWindow
                            .WindowState = crptMaximized
'                            .DataFiles(0) = gstrPathReporte & "\BDNuevaOt.mdb"
                                
                               .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
                               .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
                               .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
                               .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
                               .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
                               .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
                               .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
                               .Formulas(7) = "TMecanica=" & mcurTMec & "" '10/03/2020 wcs da error y no se visualiza en el rpt
                              .Formulas(8) = "TOtros=" & mcurTOtr & ""   '10/03/2020 wcs da error y no se visualiza en el rpt
                              .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
                               .Formulas(10) = "TRepuesto=" & mcurTRep - (mcurTLub + gcurMateriales + curSumaInsumos) & ""
'                                .Formulas(10) = "TRepuestoAux=" & mcurTRep - (mcurTLub + gcurMateriales + curSumaInsumos) & ""

                              .Formulas(11) = "TDyP=" & mcurTCar & "" ' wcs comentado por dar error
                              .Formulas(12) = "TTerceros=" & mcurTTer & "" '10/10/2020 wcs da error cambio a TTercerosAux
                                .Formulas(13) = "TMateriales=" & gcurMateriales & "" '& IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
                                .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo + curSumaInsumos, 0) & ""
                                'kjcv 27.07.20
'                                .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = "01", gcurInsumo, 0) & ""

                                .Formulas(15) = ""
                                .Formulas(16) = ""
                                .Formulas(17) = "TNetoOT=" & mcurTNeto & ""
                                .Formulas(18) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
                                .Formulas(19) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
                                .Formulas(20) = "TLubricantes=" & mcurTLub & ""
                                .Formulas(21) = "SeguroTaller=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0) & ""
                                .Formulas(22) = "NotaRecepcion='" & IIf(gstrNotaRecepcion = "", "", "OK") & "'"
                                .Formulas(23) = "TipoCargo='" & mstrIdCargo & "'"
                                .Formulas(24) = "NombreIva='" & gstrNombreIva & "'"
                                .Formulas(25) = "Tdecimal=" & gintDecimalesMoneda & ""
                                .Formulas(26) = "NombreRut='" & gstrNombreRut & "'"
                                .Formulas(27) = "NombrePatente='" & gstrNombrePatente & "'"
                                .Formulas(28) = "FamiliaInsumos='" & gstrCodigoInsumos & "'"
                                .Formulas(29) = "FamiliaLubricantes='" & gstrCodigoLubricantes & "'"
                                .Formulas(30) = "FamiliaMateriales='" & gstrCodigoMateriales & "'"
                                .Formulas(31) = "EditaRut='" & gstrEditaRut & "'"
                                .Formulas(32) = "TipodeOt='" & Me.dtcGarantia.Text & "'"

                                .Formulas(33) = "TDyP=" & mcurTCar & "" 'wcs cambiado antes era TDyP por error de connect sql
                                .Formulas(34) = "TTerceros=" & mcurTTer & ""

                                '.Connect = "Driver={SQL Server};Server=CHACLLA;UID=sa;PWD=Llosa1936;Database=elisa;" 'Conexion.ConnectionString
'                                .Connect = "Driver={SQL Server};Data Source=CHACLLA;UID=sa;PWD=Llosa1936;Initial Catalog=elisa;" 'Conexion.ConnectionString
                               ' .Connect = "Driver={SQL Server};Server=WIRACOCHA;UID=sa;PWD=Llosa1936;Database=Prueba;" 'Conexion.ConnectionString
                                .Connect = "Driver={SQL Server};Server=CHACLLA;UID=sa;PWD=Llosa1936;Database=elisa;"
                                .SelectionFormula = "{Tllr_OT.Id_Empresa}='" & gstrIdEmpresa & "' And {Tllr_OT.Id_Sucursal}='" & gstrIdSucursal & "' And {Tllr_OT.Id_OT}='" & lblNroRecepcion & "' And {Tllr_OT.Seccion_OT}='" & gstrSeccion & "'"
                                '.Connect = "Provider=Microsoft.Jet.OLEDB.4.0;DataSource=C:\Sistemas auto summit\Elisa\reportes\taller\BDNuevaOt.mdb;Persist Security Info=False;"
                                .Action = True
                            End With
                            .MoveNext
                        Else
                            .MoveNext
                        End If
                        mcurTMec = 0
                        mcurTOtr = 0
                        mcurTCar = 0
                        mcurTTer = 0
                        mcurTRep = 0
                        mcurTLub = 0
                        mcurTNeto = 0

                    Wend
                End If
            End With
        Else
            DoEvents
            Exit Sub
        End If
    Else    '/////////////////////////////////////////////////deducible <>0
        Dim mblndeducible As Boolean
        mcurDeducible = CCur(Val(txtDeduciblePesos))
        gstrSql = "SELECT ID_TIPO_CARGO FROM TLLR_TIPO_CARGO where Id_Empresa='" & gstrIdEmpresa & "'"
        If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            With gadoPrincipal
                If Not .BOF And Not .EOF Then
                    .MoveLast
                    While Not .BOF
                        mstrIdCargo = !Id_Tipo_Cargo
                        mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssMec)
                        mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssOtr)
                        mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssCar)
                        mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssTer)
                        'MODIFICADO POR FDO DIAZ EL 04/01/2001
                        mcurTLub = VerificaLubricantesTipoCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep)
                        mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep) '- IIf(mstrIdCargo = "01", gcurMateriales, 0)
                        'mcurTIns = CalculoInsumos(8)
                        mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo, 0) + IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0)
                        
                        'si solo existe deducible
                        If mcurTNeto = 0 Then
                            If mstrIdCargo = gstrCargoDeducibleMas Then
                                mblndeducible = True
                            End If
                        End If
                        If mcurTNeto > 0 Or mblndeducible = True Then
                            With rptOT
                            
                                Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
                                Me.cdImpresora.CancelError = True
                                Me.cdImpresora.Action = 5
                                
                                .CopiesToPrinter = cdImpresora.Copies
                                If gstrServiciosMarca = "S" Then
                                    .ReportFileName = gstrPathReporte & "\OTCDMM.rpt"
                                Else
                                    .ReportFileName = gstrPathReporte & "\OTCD.rpt"
                                End If
                                .Destination = crptToPrinter
                                .WindowState = crptMaximized
                                If gstrIdEmpresa = "832207004" Or InStr(gstrEmpresa, "SERINFO") = 1 Then
                                    .Destination = crptToWindow
                                End If
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
'                                .Formulas(10) = "TRepuesto=" & mcurTRep - (mcurTLub + gcurMateriales + curSumaInsumos) & ""
'kjcv 27.07.20
                                .Formulas(10) = "TRepuesto=" & mcurTRep - mcurTLub - mcurTIns & ""
                                .Formulas(11) = "TDyP=" & mcurTCar & ""
                                .Formulas(12) = "TTerceros=" & mcurTTer & ""

                                .Formulas(13) = "TMateriales=" & gcurMateriales & ""         '& IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
                                'kjcv 27.07.20
                                .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurInsumo + curSumaInsumos, 0) & ""
                                .Formulas(20) = "TLubricantes=" & mcurTLub & ""
                                
                                If mstrIdCargo = gstrCargoDeducibleMenos Then
                                    If mcurDeducible <= mcurTNeto Then
                                        .Formulas(15) = "Anexo= 'Deducible ( - )'"
                                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
                                        mcurTNeto = mcurTNeto - mcurDeducible
                                    End If
                                ElseIf mstrIdCargo = gstrCargoDeducibleMas Then
                                        .Formulas(15) = "Anexo= 'Deducible ( + )'"
                                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
                                        mcurTNeto = mcurTNeto + mcurDeducible
                                Else
                                    .Formulas(15) = ""
                                    .Formulas(16) = ""
                                End If
                                .Formulas(17) = "TNetoOT=" & mcurTNeto & ""
                                .Formulas(18) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
                                .Formulas(19) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
                                .Formulas(21) = "SeguroTaller=" & IIf(mstrIdCargo = gstrCargoDeducibleMas, gcurSeguroTaller, 0) & ""
                                .Formulas(22) = "NotaRecepcion='" & IIf(gstrNotaRecepcion = "", "", "OK") & "'"
                                .Formulas(23) = "TipoCargo='" & mstrIdCargo & "'"
                                .Formulas(24) = "NombreIva='" & gstrNombreIva & "'"
                                .Formulas(25) = "Tdecimal=" & gintDecimalesMoneda & ""
                                .Formulas(26) = "NombreRut='" & gstrNombreRut & "'"
                                .Formulas(27) = "NombrePatente='" & gstrNombrePatente & "'"
                                .Formulas(28) = "FamiliaInsumos='" & gstrCodigoInsumos & "'"
                                .Formulas(29) = "FamiliaLubricantes='" & gstrCodigoLubricantes & "'"
                                .Formulas(30) = "FamiliaMateriales='" & gstrCodigoMateriales & "'"
                                .Formulas(31) = "EditaRut='" & gstrEditaRut & "'"
                                .Formulas(32) = "TipodeOt='" & Me.dtcGarantia.Text & "'"
'                                .Connect = Conexion.ConnectionString
                                .Connect = "Driver={SQL Server};Server=CHACLLA;UID=sa;PWD=Llosa1936;Database=elisa;" 'Conexion.ConnectionString
                                .SelectionFormula = "{Tllr_OT.Id_Empresa}='" & gstrIdEmpresa & "' And {Tllr_OT.Id_Sucursal}='" & gstrIdSucursal & "' And {Tllr_OT.Id_OT}='" & lblNroRecepcion & "' And {Tllr_OT.Seccion_OT}='" & gstrSeccion & "'"

                                .Action = True
                                
                            End With
                            .MovePrevious
                        Else
                            .MovePrevious
                        End If
                        mcurTMec = 0
                        mcurTOtr = 0
                        mcurTCar = 0
                        mcurTTer = 0
                        mcurTRep = 0
                        mcurTLub = 0
                        mcurTNeto = 0
                    Wend
                End If
            End With
        Else
            DoEvents
            Exit Sub
        End If
    End If
    
  Else  '//// es presupuesto
        mcurDeducible = CCur(Val(txtDeduciblePesos))
        mcurTMec = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssMec)
        mcurTOtr = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssOtr)
        mcurTCar = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssCar)
        mcurTTer = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssTer)
        'MODIFICADO POR FDO DIAZ EL 04/01/2001
        mcurTLub = VerificaLubricantesTipoCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssRep)
        mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, "", gstrSeccion, ssRep) '- IIf(mstrIdCargo = "01", gcurMateriales, 0)
        mcurTIns = CalculoInsumos(8)
        mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep
  
  'kjcv 05.09.16 base formada en access
        Dim Dbsnueva As Database
        Dim Tabla As DAO.Recordset
        Dim i As Integer
        Dim GcamBaseTem As String
        
        gstrNombreRecepcionista = NombreRecepcionista(dtcRecepcionista.BoundText)
'        gstrNombreRecepLlamado = NombreRecepcionista(dtcRecepcionista.BoundText)

                Dim rc As Long
                Dim WinPath As String
                WinPath = Space$(300)
                rc = GetWindowsDirectory(WinPath, 300)
                GcamBaseTem = Trim$(WinPath)
                GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
                '---------------------------------------
                
                Dim wrkPredeterminado As Workspace
                Dim prpBucle As Property
                Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
                If Dir(gstrPathReporte & "\BDNuevaPresu.mdb") <> "" Then Kill gstrPathReporte & "\BDNuevaPresu.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
                Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNuevaPresu.mdb", dbLangGeneral) ' Crea a una base de datos nueva
                Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (OT text, Seccion text,Recepcionista text,FLiquida text, Cliente text,Direccion text,DNI text,Telefono text,Marca text, Modelo text,Patente text,VIN text,Color text,año text,Motor text,Siniestro text,Poliza text, Liquidador text,Compañia text, DeduSoles text,DeduDolar text, Observaciones memo,Fecha_Emision text, Kilometraje_Recepcion text)"
                Dbsnueva.Execute "CREATE TABLE T_TOTALES(OT text,TManoObra double, TRepuestos double, TPyP double,TTerceros double, Insumos double,Lubricantes double, TOtros double, SubTotal double,IVA double,Total double)"
                Dbsnueva.Execute "CREATE TABLE T_PARAMECANICA (OT text,IdServicio text,Descripcion text,PrecioU double,Cargo text,Horas text,Porcentaje_Dscto text, Monto_Dscto double, MSubtotal double)"
                Dbsnueva.Execute "CREATE TABLE T_PARASERVICIO (OT text,IdOtroServicio text,Servicio text,PrecioU double,Cargo text, Horas text,Porcentaje_Dscto text,Monto_Dscto double, OSubTotal double)"
                Dbsnueva.Execute "CREATE TABLE T_PARACARROCERIA (OT text,Carroceria text,CSubtotal double)"
                Dbsnueva.Execute "CREATE TABLE T_PARATERCEROS(OT text,IdServicioTercero text, Tercero text, Cargo text,Porcentaje_Dscto text,Monto_Dscto double,TSubtotal double)"
                'Dbsnueva.Execute "CREATE TABLE T_PARAREPUESTOS(OT text,IdItem text, Saldo text, Pieza text, Cargo text, Valor text, Cantidad text,RSubtotal double)"
                'kjcv 11.02.20
                Dbsnueva.Execute "CREATE TABLE T_PARAREPUESTOS(OT text,IdItem text, Saldo text, Pieza text, Cargo text,Valor double, Cantidad text,Porcentaje_Dscto text, Monto_Dscto double, RSubtotal double)"
                
                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
                Tabla.AddNew
                Tabla!OT = Me.lblNroRecepcion
                Tabla!Seccion = gstrSeccion
                Tabla!Recepcionista = gstrNombreRecepcionista
                Tabla!FLiquida = Me.lblFechaLiquidacion
                Tabla!Cliente = Me.lblCliente.Caption
                Tabla!Direccion = TraeDireccion(Me.lblIdCliente)
                Tabla!DNI = Me.lblIdCliente
                Tabla!Telefono = Me.lblFono
                Tabla!Marca = Me.lblMarca
                Tabla!Modelo = Me.lblModelo
                Tabla!Patente = Me.txtPatente
                Tabla!VIN = Me.lblVin
                Tabla!Color = Me.lblColorE
                Tabla!Año = Me.txtAño
                Tabla!motor = Me.lblMotor
                Tabla!Siniestro = Me.txtNroSiniestro
                Tabla!Poliza = Me.txtNroPoliza
                Tabla!Liquidador = Me.txtLiquidador
                Tabla!Compañia = Me.lblCompañia
                Tabla!DeduSoles = Me.txtDeduciblePesos
                Tabla!DeduDolar = Me.txtDeducibleUF
                Tabla!Observaciones = Me.txtComentario
                Tabla!Fecha_Emision = IIf(IsNull(pckFechaAtencion.Value), "", pckFechaAtencion.Value)
                Tabla!Kilometraje_Recepcion = IIf(IsNull(txtKilAct.Text), "", txtKilAct.Text)
                
                
                Tabla.Update
                Tabla.Close
            
                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAMECANICA")
                For i = 1 To lvwServiciosMecanica.ListItems.Count
                    Set lvwServiciosMecanica.SelectedItem = lvwServiciosMecanica.ListItems(i)
                    Tabla.AddNew
                    Tabla!idServicio = lvwServiciosMecanica.ListItems(i)
                    Tabla!Descripcion = IIf(lvwServiciosMecanica.SelectedItem.SubItems(1) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(1))
                    Tabla!PrecioU = IIf(lvwServiciosMecanica.SelectedItem.SubItems(5) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(5))
                    Tabla!CARGO = IIf(lvwServiciosMecanica.SelectedItem.SubItems(7) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(7))
                    Tabla!Horas = IIf(lvwServiciosMecanica.SelectedItem.SubItems(2) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(2))
                    Tabla!Porcentaje_Dscto = IIf(lvwServiciosMecanica.SelectedItem.SubItems(4) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(4))
                    Tabla!monto_Dscto = IIf(lvwServiciosMecanica.SelectedItem.SubItems(5) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(5))
                    Tabla!MSubtotal = IIf(lvwServiciosMecanica.SelectedItem.SubItems(10) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(10))
                    Tabla.Update
                Next i
                Tabla.Close
                
                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARASERVICIO")
                For i = 1 To lvwOtrosServicios.ListItems.Count
                    Set lvwOtrosServicios.SelectedItem = lvwOtrosServicios.ListItems(i)
                    Tabla.AddNew
                    Tabla!IdOtroServicio = lvwOtrosServicios.ListItems(i)
                    Tabla!servicio = IIf(lvwOtrosServicios.SelectedItem.SubItems(1) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(1))
                    Tabla!PrecioU = IIf(lvwOtrosServicios.SelectedItem.SubItems(3) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(3))
                    Tabla!CARGO = IIf(lvwOtrosServicios.SelectedItem.SubItems(7) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(7))
                    Tabla!Horas = IIf(lvwOtrosServicios.SelectedItem.SubItems(2) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(2))
                    Tabla!Porcentaje_Dscto = IIf(lvwOtrosServicios.SelectedItem.SubItems(4) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(4))
                    Tabla!monto_Dscto = IIf(lvwOtrosServicios.SelectedItem.SubItems(5) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(5))
                    Tabla!OSubtotal = IIf(lvwOtrosServicios.SelectedItem.SubItems(10) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(10))
                    Tabla.Update
                Next i
                Tabla.Close
                
                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARACARROCERIA")
                For i = 1 To lvwServiciosCarroceria.ListItems.Count
                    Set lvwServiciosCarroceria.SelectedItem = lvwServiciosCarroceria.ListItems(i)
                    Tabla.AddNew
                    Tabla!Carroceria = lvwServiciosCarroceria.SelectedItem.SubItems(2)
                    Tabla!CSubtotal = IIf(lvwServiciosCarroceria.SelectedItem.SubItems(16) = "", " ", lvwServiciosCarroceria.SelectedItem.SubItems(16))
                    Tabla.Update
                Next i
                Tabla.Close
                
                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARATERCEROS")
                For i = 1 To lvwServiciosTerceros.ListItems.Count
                    Set lvwServiciosTerceros.SelectedItem = lvwServiciosTerceros.ListItems(i)
                    Tabla.AddNew
                    Tabla!IdServicioTercero = lvwServiciosTerceros.ListItems(i)
                    Tabla!Tercero = IIf(lvwServiciosTerceros.SelectedItem.SubItems(3) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(3))
                    Tabla!CARGO = IIf(lvwServiciosTerceros.SelectedItem.SubItems(13) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(13))
                    Tabla!Porcentaje_Dscto = IIf(lvwServiciosTerceros.SelectedItem.SubItems(10) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(10))
                    Tabla!monto_Dscto = IIf(lvwServiciosTerceros.SelectedItem.SubItems(11) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(11))
                    Tabla!TSubtotal = IIf(lvwServiciosTerceros.SelectedItem.SubItems(12) = "", " ", lvwServiciosTerceros.SelectedItem.SubItems(12))
                    Tabla.Update
                Next i
                Tabla.Close
                
                    
                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPUESTOS")
                For i = 1 To lvwRepuestos.ListItems.Count
                    Set lvwRepuestos.SelectedItem = lvwRepuestos.ListItems(i)
                    Tabla.AddNew
                    Tabla!IdItem = lvwRepuestos.ListItems(i)
                    Tabla!Saldo = IIf(lvwRepuestos.SelectedItem.SubItems(12) = "", " ", lvwRepuestos.SelectedItem.SubItems(12))
                    Tabla!pieza = IIf(lvwRepuestos.SelectedItem.SubItems(1) = "", " ", lvwRepuestos.SelectedItem.SubItems(1))
                    Tabla!CARGO = IIf(lvwRepuestos.SelectedItem.SubItems(6) = "", " ", lvwRepuestos.SelectedItem.SubItems(6))
                    Tabla!Valor = IIf(lvwRepuestos.SelectedItem.SubItems(3) = "", " ", lvwRepuestos.SelectedItem.SubItems(3))
                    Tabla!Porcentaje_Dscto = IIf(lvwRepuestos.SelectedItem.SubItems(4) = "", " ", lvwRepuestos.SelectedItem.SubItems(4))
                    Tabla!monto_Dscto = IIf(lvwRepuestos.SelectedItem.SubItems(5) = "", " ", lvwRepuestos.SelectedItem.SubItems(5))
                    Tabla!cantidad = IIf(lvwRepuestos.SelectedItem.SubItems(2) = "", " ", lvwRepuestos.SelectedItem.SubItems(2))
                    Tabla!RSubtotal = IIf(lvwRepuestos.SelectedItem.SubItems(8) = "", " ", lvwRepuestos.SelectedItem.SubItems(8))
                    Tabla.Update
                Next i
                Tabla.Close
                
                Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_TOTALES")
                
                    Tabla.AddNew
                    'Tabla!TManoObra = mcurTMec
                    Tabla!TManoObra = mcurTMec + mcurTOtr
                    Tabla!TRepuestos = mcurTRep
                    Tabla!TPyP = mcurTCar
                    Tabla!TTerceros = mcurTTer
                    Tabla!Insumos = mcurTIns
                    Tabla!Lubricantes = mcurTLub
'                    Tabla!TOtros = mcurTOtr
                    Tabla!SubTotal = mcurTNeto
                    Tabla!IVA = Round(mcurTNeto * 0.18, 2)
                    Tabla!Total = Round(1.18 * mcurTNeto, 2)
                    Tabla.Update
              
                Tabla.Close
                
                Dbsnueva.Close
  
        
                    
        If mcurTNeto > 0 Then
            With rptOT
                            Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
                            Me.cdImpresora.CancelError = True
                            Me.cdImpresora.Action = 5
                                                              
                            .CopiesToPrinter = cdImpresora.Copies
                           
                                   

                If gstrServiciosMarca = "S" Then
                    .ReportFileName = gstrPathReporte & "\OTPresupuestoMM.rpt"
                Else
                    .ReportFileName = gstrPathReporte & "\Presupuesto.rpt"
                End If
                
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .DataFiles(0) = gstrPathReporte & "\BDNuevaPresu.mdb"
                                       
'                .Formulas(0) = "IDEmpresa='" & gstrIdEmpresa & "'"
'                .Formulas(1) = "IDSucursal='" & gstrIdSucursal & "'"
'                .Formulas(2) = "NumeroOT='" & lblNroRecepcion & "'"
'                .Formulas(3) = "SeccionOT='" & gstrSeccion & "'"
'                .Formulas(4) = "RazonSocial='" & gstrEmpresa & "'"
'                .Formulas(5) = "Sucursal='" & gstrSucursal & "'"
'                .Formulas(6) = "Direccion='" & gstrDirSuc & "'"
'
'                .Formulas(7) = "TMecanica=" & mcurTMec & ""
'                .Formulas(8) = "TOtros=" & mcurTOtr & ""
'                .Formulas(9) = "TManoObra=" & mcurTMec + mcurTOtr & ""
'                .Formulas(10) = "TRepuesto=" & mcurTRep - (mcurTLub + gcurMateriales + mcurTIns) & ""
'                .Formulas(11) = "TDyP=" & mcurTCar & ""
'                .Formulas(12) = "TTerceros=" & mcurTTer & ""
'
'                .Formulas(13) = "TMateriales=" & gcurMateriales     '& IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
'                .Formulas(14) = "TInsumos=" & mcurTIns              '& IIf(mstrIdCargo = "01", gcurInsumo, 0) & ""
'                .Formulas(20) = "TLubricantes=" & mcurTLub & ""
'                .Formulas(21) = "TelefonoE='Fono: " & gstrTelefono & " Fax: " & gstrFax & "'"
'
'                If mstrIdCargo = gstrCargoDeducibleMenos Then
'                    If mcurDeducible <= mcurTNeto Then
'                        .Formulas(15) = "Anexo= 'Deducible ( - )'"
'                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
'
'                    End If
'                ElseIf mstrIdCargo = gstrCargoDeducibleMas Then
'                        .Formulas(15) = "Anexo= 'Deducible ( + )'"
'                        .Formulas(16) = "TAnexo=" & mcurDeducible & ""
'
'                End If
'                .Formulas(17) = "TNetoOT=" & mcurTNeto & ""
'                .Formulas(18) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
'                .Formulas(19) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
'                .Formulas(22) = "NombreIva='" & gstrNombreIva & "'"
'                .Formulas(23) = "Tdecimal=" & gintDecimalesMoneda & ""
'                .Formulas(24) = "NombreRut='" & gstrNombreRut & "'"
'                .Formulas(25) = "NombrePatente='" & gstrNombrePatente & "'"
'                .Formulas(26) = "EditaRut='" & gstrEditaRut & "'"
'                .Formulas(27) = "FamiliaInsumos='" & gstrCodigoInsumos & "'"
'                .Formulas(28) = "FamiliaLubricantes='" & gstrCodigoLubricantes & "'"
'                .Formulas(29) = "FamiliaMateriales='" & gstrCodigoMateriales & "'"
              ' .Connect = "Driver={SQL Server};Server=wiracocha;UID=sa;PWD=Llosa1936;Database=elisa;" 'Conexion.ConnectionString
'                .SelectionFormula = "{Tllr_OT.Id_Empresa}='" & gstrIdEmpresa & "' And {Tllr_OT.Id_Sucursal}='" & gstrIdSucursal & "' And {Tllr_OT.Id_OT}='" & lblNroRecepcion & "' And {Tllr_OT.Seccion_OT}='" & gstrSeccion & "'"
                .Destination = crptToWindow
                .Action = True
            End With
            mcurTMec = 0
            mcurTOtr = 0
            mcurTCar = 0
            mcurTTer = 0
            mcurTRep = 0
            mcurTLub = 0
            mcurTNeto = 0
        End If
  End If
End If

Solucion:
    If Err.Number = 32755 Then
        MsgBox "Impresión Cancelada por el usuario", vbInformation, "Advertencia"
        Screen.MousePointer = 1
        Exit Sub
    End If
    If Err.Number <> 0 Then
        MsgBox "Se ha producido el siguiente error " & Chr(13) & Err.Number & " " & Err.Description, vbExclamation, "Advertencia"
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
    mcurTRep = TotalSeccionCargo(gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, mstrIdCargo, gstrSeccion, ssRep) - IIf(mstrIdCargo = "01", gcurMateriales, 0)
    mcurTNeto = mcurTMec + mcurTOtr + mcurTCar + mcurTTer + mcurTRep + IIf(mstrIdCargo = "01", gcurMateriales, 0) + IIf(mstrIdCargo = "01", gcurInsumo, 0)

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
        
        .Formulas(13) = "TMateriales=" & IIf(mstrIdCargo = "01", gcurMateriales, 0) & ""
        .Formulas(14) = "TInsumos=" & IIf(mstrIdCargo = "01", gcurInsumo, 0) & ""
        .Formulas(15) = "TNetoOT=" & mcurTNeto & ""
        .Formulas(16) = "IVA=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ""
        .Formulas(17) = "TOT=" & mcurTNeto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ""
        .Action = True
    End With
End Sub

Function ValidaDatos() As Boolean
Dim j As Integer
Dim i As Integer
Dim cont As Integer
Dim tablaParam As New ADODB.Recordset
Dim lstrSQL As String
Dim SW As Integer
Dim val_real As Double
cont = 0

ValidaDatos = True

'kjcv 19.01.16
    For i = 1 To Me.lvwServiciosTerceros.ListItems.Count
        If Trim(Me.lvwServiciosTerceros.ListItems(i).SubItems(4)) = "" Then
            MsgBox "No existe Numero Factura en la Línea " & i & " de los Servicios de Terceros" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    
    'kjcv 20.07.16
    'Asignacion de Mecanico
    'mecanica
    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
        If Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(8)) = gstrMecanicoDefectoSecMec Then
            MsgBox "Debe asignar un Mecánico de los Servicios de Mecánica" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    
     'otros servicios
    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
        If Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(8)) = gstrMecanicoDefectoSecMec Then
            MsgBox "Debe asignar un Mecánico de los Otros Servicios" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    
    'kjcv 20.07.16
    'Valida nro Horas
    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
        If Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(2)) = "0.0" And Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(6)) <> gstrIdCargoInterno Then
            MsgBox "Debe ingresar Nro Horas en Servicios de Mecánica" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    
    'otros servicios
    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
        If Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(2)) = "0.00" And Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(6)) <> gstrIdCargoInterno Then
            MsgBox "Debe ingresar Nro Horas en Otros Servicios" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    

'valida subtotales en cero (0) según parametro
If gblnValidaServiciosCero = True Then

    'mecanica
    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
        If Trim(Me.lvwServiciosMecanica.ListItems(i).SubItems(10)) = "0" Then
            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Servicios de Mecanica" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    
    'carrocería
    For i = 1 To Me.lvwServiciosCarroceria.ListItems.Count
        If Trim(Me.lvwServiciosCarroceria.ListItems(i).SubItems(16)) = "0" Then
            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Servicios de Carrocería" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    
    'otros servicios
    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
        If Trim(Me.lvwOtrosServicios.ListItems(i).SubItems(10)) = "0" Then
            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Otros Servicios" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
        
    'terceros
    For i = 1 To Me.lvwServiciosTerceros.ListItems.Count
        If Trim(Me.lvwServiciosTerceros.ListItems(i).SubItems(12)) = "0" Then
            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Servicios de Terceros" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
    
    
        
    'repuestos
    For i = 1 To Me.lvwRepuestos.ListItems.Count
        If Trim(Me.lvwRepuestos.ListItems(i).SubItems(8)) = "0" Then
            MsgBox "Existe un Valor 0 en la Línea " & i & " de los Repuestos" & Chr(13) & " La Liquidación se cancela", vbExclamation, "Liquidacion de OT"
            ValidaDatos = False
            Exit Function
        End If
    Next
End If

With lvwRepuestos
    i = 1
    j = .ListItems.Count
    For i = 1 To .ListItems.Count
           
       If Trim(lvwRepuestos.ListItems(j).SubItems(11)) = "PRESUPUESTO" Then
           
            If Trim(lvwRepuestos.ListItems(j).SubItems(2)) <> Trim(lvwRepuestos.ListItems(j).SubItems(13)) Then
                SW = 1
                val_real = Val(Trim(lvwRepuestos.ListItems(j).SubItems(2)))
            End If
           
            If MsgBox("El Repuesto " & Me.lvwRepuestos.ListItems(j).SubItems(1) & " No esta Descontado de Stock-Pro" & Chr(13) & "¿Desea Eliminarlo de la OT?", vbQuestion + vbYesNo + vbDefaultButton2, "Advertencia") = vbYes Then
                lvwRepuestos.ListItems.Remove (j)
                AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                TotalFinal
            End If
       End If
        
        j = j - 1
    Next
      
End With

If SW = 1 Then
    GrabarRegistro
    MsgBox "Se han actualizado las cantidades ", vbInformation
End If
        
End Function

Sub EstadosOT(ModeAction As gAccionEstadoOT)

Dim SW As Integer
SW = 1
gflag = False

If ModeAction = gOTActivar Then
    '//////////////////////////////////////VERIFICAR
    Act = 1
    If VeriLiq() = True And gflag = True Then
        gstrSql = "UPDATE TLLR_OT SET ESTADO = 'V' ,"
'        gstrSql = gstrSql & "Fecha_Activacion = '" & CDate(pckFechaAtencion.Value) & "' , "
'kjcv 28.05.13 Graba la fecha en que se genera la activacion
        gstrSql = gstrSql & "Fecha_Activacion = '" & CDate(Now) & "' , "
        'kjcv 06.06.16
        gstrSql = gstrSql & "Usr_Activacion = '" & gUsr_Activacion & "' ,"
        gstrSql = gstrSql & "Fecha_Activa = '" & CDate(Now) & "' , "
        
        gstrSql = gstrSql & "Quien_Activa = '" & gstrIdUsuario & "' "
        gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_OT.Id_OT = '" & lblNroRecepcion & "' AND Tllr_OT.Seccion_OT = '" & gstrSeccion & "' "
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
'kjcv 10.02.14
'Validacion si tiene perfil para anular OT, se creo nuevo perfil desde BD opcion_sistema

    If Not Atributos("Glbl", "Tllr_20_0170", False, False, False, False) Then
        MsgBox "Ud. No cuenta con Acceso para realizar esta operación...", vbInformation, "Advertencia"
'        Unload Me
        Exit Sub
    End If
'kjcv 28.04.15 se comenta diferencia del total- insumo
    'stbTotalOT.Panels(2) = CDbl(stbTotalOT.Panels(2)) - gcurInsumo
    If stbTotalOT.Panels(2) <= 0 Then  ' valida que no existan valores cargados a la OT

        If VeriLiq() = True Then
            gstrSql = "UPDATE TLLR_OT SET ESTADO = 'N' ,"
'            gstrSql = gstrSql & "Fecha_Anulacion = '" & CDate(pckFechaAtencion.Value) & "' , "
            'kjcv 28.05.13 Graba la fecha en que se genera la activacion
            gstrSql = gstrSql & "Fecha_Anulacion = '" & CDate(Now) & "' , "
            gstrSql = gstrSql & "Quien_Anula = '" & gstrIdUsuario & "' "
            gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_OT.Id_OT = '" & lblNroRecepcion & "' AND Tllr_OT.Seccion_OT = '" & gstrSeccion & "' "
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
    Else
        MsgBox "No puede Anular una OT que Tenga Valor mayor que 0", vbExclamation, "Anular OT"
    End If
ElseIf ModeAction = gOTLiquidar And SW = 1 Then

    Dim lcurInsumos As Double
    
    'guardo el parametro, porque mas adelante si lo cambia lo hace en la variable global
    lcurInsumos = gcurInsumo
    
    If ValidaDatos = False Then
        Exit Sub
    End If
    Act = 0
    
    
    frmLiquidacion.Show 1
    If gblnCierraLiq = True Then
        GrabarRegistro
        If VeriLiq() = True And gflag = True Then
            EliminaRegistros gstrIdEmpresa, gstrIdSucursal, lblNroRecepcion, gstrSeccion
            gstrSql = "UPDATE TLLR_OT SET ESTADO = 'L' ,"
            gstrSql = gstrSql & "Fecha_Liquidacion = '" & CDate(Format(Now, "dd/mm/yyyy")) & "' , "
            gstrSql = gstrSql & "Quien_Liquida = '" & gstrIdUsuario & "' ,"
            gstrSql = gstrSql & "Total_Insumos=" & gcurInsumo & " ,"
            gstrSql = gstrSql & "Total_Materiales=" & gcurMateriales & " ,"
            gstrSql = gstrSql & "Total_Iva=" & Round(gcurTotalIVA, gintDecimalesMoneda) & " ,"
            gstrSql = gstrSql & "Total_OT_IVA=" & Round(gcurTotalNetoMasIVA, gintDecimalesMoneda) & " ,"
            gstrSql = gstrSql & "Total_OT=" & Round(gcurTotalNeto, gintDecimalesMoneda) & " "
            gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_OT.Id_OT = '" & lblNroRecepcion & "' AND Tllr_OT.Seccion_OT = '" & gstrSeccion & "' "
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
                
                'vuelve al estado original del parametro
                gcurInsumo = lcurInsumos
                
            End If
            MsgBox "La OT Nº " & lblNroRecepcion & " Fue Liquidada"
            Bloqueo "L"
        Else
            MsgBox "Lo siento, La Contraseña Ingresada no es la Correcta"
        End If
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
        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoMateriales Then '"85"
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
        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoInsumos Then '"80"
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
        If Trim(.SelectedItem.SubItems(9)) = gstrCodigoLubricantes Then '"90"
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
    itmAux.SubItems(5) = FormatoValor(IIf(txtHorasCar <> "", txtHorasCar, 0), "", gintDecimalesMoneda)
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
        If Me.lvwServiciosCarroceria.SelectedItem.SubItems(17) = "N" Then
            lvwServiciosCarroceria.ListItems.Remove lvwServiciosCarroceria.SelectedItem.Index
        End If
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
        gcurMateriales = gcurMateriales + CalculoMateriales(8)
        gcurLubricantes = CalculoLubricantes(8)
        Resta = CalculoInsumos(8) + gcurLubricantes
        .Panels(2).Text = FormatoValor(TotalSeccion(lvwRepuestos, 8) - Resta, "", gintDecimalesMoneda)
        stbTotalMateriales.Panels(2).Text = FormatoValor(gcurMateriales, "", gintDecimalesMoneda)   '// sumo insumos a materiales
        StbLubricantes.Panels(2).Text = FormatoValor(gcurLubricantes, "", gintDecimalesMoneda)
        stbInsumos.Panels(2).Text = FormatoValor(gcurInsumo + CalculoInsumos(8), "", gintDecimalesMoneda)
        
        'stbTotalMateriales.Visible = IIf(gcurMateriales > 0, True, False)
        
        
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
    .stbInsumos.Panels(2).Text = "0"
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
    'dblSemiTotal = dblSemiTotal + IIf(Not IsNull(gcurInsumo), gcurInsumo, 0)
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbInsumos.Panels(2).Text, ""))
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
    mstrSQL = mstrSQL & " AND Glbl_Cliente_Proveedor.Id_Ciudad = Glbl_Comuna.Id_Ciudad "
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

Function ObtenerValorMateriales() As Double
'Para obtenber el valor de materiales configurado
Dim Valor As Double
Valor = 0

    mstrSQL = "SELECT Materiales FROM Tllr_Parametro "
    mstrSQL = mstrSQL & " Where Id_Empresa='" & gstrIdEmpresa & "'"
    mstrSQL = mstrSQL & " and Id_Sucursal='" & gstrIdSucursal & "'"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                Valor = ValorNulo(!Materiales)
            End If
        End With
    End If
    Conexion.CloseHost AdoPrincipal


ObtenerValorMateriales = Valor

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

mstrSQL = "Exec Tllr_CargaInventario_Ot " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdRecepcion & "'"

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

Sub FillCampanaOT(strIdEmpresa As String, strIdSucursal As String, strIdRecepcion As String, strSeccion As String)

SetCheckOff lvwCampana

mstrSQL = "Exec Tllr_CargaCampana_Ot " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdRecepcion & "'"

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        While Not .EOF
            Set lvwCampana.SelectedItem = lvwCampana.FindItem(CStr(!Codigo), , , 1)
            lvwCampana.SelectedItem.Checked = True
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub

Sub FillMecanicaOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
    
    lvwServiciosMecanica.ListItems.Clear

    If gstrServiciosMarca = "S" Then
        mstrSQL = "Exec Tllr_CargaServicios_Mecanica_MM " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
    Else
        mstrSQL = "Exec Tllr_CargaServicios_Mecanica " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
    End If
    Screen.MousePointer = 11
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
            If ValorNulo(!Facturado) = "N" Then
                mblnOtFacturada = True
            End If
            itmAux.SubItems(13) = ValorNulo(!HorasReales)
            itmAux.SubItems(14) = ValorNulo(!Id_tarea)
            itmAux.SubItems(15) = ValorNulo(!estado_tarea)
            itmAux.SubItems(16) = ValorNulo(!id_grupo_centro_costo)
            
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub
Sub FillRepuestosReservados(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String, strTipo As String)
If strTipo <> "Q" Then
    lvwRepuestosMantencion.ListItems.Clear
End If

mstrSQL = "SELECT Tllr_Repuestos_Reservados.Id_Item, "
mstrSQL = mstrSQL & "Stck_Item.Descripcion, Stck_Item.Id_Familia,Tllr_Repuestos_Reservados.Solicitado, "
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
mstrSQL = mstrSQL & " (Tllr_Repuestos_Reservados.Seccion_OT = '" & strSeccion & "') AND"
If strTipo <> "Q" Then
    mstrSQL = mstrSQL & " (Tllr_Repuestos_Reservados.Tipo <> 'Q')"
Else
    mstrSQL = mstrSQL & " (Tllr_Repuestos_Reservados.Tipo = 'Q')"
End If


If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = frmRecepcion.lvwRepuestosMantencion.FindItem(!Id_Item, lvwText, , 0)
            If itmAux Is Nothing Then   ' Si no hay coincidencia
                Set itmAux = lvwRepuestosMantencion.ListItems.Add(, , ValorNulo(!Id_Item))
                Set lvwRepuestosMantencion.SelectedItem = itmAux
                itmAux.SubItems(1) = ValorNulo(!Descripcion)
                itmAux.SubItems(2) = FormatoValor(!Solicitado, "", 2)
                itmAux.SubItems(3) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
                itmAux.SubItems(4) = ValorNulo(!Familia)
                
                If Me.lvwServiciosMecanica.ListItems.Count > 0 Then
                    itmAux.SubItems(5) = Me.lvwServiciosMecanica.SelectedItem.SubItems(6)
                Else
                    itmAux.SubItems(5) = gstrIdCargo
                End If
                
                If !estado = "S" Then
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HFF0000
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HFF0000
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HFF0000
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HFF0000
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HFF0000
                   ' lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HFF0000
                End If
                If !estado = "P" Then
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HC0&
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
                    lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
                   ' lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
                End If
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
            itmAux.SubItems(2) = FormatoValor(!Solicitado, "", 2)
            itmAux.SubItems(3) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
            itmAux.SubItems(4) = ValorNulo(!Familia)
            'itmAux.SubItems(5) = lvwServiciosMecanica.SelectedItem.SubItems(6)
            
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
            lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
            'lvwRepuestosMantencion.ListItems(Me.lvwRepuestosMantencion.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
                
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub


Sub FillCarroceriaOT(strIdEmpresa As String, strIdSucursal As String, strIdRecepcion As String, strSeccion As String, strIdCiaSeguro As String)

lvwServiciosCarroceria.ListItems.Clear

mstrSQL = "Exec Tllr_CargaServicios_Carroceria " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdRecepcion & "'"

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
            itmAux.SubItems(12) = IIf(IsNull(!CARGO), "", !CARGO)
            itmAux.SubItems(13) = !IDCARGO
            itmAux.SubItems(14) = IIf(ValorNulo(!Provee) = "", "(Ninguno)", !Provee)
            itmAux.SubItems(15) = ValorNulo(!IDPROV)
            itmAux.SubItems(16) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
            itmAux.SubItems(17) = ValorNulo(!Facturado)
            itmAux.SubItems(18) = IIf(IsNull(!Codigo), 1, !Codigo)
            If ValorNulo(!Facturado) = "N" Then
                mblnOtFacturada = True
            End If
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Sub
Sub FillOtrosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)

lvwOtrosServicios.ListItems.Clear

mstrSQL = "Exec Tllr_CargaServicios_Otro " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwOtrosServicios.ListItems.Add(, , !ID)            '///des concepto
            itmAux.SubItems(1) = !Des                                              '///id concepto
            itmAux.SubItems(2) = FormatoValor(!TIEMPO, "", 2)                                                 '///d_p
            itmAux.SubItems(3) = FormatoValor(!UNITARIO, "", gintDecimalesMoneda)                                               '/// des parte)
            itmAux.SubItems(4) = FormatoValor(!PORCDESC, "", 2)                                 '///valor definido Format(ValorNulo(!HORAS), "#0.0")
            itmAux.SubItems(5) = FormatoValor(!MTODESC, "", gintDecimalesMoneda)
            itmAux.SubItems(6) = !IDCARGO
            itmAux.SubItems(7) = TraeCargoDes(!IDCARGO)
            itmAux.SubItems(8) = ValorNulo(!idmec)
            itmAux.SubItems(9) = MecanicoD(ValorNulo(!idmec))
            itmAux.SubItems(10) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
            itmAux.SubItems(11) = ValorNulo(!Facturado)
            If ValorNulo(!Facturado) = "N" Then
                mblnOtFacturada = True
            End If
            itmAux.SubItems(12) = ValorNulo(!HorasReales)
            itmAux.SubItems(13) = ValorNulo(!Id_tarea)
            itmAux.SubItems(14) = ValorNulo(!estado_tarea)
            'kjcv 15.09.17
            itmAux.SubItems(16) = ValorNulo(!fecha_update)
            itmAux.SubItems(15) = ValorNulo(!id_grupo_centro_costo)
            itmAux.SubItems(17) = ValorNulo(!HorasAsignadas)
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoPrincipal

End Sub


Sub FillTercerosOT(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)

lvwServiciosTerceros.ListItems.Clear

mstrSQL = "Exec Tllr_CargaServicios_Terceros " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"

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
            itmAux.SubItems(16) = ValorNulo(!id_grupo_centro_costo)
            
            If ValorNulo(!Facturado) = "N" Then
                mblnOtFacturada = True
            End If
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

mstrSQL = "Exec Tllr_CargaServicios_Repuestos " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            'If !CanTY > 0 Then  '///valores > 0
                Set itmAux = lvwRepuestos.ListItems.Add(, , !ID)            '///des concepto
                itmAux.SubItems(1) = ValorNulo(!Item)                                              '///id concepto
                itmAux.SubItems(2) = FormatoValor(!CANTY, "", 2)
                itmAux.SubItems(3) = FormatoValor(!Valor, "", gintDecimalesMoneda)
                itmAux.SubItems(4) = FormatoValor(!PORCDES, "", 2)
                itmAux.SubItems(5) = FormatoValor(!MTODES, "", gintDecimalesMoneda)
                itmAux.SubItems(6) = TraeCargoDes(ValorNulo(!IDCARGO))
                itmAux.SubItems(7) = ValorNulo(!IDCARGO)
                itmAux.SubItems(8) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
                itmAux.SubItems(9) = FamiliaRep(!ID)
                itmAux.SubItems(10) = ValorNulo(!Facturado)
                itmAux.SubItems(11) = IIf(IsNull(!Consumo), "STOCK", IIf(!Consumo = "C", "STOCK", "PRESUPUESTO"))
                '//LREYES
                itmAux.SubItems(12) = FormatoValor(0, "", 0)
                itmAux.SubItems(13) = FormatoValor(IIf(IsNull(!realy), 0, !realy), "", 2)
                itmAux.SubItems(14) = ValorNulo(!id_grupo_centro_costo)
                'kjcv 18.03.16
                itmAux.SubItems(15) = !PrecioVentaD
                
                If ValorNulo(!Facturado) = "N" Then
                    mblnOtFacturada = True
                End If
            'End If
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
'                lblChasis = ValorNulo(!chasis)
                txtChasis = ValorNulo(!chasis)
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

Sub FillCampanas()
mstrSQL = "SELECT Id_Promo AS Codigo, Descripcion AS Nombre FROM Promocion WHERE Vigencia = 'S' and id_Empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "' Order By Id_Promo"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set itmAux = lvwCampana.ListItems.Add(, , !Codigo)
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
'                    mstrSQL = mstrSQL & " SubTotal,Facturado,Porcentaje_Recargo,Monto_Recargo,Id_Proveedor,Descripcion,Id_Servicio_Carroceria)"
                    'kjcv 21.05.15
                    ' se agrega codigo SUNAT
                    mstrSQL = mstrSQL & " SubTotal,Facturado,Porcentaje_Recargo,Monto_Recargo,Id_Proveedor,Descripcion,ID_GRUPO_CENTRO_COSTO,CodProducto , Id_Servicio_Carroceria)"
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
                    'kjcv 21.05.15
                    mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(19)) & "',"
                    'kjcv 19.11.19
                    mstrSQL = mstrSQL & " '7818151',"
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
'                mstrSQL = mstrSQL & " Porcentaje_Dscto, Monto_Dscto)"
                'kjcv 21.05.15
                mstrSQL = mstrSQL & " Porcentaje_Dscto,ID_GRUPO_CENTRO_COSTO, Monto_Dscto)"
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(2) & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem) & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(14)) & "', "
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(6), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(9), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(3) & "', "
                mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(4) & "', "
'                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(17) & "', "
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(12), "#####0.00"))) & ",'" & .SelectedItem.SubItems(15) & "',"
'                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "#####0.00"))) & ")"
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & .SelectedItem.SubItems(16) & "'," & CCur(Val(Format(.SelectedItem.SubItems(11), "#####0.00"))) & ")"
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
    mstrNombreTabla = "Tllr_Otro_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Otro_Presupuesto"
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
'                mstrSQL = mstrSQL & " SubTotal,Descripcion_Otro,Facturado,HorasReales,Id_Tarea,Estado_Tarea)"
                'kjcv 21.05.15
                'mstrSql = mstrSql & " SubTotal,Descripcion_Otro,Facturado,HorasReales,Id_Tarea,ID_GRUPO_CENTRO_COSTO,Estado_Tarea)"
                'kjcv 15.09.17
                '19.11.19 se agrego codigo de SUNAT
                mstrSQL = mstrSQL & " SubTotal,Descripcion_Otro,Facturado,HorasReales,HorasAsignadas,Id_Tarea,ID_GRUPO_CENTRO_COSTO,fecha_update,CodProducto,Estado_Tarea)"
                '
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(6)) & "', "
                mstrSQL = mstrSQL & " '" & IIf(Trim(.SelectedItem.SubItems(8)) = "", "SIN", Trim(.SelectedItem.SubItems(8))) & "', "
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & UCase(Trim(.SelectedItem.SubItems(1))) & "','" & UCase(Trim(.SelectedItem.SubItems(11))) & "',"
                If .SelectedItem.SubItems(12) = "" Then
                    mstrSQL = mstrSQL & " " & 0 & ","
                Else
                    mstrSQL = mstrSQL & " " & CDbl(.SelectedItem.SubItems(12)) & ","
                End If
                If .SelectedItem.SubItems(17) = "" Then
                    mstrSQL = mstrSQL & " " & 0 & ","
                Else
                    mstrSQL = mstrSQL & " " & CDbl(.SelectedItem.SubItems(17)) & ","
                End If
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(13)) & "',"
                  'kjcv 21.05.15
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(15)) & "',"
                'kjcv 15.09.17
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(16)) & "',"
                'kjcv 19.11.19
                'codigo SUNAT
                 mstrSQL = mstrSQL & " '7818157',"

                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(14)) & "')"
                
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
Dim adoTemp As New ADODB.Recordset
Dim j As Integer

'valida si los repuestos no han sido devueltos
'y no ha sido refrescada la pantalla
If gstrProcedencia = "Movimientos" Then
    j = Me.lvwRepuestos.ListItems.Count
    For intIndice = 1 To Me.lvwRepuestos.ListItems.Count
        If Me.lvwRepuestos.ListItems(j).SubItems(11) = "STOCK" Then
            mstrSQL = "Select count(id_item) as Cuenta from Tllr_Repuestos_Ot WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' And Consumo='C' and id_item='" & Me.lvwRepuestos.ListItems(j) & "'"
            If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If adoTemp!cuenta = 0 Then
                    lvwRepuestos.ListItems.Remove (j)
                End If
            End If
        End If
        j = j - 1
    Next
End If

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Repuestos_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Repuestos_Presupuesto"
End If

GuardaRepuestos = True

'elimina solo si son presupuestos
mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & strIdDocumento & "' AND Seccion_OT ='" & strSeccion & "' And Consumo='P'"
Conexion.SendHost mstrSQL, , , , gcTiempoEspera

With lvwRepuestos
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            If VerificaRepuesto(.SelectedItem, lblNroRecepcion, strSeccion, mstrNombreTabla) = True Then
                mstrSQL = "UPDATE " & mstrNombreTabla
                mstrSQL = mstrSQL & " SET Id_Tipo_Cargo='" & Trim(.SelectedItem.SubItems(7)) & "',"
                mstrSQL = mstrSQL & " Cantidad = " & CDbl(Val(Format(.SelectedItem.SubItems(13), "#####0.00"))) & ", "
                mstrSQL = mstrSQL & " Valor = " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " cantidad_real = " & CCur(Val(Format(.SelectedItem.SubItems(13), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " Porcentaje_Descuento = " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " Monto_Descuento = " & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " SubTotal = " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " Facturado = " & UCase(Trim(IIf(.SelectedItem.SubItems(10) = "", "'N'", "'" & .SelectedItem.SubItems(10) & "'"))) & ","
                mstrSQL = mstrSQL & " Consumo = '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "',"
                mstrSQL = mstrSQL & " ID_GRUPO_CENTRO_COSTO = '" & .SelectedItem.SubItems(14) & "',"
                'kjcv 05.02.16
                mstrSQL = mstrSQL & " PrecioVentaD = '" & Round(.SelectedItem.SubItems(15), 2) & "',"
                mstrSQL = mstrSQL & " Saldo = '" & .SelectedItem.SubItems(12) & "'"
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
                'kjcv 04.12.2019
'                mstrSql = mstrSql & " Id_Linea, "
                
                mstrSQL = mstrSQL & " Cantidad, Valor,"
                mstrSQL = mstrSQL & " Porcentaje_Descuento,Monto_Descuento,"
'                mstrSql = mstrSql & " SubTotal,Facturado,Consumo,Saldo)"
                'kjcv 27.02.13
'                mstrSQL = mstrSQL & " SubTotal,Facturado,Consumo,Saldo,precioventaD)"
                'kjcv 21.05.15
                mstrSQL = mstrSQL & " SubTotal,Facturado,Consumo,Saldo,ID_GRUPO_CENTRO_COSTO, precioventaD)"
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(7)) & "', "
                'kjcv 04.12.19
'                mstrSql = mstrSql & "  & intIndice & , "
                
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ",'" & .SelectedItem.SubItems(10) & "',"
                mstrSQL = mstrSQL & " '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "',"
'                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(12) & "')"
  'kjcv 21.05.15
                mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(14) & "',"
                'kjcv 27.02.13
                 mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(12) & "',"
'                mstrSql = mstrSql & Retorna_Valor_General("Select Precio_Venta From Stck_Item Where Id_Item = '" & .SelectedItem & "'") & ")"
'kjcv 05.02.16
                mstrSQL = mstrSQL & Round(.SelectedItem.SubItems(15), 2) & ")"
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
'kjcv 28.09.21

Private Function GuardaCampana(strIdDocumento As String, strSeccion As String, gParametro As gcParametro) As Boolean
Dim mstrNombreTabla As String

If gParametro = gcOrdenTrabajo Then
    mstrNombreTabla = "Tllr_Campana_OT"
ElseIf gParametro = gcPresupuesto Then
    mstrNombreTabla = "Tllr_Campana_Presupuesto"
End If


GuardaCampana = True
mstrSQL = "DELETE " & mstrNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND " & mstrNombreTabla & ".ID_OT='" & strIdDocumento & "' and " & mstrNombreTabla & ".Seccion_OT = '" & strSeccion & "'"
If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
    For intIndice = 1 To lvwCampana.ListItems.Count
        Set lvwCampana.SelectedItem = lvwCampana.ListItems(intIndice)
        If lvwCampana.SelectedItem.Checked = True Then
            mstrSQL = "Insert Into " & mstrNombreTabla
            mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,Id_Promo, Id_OT, Seccion_OT) "
            mstrSQL = mstrSQL & " values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "','" & lvwCampana.SelectedItem & "', '" & strIdDocumento & "', '" & strSeccion & "' )"
            If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                GuardaCampana = False
                Exit Function
            End If
        End If
    Next
    GuardaCampana = True
Else
    GuardaCampana = False
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
'            mstrSQL = mstrSQL & " SubTotal, Facturado,HorasReales,Id_Tarea,Estado_Tarea)"
            'kjcv 21.05.15
            mstrSQL = mstrSQL & " SubTotal, Facturado,HorasReales,Id_Tarea,ID_GRUPO_CENTRO_COSTO,Estado_Tarea)"
            mstrSQL = mstrSQL & " Values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
            mstrSQL = mstrSQL & " '" & strIdDocumento & "', '" & gstrSeccion & "',"
            mstrSQL = mstrSQL & " '" & Trim(lblIdMarca) & "','" & Trim(lblIdModelo) & "',"
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem) & "',"
            mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(6) & "'," & IIf(.SelectedItem.SubItems(8) = "", "NULL", " '" & .SelectedItem.SubItems(8) & "' ") & ", "
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & " , " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & " , "
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & " ," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
            mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & .SelectedItem.SubItems(11) & "',"
            If .SelectedItem.SubItems(13) = "" Then
                mstrSQL = mstrSQL & " " & 0 & ","
            Else
                mstrSQL = mstrSQL & " " & CDbl(.SelectedItem.SubItems(13)) & ","
            End If
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(14)) & "',"
            'kjcv 21.05.15
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(16)) & "',"
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(15)) & "')"
                
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
mstrSQL = "SELECT Top 1 Id_OT, "
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
mstrSQL = mstrSQL & " Kilometros_Recepcion, "
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
mstrSQL = mstrSQL & " Id_Presupuesto, "
mstrSQL = mstrSQL & " Fecha_Liquidacion, "
mstrSQL = mstrSQL & " OrdenReparacion, "
mstrSQL = mstrSQL & " Nro_Presupuesto_Origen, "
mstrSQL = mstrSQL & " NroReferencia, Bencina , "
mstrSQL = mstrSQL & " CorrelativoSpiga  "
'kjcv 17.04.18
mstrSQL = mstrSQL & " ,Cuponera, Nro_Cupon ,Id_Promo, Id_Trabajo ,Id_Tipo_Venta "
mstrSQL = mstrSQL & " ,PDI"
'wcs 16.04.2024
mstrSQL = mstrSQL & " ,Correo"
mstrSQL = mstrSQL & " ,Telefono"
mstrSQL = mstrSQL & " From Tllr_OT"
letSql = mstrSQL & " " & strWhere & " " & strOrder

End Function

Private Sub LeerCampos()

'/// inicializa variable para verificar si la ot esta totalmente facturada
mblnOtFacturada = False

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
    If !ReparacionMantencion = "M" Then
        Me.optMantencion.Value = True
    Else
        Me.optReparacion.Value = True
    End If
    If !Estado_Reserva = "R" Then
        Me.cmdReserva.Enabled = False
        Me.cmdAnularReserva.Enabled = True
    Else
        Me.cmdReserva.Enabled = True
        Me.cmdAnularReserva.Enabled = False
    End If
    lblNroRecepcion.Text = !Id_OT
    mstrIdPresupuestoOrigen = ValorNulo(!Id_Presupuesto)
    lblPresupuesto = ValorNulo(!Nro_Presupuesto_Origen)
    lblFechaLiquidacion = IIf(!estado <> "N" Or !estado <> "V", ValorNulo(!Fecha_Liquidacion), "")
    dtcGarantia.BoundText = !TipoOt
    dtcPromocion.BoundText = ValorNulo(!Id_Promo)
    dtcTrabajo.BoundText = ValorNulo(!Id_Trabajo)
'    dbcboTipoVenta.BoundText = ValorNulo(!id_tipo_Venta)
    txtNReferencia = ValorNulo(!NroReferencia)
    Me.cmbBencina.ListIndex = IIf(IsNull(!Bencina), -1, !Bencina)
    'kjcv 17.04.18
    Me.cmbCuponera.ListIndex = IIf(IsNull(!Cuponera), -1, !Cuponera)
    Me.txtNroCupon = ValorNulo(!Nro_Cupon)
    If !TipoOt = "PRE" Then
        dtcGarantia.Enabled = False
    Else
        dtcGarantia.Enabled = True
    End If
    gstrIdCargo = TraeCargo(!TipoOt)
    dtcTipoCono.BoundText = !Id_Tipo_Cono
    dtcRecepcionista.BoundText = !RealizadoPor
    txtNroCono = !Nro_Cono
    
    pckFechaAtencion.Value = !Fecha_Emision
    'jn 17.01.2024
    lblHoraAtencion = Format$(!Fecha_Emision, "HH:mm:ss")
    pckFechaEntrega.Value = !Entrega_Estimada
    cboHora.Text = ValorNulo(!Hora_Entrega)
    
    txtNroSiniestro = ValorNulo(!Nro_Siniestro)
    txtNroPoliza = ValorNulo(!Nro_Poliza)
    txtLiquidador = ValorNulo(!Liquidador)
    txtOrdenReparacion = ValorNulo(!OrdenReparacion)
    
    txtDeducibleUF = !Deducible_UF
    txtDeduciblePesos = !deducible_pesos
    lblCompañia.Tag = !Id_Compañia_Seguro
    gstrIdCompañiaSeg = !Id_Compañia_Seguro
    lblCompañia = CiaSegDes(!Id_Compañia_Seguro)
    
    txtComentario = !Comentario
    txtCorreSpiga = ValorNulo(!CorrelativoSpiga)
    txtPatente = ValorNulo(!Patente)
    txtFolioGarantia = !Folio_Garantia
    txtSolicita = !Solicitado_Por
    gcurInsumo = !Total_Insumos
    gcurMateriales = !Total_Materiales
    
    'wcs 16.06.2024
    txtCorreo.Text = ValorNulo(!Correo)
    txtTelefono.Text = ValorNulo(!Telefono)
    'gcurMateriales = !Total_Materiales
    'stbInsumos.Panels(2).Text = FormatoValor(!Total_Insumos, "", 0)
    'kjcv 12.11.13 para Bloquear el Buscar Placa
    tlbPatente.Buttons(2).Enabled = False
    If Not IsNull(!estado) Then
        If gstrProcedencia = "Movimientos" Then
            lblEstadoOTValor.Caption = IIf(!estado = "V", "VIGENTE", IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "N", "NULA", IIf(!estado = "F" Or !estado = "B", "EMITIDA", IIf(!estado = "R", "RESERVA", IIf(!estado = "P", "PRESUPUESTO", ""))))))
            'kjcv 24 10.13
            txtTipo.Text = IIf(!PDI = "S", "PDI", "")
            tlbBarraHerramientas.Buttons.Item(2).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", True, IIf(!estado = "R", True, IIf(!estado = "P", True, False))))))
            tlbBarraHerramientas.Buttons.Item(13).Enabled = IIf(!estado = "V", False, IIf(!estado = "L", True, IIf(!estado = "N", True, IIf(!estado = "F" Or !estado = "B", False, False))))    'ACTIVAR
            tlbBarraHerramientas.Buttons.Item(14).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, False))))    'ANULAR
            tlbBarraHerramientas.Buttons.Item(15).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", True, False))))    'LIQUIDAR
            tlbBarraHerramientas.Buttons.Item(20).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Separador
            tlbBarraHerramientas.Buttons.Item(21).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Confirmar Reserva
            tlbBarraHerramientas.Buttons.Item(22).Visible = IIf(!estado = "V", False, IIf(!estado = "L", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, IIf(!estado = "R", True, False))))) 'Eliminar Reserva
            tlbBarraHerramientas.Buttons.Item(24).Visible = IIf(!estado = "P", True, False) 'Liquidar presupuesto
            tlbBarraHerramientas.Buttons.Item(25).Visible = IIf(!estado = "P", True, False) 'Liquidar presupuesto
        Else
'            tlbBarraHerramientas.Buttons.item(2).Enabled = False
'kjcv 27.02.13
            tlbBarraHerramientas.Buttons.Item(2).Enabled = True
        End If
    End If
    
    'busca numeros de documentos asociados
    lblDocumentos = IIf(!estado = "F" Or !estado = "B", NumerosDocumentos(!Id_OT, gstrSeccion), "")
    
    If ValorNulo(!Patente) <> "" Then DatosVehiculo !Patente
    txtKilAct = !Kilometros_Recepcion 'trae los kilometros de la OT
    '/////////////////////////////////////////////////////////////////////////////////
    FillConceptosVsCiaSeguro dtcConceptos, datConceptos, lblCompañia.Tag
    '/////////////////////////////////////////////////////////////////////////////////
    FillInventarioOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    FillCampanaOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    '/////////////////////////////////////////////////////////////////////////////////
    FillMecanicaOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    AsignaTotal mcFichaMecanica, stbTotalMec
    
    '/////////////////////////////////////////////////////////////////////////////////
    'If !Seccion_OT = "C" Then
        FillCarroceriaOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion, lblCompañia.Tag
        AsignaTotal mcFichaCarroceria, stbTotalCarroceria
    'Else
    '    lvwServiciosCarroceria.ListItems.Clear
    '    frmRecepcion.stbTotalCarroceria.Panels(2).Text = 0
    'End If
    '/////////////////////////////////////////////////////////////////////////////////
    FillOtrosOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    AsignaTotal mcFichaOtros, stbTotalOtros
    '/////////////////////////////////////////////////////////////////////////////////
    FillTercerosOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    AsignaTotal mcFichaTerceros, stbTotalTerceros
    '/////////////////////////////////////////////////////////////////////////////////
    
    'wcs 13/07/2024
    gcurMateriales = !Total_Materiales
    'stbTotalMateriales.Visible = IIf(gcurMateriales > 0, True, False)
    
    FillRepuestosOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    AsignaTotal mcFichaRepuestos, stbTotalRepuestos
'    stbTotalMateriales.Panels(2).Text = Format(CalculoMateriales(8))
    '/////////////////////////////////////////////////////////////////////////////////
    
    If !Estado_Reserva = "R" Then
        FillRepuestosReservados gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion, "T"  'tempario
        FillRepuestosFaltantes gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    Else
        '//// Si no encuentra reserva de repuestos busca los repuestos de los servicios
        Dim i As Integer
        lvwRepuestosMantencion.ListItems.Clear
        For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
            mstrAgregaPresupuesto = False
            Repuestos_de_la_Mantencion Me.lblIdMarca, Me.lblIdModelo, lvwServiciosMecanica.ListItems(i), IIf(Me.lvwServiciosMecanica.ListItems(i).SubItems(12) = "S", True, False)
        Next
        FillRepuestosReservados gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion, "Q"  'presupuesto
    End If
    
    TotalFinal
    '/////////////////////////////////////////////////////////////////////////////////
    If ValorNulo(!estado) = "B" Or ValorNulo(!estado) = "F" Then
        tlbBarraHerramientas.Buttons.Item(15).Enabled = mblnOtFacturada  'LIQUIDAR
        tlbBarraHerramientas.Buttons.Item(2).Enabled = mblnOtFacturada   'GUARDAR
        
    End If
    
    If !estado = "B" Or !estado = "F" Or !estado = "L" Then
        Me.stbSeguroTaller.Panels(2).Text = Retorna_Valor_General("Select sum(SeguroTaller) as Seguro from Tllr_Facturacion where id_ot='" & !Id_OT & "' And Seccion_OT='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
        If Me.stbSeguroTaller.Panels(2).Text = "" Then
            gcurSeguroTaller = 0
            Me.stbSeguroTaller.Panels(2).Text = "0"
        Else
            gcurSeguroTaller = CDbl(Me.stbSeguroTaller.Panels(2).Text)
        End If
    Else
        Me.stbSeguroTaller.Panels(2).Text = "0"
    End If

    gstrEstado = ValorNulo(!estado)
    
    Bloqueo ValorNulo(!estado)
    
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
Dim AdoAnular As New ADODB.Recordset
If Me.lvwRepuestosMantencion.ListItems.Count > 0 Then

    '/// valida que la reserva no haya pasado a Consumo
    EstadoReserva = Retorna_Valor_General("Select Estado_Reserva from Stck_Regularizacion Where Id_OT='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
    If EstadoReserva = "L" Then
        MsgBox "Esta Reserva ya paso a ser un Consumo...", vbInformation, "Anular Reserva de Repuestos"
        Exit Sub
    End If
    'Levanta listview con los repuestos de la mantencion
    If MsgBox(" Esta Seguro de Anular esta esta Reserva de Repuestos ", vbQuestion + vbYesNo, "Confirma Anulación") = vbYes Then
        
        mstrSQL = "Select Id_Regularizacion as Numero from Stck_Regularizacion where id_ot='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, AdoAnular, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            With AdoAnular
                If Not .BOF And Not .EOF Then
                    .MoveFirst
                    While Not .EOF
                        NroRegularizacion = !NUMERO
                        Call Actualiza_Saldos_VS_Detalle("S", "Select Canrtidad, Id_Empresa, Id_sucursal, Id_Bodega,Id_Ubicacion,Id_Item From Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & NroRegularizacion & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Empresa = '" & gstrIdEmpresa & "'")
                        
                        EliminaReservaRepuestos NroRegularizacion, lblNroRecepcion
                        
                        .MoveNext
                    Wend
                    '/// Actualiza estado de reserva
                    mstrSQL = "UPDATE TLLR_OT SET Estado_Reserva='N' "
                    mstrSQL = mstrSQL & "Where Id_OT='" & frmRecepcion.lblNroRecepcion & "' "
                    mstrSQL = mstrSQL & "And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Seccion_OT='" & gstrSeccion & "'"
                    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
                    DesactivaBotonAnularReserva
                End If
            End With
        End If
    Else
        Exit Sub
    End If
End If

End Sub
Sub DesactivaBotonAnularReserva()
    cmdAnularReserva.Enabled = False
    cmdReserva.Enabled = True
End Sub

Private Sub cmdConsultaSaldo_Click()
If Me.lvwRepuestosMantencion.ListItems.Count > 0 Then
    'Levanta listview con los repuestos de la mantencion
    gstrProcedencia = "Consulta"  'para que solo consulte y no reserve
    frmRepuestosReservados.Show vbModal
    gstrProcedencia = "Movimientos"  'vuelve al estado original
End If
End Sub

Private Sub cmdConsultaStock_Click()
    If Me.lvwRepuestos.ListItems.Count > 0 Then
        'Levanta listview con los repuestos del presupuesto
        frmRepuestosReservados.Show vbModal
        ActualizarSaldoRepuestos lblNroRecepcion, gstrSeccion
    End If
End Sub

Private Sub cmdImprimir_Click()
If Me.txtComentario.Text <> "" Then

    ImprimirConsulta

End If
End Sub

Sub ImprimirConsulta()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
Dim vNombreContacto As String
Dim vE_mail As String
Dim vTelefono As String
Dim vDireccion As String
Dim lstrQuery As String
Dim tbCliente As New ADODB.Recordset
Dim vCelularRecepcionista As String

Dim lstrArchivoIni As String
lstrArchivoIni = Command()
gstrPathReporte = LetConnectionString("TLLR", "RPT", lstrArchivoIni, 256)

    
    If Me.txtComentario.Text = "" Then
      MsgBox "No existen informacion en Comentarios", vbExclamation, "Imprimir"
      Exit Sub
    End If
    
    'jn 17.01.2024 obtención de datos adicionales de cliente y asesor
        lstrQuery = "SELECT ISNULL(NombreContacto,'') NombreContacto, ISNULL(E_mail,'') E_mail, ISNULL(Telefono,'') Telefono, ISNULL(Direccion,'') Direccion FROM Glbl_cliente_proveedor WHERE Id_Cliente_Proveedor ='" & lblIdCliente & "'"
    If Conexion.SendHost(lstrQuery, tbCliente, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        vNombreContacto = tbCliente!NombreContacto
        vE_mail = tbCliente!E_Mail
        vTelefono = tbCliente!Telefono
        vDireccion = tbCliente!Direccion
    End If
    Conexion.CloseHost tbCliente
    
    vCelularRecepcionista = Retorna_Valor_General("Select ISNULL(Movil,'') Movil From Tllr_Mecanicos where Id_Mecanico='" & Me.dtcRecepcionista.BoundText & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'")
    

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)
    If Dir(gstrPathReporte & "\BDInventario.mdb") <> "" Then Kill gstrPathReporte & "\BDInventario.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDInventario.mdb", dbLangGeneral) ' Crea a una base de datos nueva
   
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,FechaEmision Text,Recepcionista text,Seccion text,Kilometros text,Comentario memo,Patente Text,Cliente Text,Inventario text,Marca text,Modelo text, Año text, Color text, Contacto text, E_Mail text, Telefono text, Direccion text, CelularRecepcionista text, Fecha text, Hora text)"
    Dbsnueva.Execute "CREATE TABLE T_PARAINVENTARIO (NroOT text,Inventario text)"
    
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
         
        Tabla.AddNew
        Tabla!NroOT = IIf(lblNroRecepcion = "", " ", lblNroRecepcion)
        'Tabla!FechaEmision = IIf(pckFechaAtencion = "", "", pckFechaAtencion)
        Tabla!Fecha = Format(pckFechaAtencion, "dd/MM/yyyy")
        Tabla!Hora = Format$(lblHoraAtencion, "HH:mm:ss")
        Tabla!Recepcionista = ValorNulo(dtcRecepcionista.Text)
        Tabla!Seccion = "M"
        Tabla!Kilometros = txtKilAct.Text
        Tabla!Comentario = IIf(txtComentario.Text = "", " ", txtComentario.Text)
        Tabla!Patente = txtPatente.Text
        Tabla!Cliente = lblCliente.Caption
        Tabla!Marca = lblMarca.Caption
        Tabla!Modelo = lblModelo.Caption
        Tabla!Año = txtAño.Text
        Tabla!Color = lblColorE.Caption
        Tabla!Contacto = vNombreContacto
        Tabla!E_Mail = vE_mail
        Tabla!Telefono = vTelefono
        Tabla!Direccion = vDireccion
        Tabla!CelularRecepcionista = vCelularRecepcionista
        Tabla.Update
        Tabla.Close
        
        
     Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAINVENTARIO")
     Tabla.AddNew
        Tabla!NroOT = IIf(lblNroRecepcion = "", " ", lblNroRecepcion)
        Tabla.Update
     
     Dim cadItems As String
     Dim cadComa As String
     cadComa = ","
        
        For i = 1 To lvwInventario.ListItems.Count
            Set lvwInventario.SelectedItem = lvwInventario.ListItems(i)
            If lvwInventario.SelectedItem.Checked Then
               
                cadItems = cadItems & IIf(lvwInventario.SelectedItem.SubItems(1) = "", "", lvwInventario.SelectedItem.SubItems(1)) & cadComa
                
            End If
        Next i
        
        If Len(cadItems) > 0 Then cadItems = Left(cadItems, Len(cadItems) - 1)
        

        
        
'          For i = 1 To lvwInventario.ListItems.Count
'            Set lvwInventario.SelectedItem = lvwInventario.ListItems(i)
'            If lvwInventario.SelectedItem.Checked Then
'                Tabla.AddNew
'                Tabla!NroOT = IIf(lblNroRecepcion = "", " ", lblNroRecepcion)
'                Tabla!Inventario = IIf(lvwInventario.SelectedItem.SubItems(1) = "", "", lvwInventario.SelectedItem.SubItems(1))
'                Tabla.Update
'            End If
'        Next i
        
        
        
        Tabla.Close
        
   Dbsnueva.Close
  
  
  With crInventario
  
    Me.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
    Me.cdImpresora.CancelError = True
    Me.cdImpresora.Action = 5
                                      
    .CopiesToPrinter = cdImpresora.Copies
    .ReportFileName = gstrPathReporte & "\rptInventario.rpt"
  
    .Destination = crptToPrinter
    .WindowState = crptMaximized
    .DataFiles(0) = gstrPathReporte & "\BDInventario.mdb"
    
    .Formulas(0) = "invItems='" & cadItems & "'"
    
     .Action = True
     
  End With
   
'   With rptPatente
'        .ReportFileName = gstrPathReporte & "\Inventario.Rpt"
'        .WindowTitle = "Historico Por " & gstrNombrePatente
'        .WindowState = crptMaximized
'        .DataFiles(0) = gstrPathReporte & "\BDInventario.mdb"
'        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
'        .Formulas(1) = "TITULO='Inventario de " & gstrNombrePatente & "'"
'        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
'        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
'        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
'        .Formulas(7) = "TDecimal=" & gintDecimalesMoneda
'        .Formulas(8) = "TSigla='" & gstrMonedaLocal & "'"
'        .Formulas(9) = "NombrePatente='" & gstrNombrePatente & "'"
'
'        .Destination = crptToWindow
'        .Action = True
'   End With
   
   Screen.MousePointer = 1

End Sub


Private Sub cmdPrueba_Click()
ImprimirComentario (Me.txtComentario)
End Sub

Private Sub ImprimirComentario(Comentario As String)
Dim LARGO As Integer
Dim linea As String
Dim Blo80 As String
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim L As Integer
Dim Y As Integer
Dim x As Integer

Dim PosI As Integer
Dim posF As Integer

j = 20 'caracteres x linea
LARGO = Len(Comentario)
i = 1
Y = 8800
x = 3800
Dim Lineas As Variant
'Dim i As Integer

Lineas = Split(Comentario, vbCrLf)

For i = LBound(Lineas) To UBound(Lineas)
    Printer.CurrentX = x
    Printer.CurrentY = Y
    Printer.Print Lineas(i)
    Y = Y + 300
Next



'' PosI = InStr(1, Comentario, vbCrLf, 0)
''' linea = Mid(Comentario, i, PosI)
''    Printer.CurrentX = X
''    Printer.CurrentY = Y
''    Printer.Print linea
''
''
''
'' While i < LARGO
''
'' posF = InStr(i, Comentario, vbCrLf, 0)
'' If i = 1 Then
''  linea = Mid(Comentario, i, PosI)
''  Else
''
''    linea = Mid(Comentario, i, posF - PosI)
''
''
'' End If
'' Printer.CurrentX = X
''    Printer.CurrentY = Y
''    Printer.Print linea
''i = i + PosI
''Y = Y + 300
''
'' Wend
 

'While I < Largo
'Linea = Mid(Comentario, I, J)
'Printer.CurrentX = 3800
'Printer.CurrentY = Y
'Printer.Print Linea
'I = I + J
'Y = Y + 240
'Wend


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

Private Sub cmdTemparios_Click()
frmTemparios.Show
End Sub

Private Sub dtcConceptos_Change()
txtSeccion = TipoConcepto(dtcConceptos.BoundText)
End Sub
Private Sub dtcGarantia_Change()
mstrCargo = TraeCargo(dtcGarantia.BoundText)
TipoOt dtcGarantia.BoundText
gstrIdCargo = mstrCargo
 dtcTrabajo.Enabled = True
If dtcGarantia.BoundText = "PDI" Or dtcGarantia.BoundText = "GFB" Or dtcGarantia.BoundText = "RCL" Then
    dtcTrabajo.BoundText = "NIN"
    dtcTrabajo.Enabled = False
End If
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
Dim tbRegistros As New ADODB.Recordset
Dim lstrQuery As String

    mblnSW = True
    gstrSeccion = "M"
    stbServicios.tab = 0
    gstrKmsAutoNuevo = ""
    mstrLiquidaPresupuesto = False
    ' kjcv 02.02.22
    '// Tipo Venta Forma de Pago
    Set tbRegistros = New ADODB.Recordset
        lstrQuery = "SELECT * FROM Glbl_Tipo_Venta WHERE Vigencia = 'S' ORDER BY Descripcion"
 
    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        Set datTipoVenta.Recordset = tbRegistros
    End If
    
    
'    gcurMateriales = ObtenerValorMateriales()
    
    

    
    'gcurInsumoDef = gcurInsumo
End Sub

Private Sub Form_Resize()
''Dim ldblAncho As Double
''Dim ldblAnchoCol As Double
''Dim ldblAnchoBtnSmall As Double
''
''Screen.MousePointer = vbHourglass
'''kjcv 20-01-12
''ldblAncho = 120
''ldblAnchoBtnSmall = 240
'''
''Me.Frame8.Left = ldblAncho
''Me.Frame8.Width = Me.Frame8.Width
''
''Me.stbServicios.Left = ldblAncho
''Me.stbServicios.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0)
''
''Me.fmePat.Left = ldblAncho
''Me.fmePat.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''
''Me.fmeCia.Left = ldblAncho
''Me.fmeCia.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''
''Me.fmeInv.Width = IIf(Me.ScaleWidth / 2 - (ldblAncho * 2) >= 0, Me.ScaleWidth / 2 - (ldblAncho * 2), 0) - ldblAncho
''Me.lvwInventario.Width = Me.lvwInventario.Width
''
''Me.txtComentario.Width = IIf(Me.ScaleWidth / 2 - (ldblAncho * 2) >= 0, Me.ScaleWidth / 2 - (ldblAncho * 2), 0) - 4 * ldblAncho
''
''Me.fmeCom.Left = Me.fmeInv.Left + Me.fmeInv.Width + ldblAncho
''Me.fmeCom.Width = IIf(Me.ScaleWidth / 2 - (ldblAncho * 2) >= 0, Me.ScaleWidth / 2 - (ldblAncho * 2), 0)
''
''Me.fmeMec.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''
''Me.lvwServiciosMecanica.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''
''Me.lvwRepuestosMantencion.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''
''Me.fmeCar.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''Me.lvwServiciosCarroceria.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''
''Me.fmeOtr.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''Me.lvwOtrosServicios.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''
''Me.fmeTer.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''Me.lvwServiciosTerceros.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho
''
''Me.fmeRep.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 2 * ldblAncho
''
''Me.lvwRepuestos.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0) - 4 * ldblAncho

End Sub




Private Sub lblIdCliente_Change()
If DatosCliente(lblIdCliente) Then DoEvents
End Sub

Private Sub lblNroRecepcion_DblClick()
If gstrImpresion = "O" And Me.lblNroRecepcion <> "" Then
    gstrBusca = InputBox("Ingrese El Numero de O/T Deseado :", "Ir a....", CStr(Val(Mid(lblNroRecepcion, 6, Len(lblNroRecepcion) - 5))))
    gstrBusca = FormatOT(gstrBusca)
    If gstrBusca <> "" Then
'        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
'kjcv 02.01.13
mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT like  '%" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
        gstrSql = letSql(mstrWhere, mstrOrderBy)
        If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost AdoPrincipal
    End If
    Screen.MousePointer = vbDefault
    Me.SetFocus
End If
End Sub

Private Sub lblVin_Change()
VerificaCampañas
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
    If Me.lvwOtrosServicios.ListItems.Count > 0 Then
        Select Case Button
            Case vbRightButton  '//BOTON DERECHO
                gstrProcedenciaBotonDerecho = "Otros"
                frmMain.popup(5).Enabled = True
                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
        End Select
    End If
End If


'    Dim i As Integer
'    Dim gstrBusca As String
'
'    Select Case Button
'        Case vbRightButton  '//BOTON DERECHO
'            gstrBusca = InputBox("Ingrese el Descuento :", "Descuentos Multiples")
'            If IsNumeric(gstrBusca) Then
'                If Val(gstrBusca) >= 0 And Val(gstrBusca) < 101 Then
'                    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
'                        If Me.lvwOtrosServicios.ListItems(i).Selected Then
'                            dblTotalInicial = Round(CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(2)) * CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(3)), 2)
'                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(5) = FormatoValor(ValorPorcentaje(dblTotalInicial, Val(gstrBusca)), "", gintDecimalesMoneda)
'                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(10) = FormatoValor(dblTotalInicial - CDbl(Me.lvwOtrosServicios.ListItems.Item(i).SubItems(5)), "", gintDecimalesMoneda)
'                            Me.lvwOtrosServicios.ListItems.Item(i).SubItems(4) = FormatoValor(Val(gstrBusca), "", 2)
'                        End If
'
'                    Next
'                    AsignaTotal mcFichaOtros, stbTotalOtros
'                    TotalFinal
'                Else
'                    MsgBox "El Descuento fue mal Ingresado", vbExclamation, "Valor Descuento"
'                End If
'            Else
'                MsgBox "El valor debe Ser Numerico", vbExclamation, "Valor Descuento"
'            End If
'    End Select

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
    If Me.lvwRepuestos.ListItems.Count > 0 Then
        Select Case Button
            Case vbRightButton  '//BOTON DERECHO
                gstrProcedenciaBotonDerecho = "Repuestos"
                frmMain.popup(5).Enabled = False
                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
        End Select
    End If
End If
End Sub

Private Sub lvwRepuestosMantencion_DblClick()
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
    If Me.lvwServiciosCarroceria.ListItems.Count > 0 Then
        Select Case Button
            Case vbRightButton  '//BOTON DERECHO
                gstrProcedenciaBotonDerecho = "Carroceria"
                frmMain.popup(5).Enabled = False
                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
        End Select
    End If
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
    If Me.lvwServiciosMecanica.ListItems.Count > 0 Then
        Select Case Button
            Case vbRightButton  '//BOTON DERECHO
                gstrProcedenciaBotonDerecho = "Mecanica"
                frmMain.popup(5).Enabled = True
                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
        End Select
    End If
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
    If Me.lvwServiciosTerceros.ListItems.Count > 0 Then
        Select Case Button
            Case vbRightButton  '//BOTON DERECHO
                gstrProcedenciaBotonDerecho = "Terceros"
                frmMain.popup(5).Enabled = False
                PopupMenu frmMain.MenuPopup, , , , frmMain.popup(1)
        End Select
    End If
End If
End Sub

Private Sub optRecepcion_Click(Index As Integer)
Select Case Index
Case 0
    stbServicios.tab = 0
    gstrSeccion = "M"
    If Me.Tag = "" Then
        Renovar
    End If
    'stbServicios.TabEnabled(3) = False
    Screen.MousePointer = vbDefault
Case 1
    stbServicios.tab = 0
    gstrSeccion = "C"
    If Me.Tag = "" Then
        Renovar
    End If
    'stbServicios.TabEnabled(3) = True
    Screen.MousePointer = vbDefault
End Select
End Sub



Private Sub tlbAddRep_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Agregar" ' ////////////////AGREGAR
        If Trim(txtPatente.Text) <> "" Then
            mstrProcedenciaAux = gstrProcedencia
            gstrProcedencia = "Presupuestos"
            frmSelTempRepuestos.Show vbModal
            AsignaTotal mcFichaRepuestos, stbTotalRepuestos
            TotalFinal
            gstrProcedencia = mstrProcedenciaAux
        End If
    Case "Quitar" ' ////////////////QUITAR
        If Me.lvwRepuestos.ListItems.Count > 0 Then
            If Me.lvwRepuestos.SelectedItem.SubItems(11) = "PRESUPUESTO" Then
                If Not lvwRepuestos.SelectedItem Is Nothing Then
                    If AccesoEliminar(lvwRepuestos.SelectedItem) = True Then
                        lvwRepuestos.ListItems.Remove (lvwRepuestos.SelectedItem.Index)
                        AsignaTotal mcFichaRepuestos, stbTotalRepuestos
                        TotalFinal
                    Else
                        MsgBox ""
                    End If
                End If
            End If
        End If
    End Select
End Sub

Private Sub tlbAddServicioCar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case Is = "Agregar"
    If Trim(txtPatente) <> "" Then
        frmAddTrabajosCarroceria.Show vbModal
        AsignaTotal mcFichaCarroceria, stbTotalCarroceria
        TotalFinal
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
        mstrProcedenciaAux = gstrProcedencia
        gstrProcedencia = "Movimientos"
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
        gstrProcedencia = mstrProcedenciaAux
    Else
        MsgBox LoadResString(301), vbOKOnly, LoadResString(4)
    End If
Case Is = "Quitar"
    If (lvwServiciosMecanica.ListItems.Count > 0 And Me.cmdReserva.Enabled = True) Or Me.dtcGarantia.BoundText = "PRE" Then
        lstrServicioMecanica = lvwServiciosMecanica.SelectedItem
        'kjcv 11.09.12 Cambio de subitems 12 a subitems(11), no se podia quitar Servicio de Mecanica
        If Me.lvwServiciosMecanica.SelectedItem.SubItems(11) = "N" Then
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
        End If
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
            If Me.lvwOtrosServicios.SelectedItem.SubItems(11) = "N" Then
                lvwOtrosServicios.ListItems.Remove lvwOtrosServicios.SelectedItem.Index
                AsignaTotal mcFichaOtros, stbTotalOtros
                TotalFinal
            End If
        End If
    End If
End Select
End Sub

Private Sub tlbAddServicioTer_ButtonClick(ByVal Button As MSComctlLib.Button)

If Not Atributos("Glbl", "Tllr_20_0180", False, False, False, False) Then
        MsgBox "Ud. No cuenta con Acceso para realizar esta operación...", vbInformation, "Advertencia"
        Exit Sub
End If

Select Case Button.Key
Case "Agregar" ' ////////////////AGREGAR
    If Trim(txtPatente.Text) <> "" Then
        frmAddTrabajosTercero.Show vbModal
        AsignaTotal mcFichaTerceros, stbTotalTerceros
        TotalFinal
    End If
Case "Quitar" ' ////////////////QUITAR
    If Not lvwServiciosTerceros.SelectedItem Is Nothing Then
        If Me.lvwServiciosTerceros.SelectedItem.SubItems(15) = "N" Then
            If Mid(Me.lvwServiciosTerceros.SelectedItem, 1, 2) = "OC" Then
                MsgBox "No puede Eliminar este Item, porque fue registrado desde una Orden De Compra", vbInformation, "Advertencia"
            Else
                lvwServiciosTerceros.ListItems.Remove (lvwServiciosTerceros.SelectedItem.Index)
                AsignaTotal mcFichaTerceros, stbTotalTerceros
                TotalFinal
            End If
        End If
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
            gstrProcedenciaRptos = ""
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
        Case "Editar"
            frmHistoricoOT.Show
        Case "ValoresCargo"
            If Me.lblEstadoOTValor.Caption <> "VIGENTE" Then
                frmValoresPorCargo.Show vbModal
            Else
                MsgBox "La OT aún está Vigente"
            End If
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSW Then
        mstrProcedencia = gstrProcedencia
        mblnSW = False
        If mstrProcedencia = "Movimientos" Then
            If Not Atributos("Glbl", "Tllr_20_0020", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
                MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
                Unload Me
                Exit Sub
            End If '/////////ojo
        ElseIf mstrProcedencia = "Recepcion" Then
            If Not Atributos("Glbl", "Tllr_20_0010", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
                MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
                Unload Me
                Exit Sub
            End If '/////////ojo
        Else
            If Not Atributos("Glbl", "Tllr_20_0030", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
                MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
                Unload Me
                Exit Sub
            End If '/////////ojo
        End If
        
        tlbAgregarRepuestos.Visible = True

        FillConceptosInventario
        FillCampanas
        FillGarantia dtcGarantia, datGarantia, IIf(gstrProcedencia = "Presupuestos", True, False)
        FillRecepcionista dtcRecepcionista, datRecepcionista
        
        FillPromocion dtcPromocion, datPromocion
        
        FillTrabajos dtcTrabajo, datTrabajo
   
       
        FillTipoCono dtcTipoCono, datTipoCono
        FillTime gintHoraInicio, gintHoratermino, cboHora
        
        If gstrIdEmpresa = "20604506078" Then
            Me.txtCorreSpiga.Visible = True
            Me.Label9.Visible = True
        Else
            Me.txtCorreSpiga.Visible = False
            Me.Label9.Visible = False
        End If
        
        'FillTipoCargo dtcCargoCar, datCargoCar
        'FillMecanicos dtcMecanicoCar, datMecanico
        'FillPartePieza dtcPartePieza, datPartesPiezas
        
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
                mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT='" & gstrBusca & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
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
        Case 17 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    Bloqueo "V"
    ParametrosDefecto gstrIdEmpresa, gstrIdSucursal
    lblEstadoOTValor = ""
    txtTipo = ""
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
    'kjcv 12.11.13 para DesBloquear el Buscar Placa
    tlbPatente.Buttons(2).Enabled = True
    '//// que obligatoriamente elija un tipo de OT
    If InStr(gstrEmpresa, "AUTO SUMMIT") = 1 Then
        If mstrProcedencia <> "Presupuestos" Then
            frmElegirTipoOT.Show vbModal
            dtcGarantia.Enabled = False
        End If
    End If
    
    '////si es nuevo muestra la ot PRESUPUESTO
    If mstrProcedencia = "Presupuestos" Then
        dtcGarantia.BoundText = "PRE"
        dtcGarantia.Enabled = False
        lblEstadoOTValor = "PRESUPUESTO"
    Else
        dtcGarantia.BoundText = gstrIdTipoOtDefecto
    End If
    Me.Tag = "Crear"
    mstrIdPresupuestoOrigen = ""
'    gcurInsumoDef = gcurInsumo
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones                                                                       'AND Tllr_OT.ID_OT = Tllr_OT.ID_OT >'" & Trim(lblNroRecepcion) & "'
    If mstrProcedencia = "Presupuestos" Then
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
    Else
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
    End If
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
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
Dim lstrIdTipoCono As String

    If Not validacion() Then
        Exit Sub
    End If
    
    If Me.Tag = "Crear" Then
        If Me.dtcGarantia.BoundText <> "PRE" Then  '  And mstrLiquidaPresupuesto = True Then
            lblNroRecepcion = TraeCorrelativo(gcOrdenTrabajo, gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
        Else
            lblNroRecepcion = "P-" & TraeCorrelativoPresupuesto(gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
            mstrIdPresupuestoOrigen = lblNroRecepcion
            If Me.dtcTipoCono = "" Then
                lstrIdTipoCono = Retorna_Valor_General("Select Top 1 Id_Tipo_Cono from Tllr_Tipo_Cono", gcdynamic)
                dtcTipoCono.BoundText = lstrIdTipoCono
            Else
                lstrIdTipoCono = dtcTipoCono.BoundText
            End If
        End If
              
' Trim$(frmVentas.dbcboTipoVenta.BoundText)



        Dim valorMateriales As Double
        valorMateriales = ObtenerValorMateriales()
        gcurMateriales = valorMateriales
        
        stbTotalMateriales.Panels(2).Text = CStr(valorMateriales)
        
        
        gstrBusca = lblNroRecepcion
        mstrSQL = "INSERT INTO Tllr_OT "
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
        'kjcv 27.04.20
        mstrSQL = mstrSQL & " CorrelativoSpiga, "
        mstrSQL = mstrSQL & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor,"
'        mstrSql = mstrSql & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto,OrdenReparacion,NroReferencia,Bencina ) "
        'kjcv 19.09.13 se incluyo usuario y fecha de quien genera OT
'        mstrSQL = mstrSQL & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto,OrdenReparacion,NroReferencia,Bencina,Usr_Id,Usr_Fecha ) "
        'kjcv 24.10.13 se incluye campo de PDI
        mstrSQL = mstrSQL & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto,OrdenReparacion,NroReferencia,Bencina,Cuponera,Nro_Cupon,Id_Promo,Id_Trabajo,Usr_Id,Usr_Fecha,PDI, Correo, Telefono ) "
        'kjcv 03.02.22 Se agrega Forma de Pago
'        mstrSQL = mstrSQL & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto,OrdenReparacion,NroReferencia,Bencina,Cuponera,Nro_Cupon,Id_Promo,Id_Trabajo,Usr_Id,Usr_Fecha,Id_Tipo_Venta, PDI ) "
        mstrSQL = mstrSQL & " VALUES ("
        mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
        mstrSQL = mstrSQL & " '" & lblNroRecepcion & "', '" & gstrSeccion & "',"
        mstrSQL = mstrSQL & " '" & Trim(dtcGarantia.BoundText) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), "S/F") & "',"
        mstrSQL = mstrSQL & " '" & IIf(dtcGarantia.BoundText = "PRE", lstrIdTipoCono, dtcTipoCono.BoundText) & "', " & CLng(txtNroCono.Text) & ","
        mstrSQL = mstrSQL & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
        mstrSQL = mstrSQL & " " & CLng(txtKilAct) & ", '" & IIf(lblCompañia.Tag <> "", lblCompañia.Tag, "00") & "',"   'OJO
        mstrSQL = mstrSQL & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
        mstrSQL = mstrSQL & " '" & IIf(Me.dtcGarantia.BoundText = "PRE", "P", "V") & "','" & CDate(pckFechaAtencion.Value) & "', "
        mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
        mstrSQL = mstrSQL & " '" & "S/N" & "', '" & IIf(mstrIdPresupuestoOrigen <> "", mstrIdPresupuestoOrigen, "S/N") & "',"
        mstrSQL = mstrSQL & " '" & IIf(txtNroSiniestro <> " ", UCase(Trim(txtNroSiniestro)), "S/N") & " ','" & IIf(txtNroPoliza <> " ", UCase(Trim(txtNroPoliza)), "S/N") & "','" & IIf(txtLiquidador <> " ", UCase(Trim(txtLiquidador)), "S/L") & "' , "
        mstrSQL = mstrSQL & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), "S/S") & "' ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(txtDeducibleUF, ""))) & " , " & CCur(Val(SacarFormatoValor(txtDeduciblePesos, ""))) & " ,"
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMec.Panels(2).Text, ""))) & " ," & CCur(Val(SacarFormatoValor(stbTotalCarroceria.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalDesabolladura.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalPintura.Panels(2).Text, ""))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalTerceros.Panels(2).Text, ""))) & "," & CCur(Val(SacarFormatoValor(stbTotalRepuestos.Panels(2).Text, ""))) & ","
        
        
        'wcs 09/07/2024
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalMateriales.Panels(2).Text, ""))) & ", " & IIf(Me.dtcGarantia.BoundText = "PRE", 0, gcurInsumo) & ", "
       
       'se reasigna para que lo que se inserta por defecto en Insumos pase a Materiales
'       If gcurInsumo > 0 Then
'        gcurMateriales = gcurInsumo
'        gcurInsumo = 0
'       End If
'        mstrSQL = mstrSQL & " " & IIf(Me.dtcGarantia.BoundText = "PRE", 0, gcurMateriales) & ", " & CCur(Val(SacarFormatoValor(stbInsumos.Panels(2).Text, ""))) & ", "

        
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOtros.Panels(2).Text, ""))) & ", " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & " ,"
        'kjcv 27.04.20
        If gstrIdEmpresa = "20604506078" Then
        mstrSQL = mstrSQL & "'" & Me.txtCorreSpiga & "',"
        Else
        mstrSQL = mstrSQL & "'.',"
        End If
        
        mstrSQL = mstrSQL & " " & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & " ," & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) & ","
        mstrSQL = mstrSQL & " '" & lblIdCliente & "',"
        mstrSQL = mstrSQL & " '" & IIf(optMantencion.Value = True, "M", "R") & "',"
        mstrSQL = mstrSQL & " '" & IIf(cmdReserva.Enabled = False, "R", "N") & "',"
        mstrSQL = mstrSQL & " '" & mstrIdPresupuestoOrigen & "',"
'        mstrSql = mstrSql & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & ")"
        'kjcv 19.09.13 se agrego usuario y fecha de generacion de OT
'        mstrSQL = mstrSQL & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & ",'" & gstrIdUsuario & "','" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "')"
'         mstrSql = mstrSql & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & ",'" & gstrIdUsuario & "','" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "', '" & IIf(Len(txtPatente) > 16, "S", "N") & "' )"
        'kjcv 17.04.18
'        mstrSQL = mstrSQL & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & "," & cmbCuponera.ListIndex & ",'" & txtNroCupon.Text & "','" & gstrIdUsuario & "','" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "', '" & IIf(Len(txtPatente) > 16, "S", "N") & "' )"
        mstrSQL = mstrSQL & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & "," & cmbCuponera.ListIndex & ",'" & txtNroCupon.Text & "', " & IIf(Trim(dtcPromocion.BoundText) = "", 0, Trim(dtcPromocion.BoundText)) & " , '" & dtcTrabajo.BoundText & "' ,'" & gstrIdUsuario & "','" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "', '" & IIf(Len(txtPatente) > 16, "S", "N") & "',"
'        mstrSQL = mstrSQL & " '" & txtOrdenReparacion & "','" & txtNReferencia & "'," & cmbBencina.ListIndex & "," & cmbCuponera.ListIndex & ",'" & txtNroCupon.Text & "', " & IIf(Trim(dtcPromocion.BoundText) = "", 0, Trim(dtcPromocion.BoundText)) & " , '" & dtcTrabajo.BoundText & "' ,'" & gstrIdUsuario & "','" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "', '" & dbcboTipoVenta.BoundText & "' ,'" & IIf(Len(txtPatente) > 16, "S", "N") & "' )"
        mstrSQL = mstrSQL & " '" & Trim(txtCorreo.Text) & "',"
        mstrSQL = mstrSQL & " '" & Trim(txtTelefono.Text) & "')"
       'Trim$(dbcboTipoVenta.BoundText)
        mstrIdPresupuestoOrigen = ""
    Else
        mstrSQL = "UPDATE Tllr_OT "
        mstrSQL = mstrSQL & " SET Id_Garantia='" & Trim(dtcGarantia.BoundText) & "', "
        mstrSQL = mstrSQL & " Folio_Garantia='" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), ".") & "', "
        mstrSQL = mstrSQL & " Id_Tipo_Cono='" & dtcTipoCono.BoundText & "', "
        mstrSQL = mstrSQL & " Nro_Cono=" & CLng(txtNroCono.Text) & ", "
        mstrSQL = mstrSQL & " Patente='" & txtPatente.Text & "', "
        mstrSQL = mstrSQL & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
        mstrSQL = mstrSQL & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
        mstrSQL = mstrSQL & " Entrega_Estimada='" & CDate(pckFechaEntrega) & "', "
        mstrSQL = mstrSQL & " Hora_Entrega='" & cboHora.Text & "', "
        mstrSQL = mstrSQL & " Nro_Siniestro='" & IIf(txtNroSiniestro <> " ", UCase(Trim(txtNroSiniestro)), "S/N") & " ', "
        mstrSQL = mstrSQL & " Nro_Poliza='" & IIf(txtNroPoliza <> " ", UCase(Trim(txtNroPoliza)), "S/N") & "', "
        mstrSQL = mstrSQL & " Liquidador='" & IIf(txtLiquidador <> " ", UCase(Trim(txtLiquidador)), "S/L") & "', "
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
        mstrSQL = mstrSQL & " Total_Insumos=" & IIf(Me.dtcGarantia.BoundText = "PRE", 0, gcurInsumo) & ", "
        
        'kjcv 27.04.20
        mstrSQL = mstrSQL & " CorrelativoSpiga= '" & Me.txtCorreSpiga & "',"
'        mstrSQL = mstrSQL & " Total_Ot=" & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) + gcurInsumo & "  ,"
        'kjcv 14.10.15 se quito el valor de insumos
        mstrSQL = mstrSQL & " Total_Ot=" & CCur(Val(SacarFormatoValor(stbTotalOT.Panels(2).Text, ""))) & "  ,"
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
        mstrSQL = mstrSQL & " OrdenReparacion='" & txtOrdenReparacion & "',"
        mstrSQL = mstrSQL & " NroReferencia='" & txtNReferencia & "',"
        mstrSQL = mstrSQL & " Id_Promo=" & IIf(Trim(dtcPromocion.BoundText) = "", 0, Trim(dtcPromocion.BoundText)) & ","
        mstrSQL = mstrSQL & " Id_Trabajo= '" & Trim(dtcTrabajo.BoundText) & "',"
        'kjcv 19.09.19
        mstrSQL = mstrSQL & " Usr_Id='" & gstrIdUsuario & "',"
'        mstrSQL = mstrSQL & " Usr_Fecha='" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "'"
        'kjcv 24.10.13tlbBarraHerramientas
        mstrSQL = mstrSQL & " Usr_Fecha='" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "',"
'        mstrSQL = mstrSQL & " Id_Tipo_Venta='" & Trim(dbcboTipoVenta.BoundText) & "',"
        mstrSQL = mstrSQL & " PDI='" & IIf(Len(txtPatente) > 16, "S", "N") & "',"
        mstrSQL = mstrSQL & " Correo='" & Trim(txtCorreo.Text) & "',"
        mstrSQL = mstrSQL & " Telefono='" & Trim(txtTelefono.Text) & "'"
        
        'mstrSql = mstrSql & " Id_Presupuesto='" & mstrIdPresupuestoOrigen & "'"
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
            'kjcv 28.09.21
            If GuardaCampana(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
                MsgBox LoadResString(322)
            End If
            
        
            If GuardaMecanica(lblNroRecepcion, gcOrdenTrabajo) = False Then
                MsgBox LoadResString(321)
            End If
    
            If GuardaCarroceria(lblNroRecepcion, gstrSeccion, lblCompañia.Tag, gcOrdenTrabajo) = False Then
                MsgBox LoadResString(320)
            End If
        
            If GuardaOtros(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
                MsgBox LoadResString(328)
            End If
       
            If GuardaTerceros(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
                MsgBox LoadResString(319)
            End If
 
        
        
      
        'traspasa los repuestos de un presupuesto a una ot segun parametro
        If mstrLiquidaPresupuesto = True Then
            If gblnTraspasaRepuestos = True Then
                If GuardaRepuestosPresupuesto(lblNroRecepcion, gstrSeccion) = False Then
                    MsgBox LoadResString(318)
                End If
            End If
        Else
            If GuardaRepuestos(lblNroRecepcion, gstrSeccion, gcOrdenTrabajo) = False Then
                MsgBox LoadResString(318)
            End If
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
            'If MsgBox("Imprimir inventario de vehículo?", 4 + 32, "Reporte") = vbYes Then
            '    ImprimirConsulta
            'End If
            If MsgBox("Imprimirá la OT Nº " & lblNroRecepcion & ", Confirma el Documento", 4 + 32, "Imprime OT(Recepción)") = vbYes Then
                PrintOT
                
                'jn 17.01.2024
                'frmConfImprimirInventarioVehiculo.Show 1
                'Screen.MousePointer = vbHourglass
                'Screen.MousePointer = vbDefault
                'If (ConfirmarImprimirInventarioVehiculo = "S") Then
                '    ImprimirConsulta
                'End If
                If MsgBox("Imprimir inventario de vehículo?", 4 + 32, "Reporte") = vbYes Then
                    ImprimirConsulta
                    
                End If
                
                If mstrLiquidaPresupuesto = False Then 'cuando liquida presupuesto no borre la pantalla
                    AgregarRegistro
                End If
            Else
                If MsgBox("Imprimir inventario de vehículo?", 4 + 32, "Reporte") = vbYes Then
                    ImprimirConsulta
                End If
                
                If mstrLiquidaPresupuesto = False Then
                    AgregarRegistro
                End If
            End If
            
            
        End If
    End If '//////////////
End Sub
Sub GrabarPresupuesto(NumeroPresupuesto As String, NumeroOT As String, EstadoPresupuesto As String, MotivoAnula As String)

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
    mstrSQL = mstrSQL & " ReparacionMantencion, Estado_Reserva, Id_Presupuesto, Descripcion_Anula, Fecha_Liquidacion, Correo, Telefono ) "
    mstrSQL = mstrSQL & " VALUES ("
    mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
    mstrSQL = mstrSQL & " '" & NumeroOT & "', '" & gstrSeccion & "',"
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
    mstrSQL = mstrSQL & " '" & "M" & "',"
    mstrSQL = mstrSQL & " '" & "N" & "',"
    mstrSQL = mstrSQL & " '" & NumeroPresupuesto & "',"
    mstrSQL = mstrSQL & " '" & MotivoAnula & "',"
    mstrSQL = mstrSQL & " '" & Format(Date, "DD/MM/YYYY") & "',"
    mstrSQL = mstrSQL & " '" & Trim(txtCorreo.Text) & "',"
    mstrSQL = mstrSQL & " '" & Trim(txtTelefono.Text) & "')"
    
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        If GuardaInventario(NumeroPresupuesto, gstrSeccion, gcPresupuesto) = False Then
            MsgBox LoadResString(322)
        End If
        If GuardaMecanica(NumeroPresupuesto, gcPresupuesto) = False Then
            MsgBox LoadResString(321)
        End If
        If GuardaCarroceria(NumeroPresupuesto, gstrSeccion, lblCompañia.Tag, gcPresupuesto) = False Then
            MsgBox LoadResString(320)
        End If
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
        mstrSQL = "DELETE FROM Tllr_OT WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
            mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
            gstrSql = letSql(mstrWhere, mstrOrderBy)
            If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
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
frmBuscaOT.Show vbModal
Screen.MousePointer = 1
If gstrBusca <> "" Then
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
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
    If mstrProcedencia = "Presupuestos" Then
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
    Else
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
    End If
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
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
    If mstrProcedencia = "Presupuestos" Then
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
    Else
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
    End If
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
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
    If mstrProcedencia = "Presupuestos" Then
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
    Else
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
    End If
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT "
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
    If mstrProcedencia = "Presupuestos" Then
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
    Else
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
    End If
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
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
    
    If mstrProcedencia = "Presupuestos" Then
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado='P'"
    Else
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Estado<>'P'"
    End If
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT "
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
        .Item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(gstrProcedencia = "Recepcion", False, IIf(mblnAccesoEditar, True, False)))
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
    .Item("Liquidar").Enabled = True
    .Item("Anular").Enabled = True
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
    SetCheckOff .lvwCampana
    .lvwServiciosCarroceria.ListItems.Clear
    .lvwServiciosMecanica.ListItems.Clear
    .lvwServiciosTerceros.ListItems.Clear
    .lvwRepuestos.ListItems.Clear
    .lblNroRecepcion.Text = ""
    .dtcGarantia.BoundText = ""
    .dtcGarantia.Enabled = True
    .dtcPromocion.BoundText = ""
    .dtcPromocion.Enabled = True
    .dtcTrabajo.BoundText = ""
    .dtcTrabajo.Enabled = True
'    .dbcboTipoVenta.BoundText = ""
'    .dbcboTipoVenta.Enabled = True
    
    .pckFechaAtencion.Value = Now
    .lblHoraAtencion = Format$(Now, "HH:mm:ss")
    .ConfirmarImprimirInventarioVehiculo = ""
    .txtPatente.Text = ""
    .lblMarca.Caption = "": .lblIdMarca = ""
    .lblModelo.Caption = "": .lblIdModelo = ""
    .txtAño.Text = ""
    .lblColorE.Caption = ""
'    .lblChasis.Caption = ""
    .txtChasis.Text = ""
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
    .txtCorreSpiga = ""
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
    .lblFechaLiquidacion = ""
    .txtOrdenReparacion = ""
    .lblDocumentos = ""
    .txtNReferencia = ""
    .txtTipo = ""
    .txtCorreo = ""
    .txtTelefono = ""
End With
End Sub
Private Sub ValoresporDefecto()
    txtAño.Text = Year(Now)
    txtDeducibleUF.Text = "0"
    txtNroCono.Text = "0"
    txtDeduciblePesos.Text = "0"
    txtNroSiniestro.Text = " "
    txtNroPoliza.Text = " "
    txtLiquidador.Text = " "
    txtKilAct.Text = "0"
    lblEstadoOTValor = "VIGENTE"
    lblEstadoOTValor.Tag = "V"
End Sub
Private Function validacion() As Boolean
Dim intIndice As Integer
    validacion = True
With Me
    If .dtcGarantia.BoundText = "" Then
        MsgBox LoadResString(317), vbInformation, "Advertencia"
        dtcGarantia.Enabled = True
        dtcGarantia.SetFocus
        validacion = False
        Exit Function
    End If
    
    If .lvwCampana.SelectedItem.Selected = False Then
            MsgBox "Elegir una Campaña debe Especificarse..", vbInformation, "Advertencia"
            lvwCampana.Enabled = True
            lvwCampana.SetFocus
            validacion = False
            Exit Function
    End If
    
'    For intIndice = 1 To lvwCampana.ListItems.Count
'        Set .lvwCampana.SelectedItem = lvwCampana.ListItems(intIndice)
'        If .lvwCampana.SelectedItem.Checked = False Then
'            MsgBox "Elegir una Campaña debe Especificarse..", vbInformation, "Advertencia"
'            lvwCampana.Enabled = True
'            lvwCampana.SetFocus
'            validacion = False
'            Exit Function
'
'        End If
'    Next
'     'Valida Forma de Pago
'     If .dbcboTipoVenta.BoundText = "" Then
'        MsgBox "La Forma de Pago debe especificarse....", vbInformation, "Advertencia"
'        .dbcboTipoVenta.Enabled = True
'        .dbcboTipoVenta.SetFocus
'        validacion = False
'        Exit Function
'     End If
  
    
    If .dtcTrabajo.BoundText = "" Then
        MsgBox "El Tipo de Trabajo debe Especificarse..", vbInformation, "Advertencia"
        dtcTrabajo.Enabled = True
        dtcTrabajo.SetFocus
        validacion = False
        Exit Function
    End If
    
    Dim valKM As Double
    valKM = Val(txtKilAct.Text)
    If valKM <= 0 Then
        MsgBox "El campo Kms. Act. debe ser mayor a cero..", vbInformation, "Advertencia"
        txtKilAct.SetFocus
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
    
    If .txtCorreo.Text = "" Then
        MsgBox "Se debe ingresar el correo", vbInformation, "Advertencia"
        txtCorreo.SetFocus
        validacion = False
        Exit Function
    End If
    If .txtTelefono.Text = "" Then
        MsgBox "Se debe ingresar el telefono", vbInformation, "Advertencia"
        txtTelefono.SetFocus
        validacion = False
        Exit Function
    End If
    
    If Len(Trim(.txtTelefono.Text)) < 9 Then
        MsgBox "Se debe ingresar al menos 9 digitos", vbInformation, "Advertencia"
        txtTelefono.SetFocus
        validacion = False
        Exit Function
    End If
    
    If Not IsValidEmail(txtCorreo.Text) Then
        MsgBox "Por favor, ingrese una dirección de correo electrónico válida.", vbExclamation, "Correo electrónico inválido"
        txtCorreo.SetFocus
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
        If Me.dtcGarantia.BoundText <> "PRE" Then
            MsgBox LoadResString(312), vbInformation, "Advertencia"
            dtcTipoCono.SetFocus
            validacion = False
            Exit Function
        End If
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
    If gstrProcedencia = "Recepcion" Then
        If Me.cmbBencina.Text = "" Then
            MsgBox "El estado del Estanque de Gasolina debe contener un valor", vbExclamation, "Recepción"
            stbServicios.tab = 1
            Me.cmbBencina.SetFocus
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
        mstrSQL = "select ID_OT from TLLR_OT where SECCION_OT = '" & gstrSeccion & "' AND ID_OT ='" & lblNroRecepcion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción "
                validacion = False
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End Function

Public Function IsValidEmail(email As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$"
    re.IgnoreCase = True
    re.Global = False
    
    IsValidEmail = re.Test(email)
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

Private Sub tlbBusca_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim lstrNombre As String
Dim lstrSQL As String

Select Case Button.Key
    Case "Nuevo"
        gstrBusca = ""
        lstrNombre = ""
'        gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, gstrBusca, lstrNombre, apcrear, "Cliente - Proveedor", gstrIdSucursal)

        
        lblIdCliente = gstrBusca
        'ACTUALIZA PATENTE V/S CLIENTE
        lstrSQL = "Update Tllr_Vehiculo_Cliente set Id_Cliente_Proveedor='" & lblIdCliente & "' Where Patente='" & txtPatente & "'"
        Conexion.SendHost lstrSQL, , , , gcTiempoEspera
    Case "Buscar"
'        gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, lblIdCliente, lstrNombre, apeditar, "Cliente - Proveedor", gstrIdSucursal)
'        lblIdCliente = gstrBusca
'       Me.lblCliente.Caption = lblIdCliente
'       DatosCliente (lblIdCliente)
'kjcv 02-02-2012
        gstrRutCliente = ""
        gstrNombreCliente = ""
        Libreria.ClienteBuscar Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
         If gstrRutCliente <> "" Then
         '06.07.18 Valida Cliente Bloqueado
            If ValidaCliente(gstrRutCliente) Then
               Me.lblCliente.Caption = gstrNombreCliente
               Me.lblCliente.Tag = gstrRutCliente
            End If
        End If
    End Select
End Sub

Private Sub tlbCiaSeg_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case Is = "Nueva"
    gstrProcedencia = "Movimientos"
    frmMantenedorCompañiaSeguro.Show 1
    
Case Is = "Buscar"
    'gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Compañia_Seguro", "Id_Compañia_Seguro", "Nombre", "Busca Compañia de Seguro")
    gstrBusca = ""
    frmBuscarCiaSeguros.Show vbModal
    lblCompañia = NombreCiaSeg(gstrBusca)
    lblCompañia.Tag = gstrBusca
    FillConceptosVsCiaSeguro dtcConceptos, datConceptos, lblCompañia.Tag
    txtDeduciblePesos.SetFocus
End Select

End Sub

Private Sub tlbPatente_ButtonClick(ByVal Button As MSComctlLib.Button)
Screen.MousePointer = vbHourglass
Dim strPatente As String


If Me.Tag = "Crear" Then
    Select Case Button.Key
    Case "Nuevo"
        strPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apcrear)
        If gstrPresionoEnter = "OK" Then
            txtPatente = strPatente
            DatosVehiculo txtPatente
        End If
        
        
        
    Case "Buscar"
        gstrProcedencia = "Movimientos"
        frmBuscaVehiculo.Show vbModal
        'kjcv 30.10.15
    Case "Historial"
        frmHistorialPlaca.Show vbModal
        
    Case "Presupuesto"
              
                
        
    End Select
Else
    Select Case Button.Key
    Case "Nuevo"
        strPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apeditar)
         If gstrPresionoEnter = "OK" Then
            txtPatente = strPatente
             DatosVehiculo txtPatente
         End If
       
        
    Case "Buscar"
        gstrProcedencia = "Movimientos"
        frmBuscaVehiculo.Show vbModal
        'kjcv 30.10.15
    Case "Historial"
        frmHistorialPlaca.Show vbModal
        
    Case "Presupuesto"
        gstrProcedencia = "Presupuestos"
        HistorialPresupuesto
    End Select
End If
Screen.MousePointer = vbDefault
End Sub
Private Sub HistorialPresupuesto()
Screen.MousePointer = 1
frmHistorialPresupuesto.Show vbModal
Screen.MousePointer = 1
If gstrProcedencia = "Presupuestos" Then
If gstrBusca <> "" Then
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' and Estado='P'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
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

Private Sub tlbPatente_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If tlbPatente.Buttons(1).Key = "Nuevo" Then
    tlbPatente.Buttons(1).ToolTipText = IIf(Me.Tag = "Crear", "Nuevo Vehiculo", "Editar Vehiculo")
Else
    tlbPatente.Buttons(2).ToolTipText = IIf(Me.Tag = "Crear", "Buscar Vehiculo", "Buscar Vehiculo")
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub tlbTemparioCarroceria_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case Is = "Temparios"
        frmTemparios.Show
           
End Select
End Sub

Private Sub txtConcesionario_GotFocus()
MarcaTexto txtConcesionario
End Sub



Private Sub txtCorreo_KeyPress(KeyAscii As Integer)

 ' Verificar si el carácter ingresado es válido para un correo electrónico
    If Not IsValidEmailChar(KeyAscii) Then
        KeyAscii = 0 ' Cancelar la entrada del carácter no válido
        MsgBox "Solo se permiten letras, números, guiones bajos, puntos y arrobas en un correo electrónico.", vbExclamation, "Error de validación"
    End If

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

Private Sub txtLiquidador_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtNroCono_GotFocus()
MarcaTexto txtNroCono
End Sub




Private Sub txtNroCupon_LostFocus()
' formatea codigo correctamente
If Len(Me.txtNroCupon.Text) < 4 Then
    Me.txtNroCupon.Text = Lpad(Me.txtNroCupon.Text, "0", 4)
End If
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
''        kjcv 24 - 01 - 12
''        CheckPatente txtPatente, str1, str2  '/// devuelve el rut de la patente
'        txtFolioGarantia = str2
        If txtPatente <> "" Then
            'If Len(txtPatente) = 6 And lblPat.Caption = gstrNombrePatente Or Me.dtcGarantia.BoundText = "PEX" Then
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
                    'kjcv 15.11.13
                    If ConsultaPatente(txtPatente) = True Then
                        MsgBox "No hay Cupo en el Taller...", vbCritical, "Elisa"
                        Call DatosVehiculo(txtPatente)
                    Else
                        Call DatosVehiculo(txtPatente)
                    End If
'                    Call DatosVehiculo(txtPatente)
                Else
                    gstrProcedencia = "Movimientos"
                    gapAccion = apcrear
                    gstrKmsAutoNuevo = "Nuevo"
                    frmMantenedorVehiculoCliente.Show vbModal
                End If
                
'            ElseIf dtcGarantia.BoundText = "INW" Or dtcGarantia.BoundText = "INC" Then
'                If ConsultaVinExistencia(txtPatente) = True Then
'                    If ConsultaVehiculo(txtPatente) = True Then
'                        If MsgBox("La " & gstrNombrePatente & " " & txtPatente & " Ya Existe, Desea Desplegar los Datos", 4 + 32, "Patente Existente") = vbYes Then
'                            Call DatosVehiculo(txtPatente)
'                        Else
'                            LimpiaCampos
'                        End If
'                    Else
'                        gstrProcedencia = "Movimientos"
'                        gapAccion = apcrear
'                        frmMantenedorVehiculoCliente.Show vbModal
'                    End If
'                End If
'
'            ElseIf gstrValidaPatente = "N" Then
'                If ConsultaVehiculo(txtPatente) = True Then
'                    Call DatosVehiculo(txtPatente)
'                Else
'                    gstrProcedencia = "Movimientos"
'                    gapAccion = apcrear
'                    gstrKmsAutoNuevo = "Nuevo"
'                    frmMantenedorVehiculoCliente.Show vbModal
'                End If
'            Else
'                'MsgBox LoadResString(326)
'            End If
'
'        Else
'            MsgBox LoadResString(327)
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

'kjcv 24-01-12
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
        mstrSQL = mstrSQL & " Set Id_OT = '" & gstrBusca & "',"
        mstrSQL = mstrSQL & " Fecha_Emision_OT = getdate()" & ","
        mstrSQL = mstrSQL & " Estado='R'"
        mstrSQL = mstrSQL & " Where Id_Reserva='" & Mid(gstrBuscaReserva, 3, 5) & "'"
        mstrSQL = mstrSQL & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
            MsgBox "Error Al Actualizar Los Datos De La Reserva de Hora"
        End If
        
        '/// actualiza la reserva de repuestos
        mstrSQL = "Update Stck_Regularizacion "
        mstrSQL = mstrSQL & " Set Id_OT = '" & gstrSeccion & gstrBusca & "'"
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
    mstrSQL = "DELETE FROM Tllr_Otro_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
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
    mstrSQL = "DELETE FROM Tllr_OT WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT='" & Trim(pstrNroReserva) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
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
Dim AdoAnular As New ADODB.Recordset

    If MsgBox("¿ Realmente Desea eliminar esta Reserva de Hora?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        If TieneReservadeRepuestos Then
        
            mstrSQL = "Select Id_Regularizacion as Numero from Stck_Regularizacion where id_ot='" & gstrSeccion & lblNroRecepcion & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            If Conexion.SendHost(mstrSQL, AdoAnular, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                With AdoAnular
                    If Not .BOF And Not .EOF Then
                        .MoveFirst
                        While Not .EOF
                            NroRegularizacion = !NUMERO
                            Call Actualiza_Saldos_VS_Detalle("S", "Select Canrtidad, Id_Empresa, Id_sucursal, Id_Bodega,Id_Ubicacion,Id_Item From Stck_Regularizacion_Detalle Where Id_Regularizacion = '" & NroRegularizacion & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Empresa = '" & gstrIdEmpresa & "'")
                            
                            EliminaReservaRepuestos NroRegularizacion, lblNroRecepcion
                            
                            .MoveNext
                        Wend
                    End If
                End With
            End If
        
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
mstrSQL = mstrSQL & " Stck_Item.Precio_Venta as Precio, "
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
                    lsiItem.SubItems(2) = FormatoValor(!CANTY, "", 2)
                    lsiItem.SubItems(3) = FormatoValor(ValorNulo(!Precio), "", gintDecimalesMoneda)
                    lsiItem.SubItems(4) = ValorNulo(!Familia)
                    lsiItem.SubItems(5) = Me.lvwServiciosMecanica.SelectedItem.SubItems(6)
                    
                    If Me.dtcGarantia.BoundText = "PRE" And mstrAgregaPresupuesto = True And blnLlenaRepuestos = True Then
                        Set itmAux = lvwRepuestos.ListItems.Add(, , ValorNulo(!Codigo))
                        itmAux.SubItems(1) = ValorNulo(!Nombre)
                        itmAux.SubItems(2) = FormatoValor(!CANTY, "", 2)
                        itmAux.SubItems(3) = FormatoValor(!Precio, "", gintDecimalesMoneda)
                        itmAux.SubItems(4) = FormatoValor(0, "", 2)
                        itmAux.SubItems(5) = FormatoValor(0, "", gintDecimalesMoneda)
                        itmAux.SubItems(6) = "" 'TraeCargoDes(gstrIdCargo)
                        itmAux.SubItems(7) = gstrIdCargo
                        itmAux.SubItems(8) = Format(Val(SacarFormatoValor(itmAux.SubItems(2), "")) * Val(SacarFormatoValor(itmAux.SubItems(3), "")), "###,##0.00")
                        itmAux.SubItems(9) = ValorNulo(!IDFAM)
                        itmAux.SubItems(10) = "N"
                        itmAux.SubItems(11) = "PRESUPUESTO"
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

    lstrEstadoReserva = Retorna_Valor_General("Select estado_reserva As Codigo From Tllr_Ot where id_ot='" & Me.lblNroRecepcion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Seccion_OT='" & gstrSeccion & "'")
    If lstrEstadoReserva = "R" Then
        TieneReservadeRepuestos = True
    End If
End Function
Sub GrabaReservaRepuestosRecepcion()
    If Me.Tag = "Crear" Then
        lblNroRecepcion = TraeCorrelativo(gcOrdenTrabajo, gstrIdEmpresa, gstrIdSucursal, gstrSeccion)
        gstrBusca = lblNroRecepcion
        mstrSQL = "INSERT INTO Tllr_OT "
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
        mstrSQL = mstrSQL & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor, ReparacionMantencion, Estado_Reserva,Correo,Telefono ) "
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
        mstrSQL = mstrSQL & " '" & IIf(cmdReserva.Enabled = False, "R", "N") & "',"
        mstrSQL = mstrSQL & " '" & Trim(txtCorreo.Text) & "',"
        mstrSQL = mstrSQL & " '" & Trim(txtTelefono.Text) & "')"
        
    Else
        mstrSQL = "UPDATE Tllr_OT "
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
        mstrSQL = mstrSQL & " Estado_Reserva='" & IIf(Me.cmdReserva = False, "R", "N") & "',"
        mstrSQL = mstrSQL & " Correo='" & Trim(txtCorreo.Text) & "',"
        mstrSQL = mstrSQL & " Telefono='" & Trim(txtTelefono.Text) & "'"
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
    
    Screen.MousePointer = vbDefault
    'Pregunto si el presupuesto lo va agregar a una ot existente
    gstrProcedencia = "Movimientos"
    frmPresupuestoAdicional.Show vbModal
    gstrProcedencia = "Presupuestos"
    mstrLiquidaPresupuesto = True
    Screen.MousePointer = vbHourglass
    If gintOtExistente = 2 Then         'ot nueva
        gstrBuscaReserva = lblNroRecepcion
        mstrIdPresupuestoOrigen = lblNroRecepcion
        Me.Tag = "Crear"
        dtcGarantia.BoundText = gstrIdTipoOtDefecto
        GrabarRegistro                                              '/// graba el presupuesto en una ot definitiva
        GrabarPresupuesto gstrBuscaReserva, gstrBusca, "L", ""      '/// Graba presupuesto en tablas de presupuesto
        EliminaReserva gstrBuscaReserva                             '/// elimina el presupuesto que fue grabado anteriormente como OT
        
        MsgBox "Fue Creada la OT Numero : " & gstrBusca, vbInformation, "OT"
        
        UltimoRegistro
        
    ElseIf gintOtExistente = 1 Then         'ot existente
    
        Dim lstrNumeroPresupuesto As String
        Dim lstrSQL As String
        
        If GuardaMecanicaPresupuesto(gstrBusca, gstrSeccion) = False Then
            MsgBox LoadResString(321)
        End If
        If GuardaCarroceriaPresupuesto(gstrBusca, gstrSeccion) = False Then
            MsgBox LoadResString(320)
        End If
        If GuardaOtrosPresupuesto(gstrBusca, gstrSeccion) = False Then
            MsgBox LoadResString(328)
        End If
        If GuardaTercerosPresupuesto(gstrBusca, gstrSeccion) = False Then
            MsgBox LoadResString(319)
        End If
        If gblnTraspasaRepuestos = True Then
            If GuardaRepuestosPresupuesto(gstrBusca, gstrSeccion) = False Then
                MsgBox LoadResString(318)
            End If
        End If
        mstrIdPresupuestoOrigen = lblNroRecepcion
        GrabarPresupuesto lblNroRecepcion, gstrBusca, "L", ""       '/// Graba presupuesto en tablas de presupuesto
        EliminaReserva lblNroRecepcion                              '/// elimina el presupuesto que fue grabado anteriormente como OT
        'actualiza id_presupuesto de tllr_ot
        lstrNumeroPresupuesto = Retorna_Valor_General("Select Nro_Presupuesto_Origen from Tllr_OT Where id_ot='" & gstrBusca & "' And Seccion_OT='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
        
        lstrSQL = "Update Tllr_OT Set Nro_Presupuesto_Origen='" & lstrNumeroPresupuesto & "/" & lblNroRecepcion & "' "
        lstrSQL = lstrSQL & "Where Id_ot='" & gstrBusca & "' And Seccion_OT='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(lstrSQL, , , , gcTiempoEspera) = apAbort Then
            MsgBox "Problemas para actualizar el numero de presupuesto", vbInformation, "Actualización"
        End If
        
        UltimoRegistro
    End If
    mstrLiquidaPresupuesto = False
End Sub
Sub AnularPresupuesto()
Dim mstrMotivoAnula As String

    Screen.MousePointer = vbHourglass
    mstrMotivoAnula = InputBox("Ingrese El Motivo por que Anula :", "Por que Anula Presupuesto....")
    If mstrMotivoAnula <> "" Then
        gstrBuscaReserva = lblNroRecepcion
        GrabarPresupuesto gstrBuscaReserva, "S/N", "N", mstrMotivoAnula    '/// Graba presupuesto en tablas de presupuesto
        EliminaReserva gstrBuscaReserva          '/// elimina el presupuesto que fue grabado anteriormente como OT
        Renovar
    End If
End Sub
Function NumerosDocumentos(IdOT As String, SeccionOT As String) As String
Dim adoTemp As New ADODB.Recordset
Dim lstrSQL As String

    NumerosDocumentos = ""
    lstrSQL = "Select Nro_Factura_Emitida from Tllr_Facturacion where id_Ot='" & IdOT & "' And Seccion_OT='" & SeccionOT & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    If Conexion.SendHost(lstrSQL, adoTemp, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoTemp
            While Not .EOF
                NumerosDocumentos = NumerosDocumentos & ValorNulo(!Nro_Factura_Emitida) & "/"
                adoTemp.MoveNext
            Wend
        End With
    End If
End Function
Private Function GuardaMecanicaPresupuesto(strIdOt As String, strSeccion As String) As Boolean

'valida que no exista ya el servicio
If Me.lvwServiciosMecanica.ListItems.Count > 0 Then
    If ValidaServicioMecanica(strIdOt, strSeccion, Trim(lblIdMarca), Trim(lblIdModelo), IIf(Me.lvwServiciosMecanica.ListItems.Count > 0, Trim(Me.lvwServiciosMecanica.SelectedItem), "")) = True Then
        Exit Function
    End If
End If
GuardaMecanicaPresupuesto = True
With lvwServiciosMecanica
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intIndice)
        mstrSQL = "Insert Into Tllr_Mecanica_OT "
        mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
        mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
        mstrSQL = mstrSQL & " Id_Marca, Id_Modelo, "
        mstrSQL = mstrSQL & " Id_Servicio, "
        mstrSQL = mstrSQL & " Id_Tipo_Cargo,Mecanico_Designado,"
        mstrSQL = mstrSQL & " Horas,Valor,"
        mstrSQL = mstrSQL & " Porcentaje_Descuento, Monto_Descuento, "
        mstrSQL = mstrSQL & " SubTotal, Facturado)"
        mstrSQL = mstrSQL & " Values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
        mstrSQL = mstrSQL & " '" & strIdOt & "', '" & strSeccion & "',"
        mstrSQL = mstrSQL & " '" & Trim(lblIdMarca) & "','" & Trim(lblIdModelo) & "',"
        mstrSQL = mstrSQL & " '" & Trim(.SelectedItem) & "',"
        mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(6) & "'," & IIf(.SelectedItem.SubItems(8) = "", "NULL", " '" & .SelectedItem.SubItems(8) & "' ") & ", "
        mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & " , " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & " , "
        mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & " ," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
        mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & .SelectedItem.SubItems(11) & "' )"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
            GuardaMecanicaPresupuesto = False
            Exit Function
        End If
        Next
    Else
        GuardaMecanicaPresupuesto = True
    End If
End With
End Function

Private Function GuardaCarroceriaPresupuesto(strIdOt As String, strSeccion As String) As Boolean

GuardaCarroceriaPresupuesto = True
With lvwServiciosCarroceria
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            '/////////////////////////////////////////////////VALIDAR SI EXISTE EN PARENT
            'If ExisteRegistro(strCiaSeguro, .SelectedItem.SubItems(1), .SelectedItem.SubItems(4)) = True Then
                mstrSQL = "INSERT INTO Tllr_Carroceria_OT"
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
                mstrSQL = mstrSQL & " '" & strIdOt & "', '" & strSeccion & "',"                  '///nro ot, seccion
                mstrSQL = mstrSQL & " '" & frmRecepcion.lblCompañia.Tag & "', "                                         '///cia seguro
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
                    GuardaCarroceriaPresupuesto = False
                    Exit Function
                End If
            'End If
        Next
    Else
        GuardaCarroceriaPresupuesto = True
    End If
End With
End Function
Private Function GuardaOtrosPresupuesto(strIdOt As String, strSeccion As String) As Boolean

GuardaOtrosPresupuesto = True
With lvwOtrosServicios
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            mstrSQL = "INSERT INTO Tllr_Otro_OT"
            mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
            mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
            mstrSQL = mstrSQL & " Id_Otro_Servicio, "
            mstrSQL = mstrSQL & " Id_Tipo_Cargo,"
            mstrSQL = mstrSQL & " Mecanico_Asignado, "
            mstrSQL = mstrSQL & " Horas,Valor,"
            mstrSQL = mstrSQL & " Porcentaje_Descuento,Monto_Descuento,"
            mstrSQL = mstrSQL & " SubTotal,Descripcion_Otro,Facturado)"
            mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
            mstrSQL = mstrSQL & " '" & strIdOt & "', '" & strSeccion & "',"
            mstrSQL = mstrSQL & " '" & .SelectedItem & "', "
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(6)) & "', "
            mstrSQL = mstrSQL & " '" & IIf(Trim(.SelectedItem.SubItems(8)) = "", "SIN", Trim(.SelectedItem.SubItems(8))) & "', "
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
            mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & ",'" & UCase(Trim(.SelectedItem.SubItems(1))) & "','" & UCase(Trim(.SelectedItem.SubItems(11))) & "')"
            If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                GuardaOtrosPresupuesto = False
                Exit Function
            End If
        Next
    Else
        GuardaOtrosPresupuesto = True
    End If
End With
End Function
Private Function GuardaTercerosPresupuesto(strIdOt As String, strSeccion As String) As Boolean

GuardaTercerosPresupuesto = True
With lvwServiciosTerceros
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            mstrSQL = "INSERT INTO Tllr_Terceros_OT"
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
            mstrSQL = mstrSQL & " '" & strIdOt & "', '" & strSeccion & "',"
            mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(2) & "', "
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem) & "', "
            mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(14)) & "', "
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(6), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(7), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
            mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(9), "#####0.00"))) & ","
            mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(3) & "', "
            mstrSQL = mstrSQL & " '" & .SelectedItem.SubItems(4) & "', "
            mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(12), "#####0.00"))) & ",'" & .SelectedItem.SubItems(15) & "',"
            mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(10), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(11), "#####0.00"))) & ")"
            If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                GuardaTercerosPresupuesto = False
                Exit Function
            End If
        Next
    Else
        GuardaTercerosPresupuesto = True
    End If
End With
End Function

Private Function GuardaRepuestosPresupuesto(strIdOt As String, strSeccion As String) As Boolean

'primero actualiza tllr_repuestos_ot

GuardaRepuestosPresupuesto = True

With lvwRepuestos
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            If VerificaRepuesto(.SelectedItem, strIdOt, strSeccion, "Tllr_Repuestos_OT") = True Then
                mstrSQL = "UPDATE Tllr_Repuestos_OT"
                mstrSQL = mstrSQL & " SET Id_Tipo_Cargo='" & Trim(.SelectedItem.SubItems(7)) & "',"
                mstrSQL = mstrSQL & " Cantidad = " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & ", "
                mstrSQL = mstrSQL & " Valor = " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " Porcentaje_Descuento = " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " Monto_Descuento = " & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " SubTotal = " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " Facturado = " & UCase(Trim(IIf(.SelectedItem.SubItems(10) = "", "'N'", "'" & .SelectedItem.SubItems(10) & "'"))) & ","
           '     mstrSql = mstrSql & " cantidad_real = " & CDbl(Val(Format(.SelectedItem.SubItems(13), "#####0.0"))) & ", "
                mstrSQL = mstrSQL & " Consumo = '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "'"
                mstrSQL = mstrSQL & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
                mstrSQL = mstrSQL & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
                mstrSQL = mstrSQL & " Id_OT = '" & strIdOt & "' AND  "
                mstrSQL = mstrSQL & " Seccion_OT = '" & strSeccion & "' AND "
                mstrSQL = mstrSQL & " Id_Item = '" & .SelectedItem & "' "
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaRepuestosPresupuesto = False
                    Exit Function
                End If
            Else
                '///////////////////////////////////VALIDAR SI EXISTE EN PARENT
                mstrSQL = "INSERT INTO Tllr_Repuestos_OT"
                mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
                mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
                mstrSQL = mstrSQL & " Id_Item, "
                mstrSQL = mstrSQL & " Id_Tipo_Cargo, "
                mstrSQL = mstrSQL & " Cantidad, Valor,"
                mstrSQL = mstrSQL & " Porcentaje_Descuento,Monto_Descuento,"
                mstrSQL = mstrSQL & " SubTotal,Facturado,Consumo)"
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdOt & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem & "', "
                mstrSQL = mstrSQL & " '" & Trim(.SelectedItem.SubItems(7)) & "', "
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(4), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(5), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & CCur(Val(Format(.SelectedItem.SubItems(8), "#####0.00"))) & ",'" & .SelectedItem.SubItems(10) & "',"
                mstrSQL = mstrSQL & " '" & IIf(Mid(.SelectedItem.SubItems(11), 1, 1) = "P", "P", "C") & "')"
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaRepuestosPresupuesto = False
                    Exit Function
                End If
            End If
        Next
    Else
        GuardaRepuestosPresupuesto = True
    End If
End With


'Ahora actualiza repuestos reservados
GuardaRepuestosPresupuesto = True

With lvwRepuestos
    If .ListItems.Count > 0 Then
        For intIndice = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intIndice)
            If VerificaRepuesto(.SelectedItem, strIdOt, strSeccion, "Tllr_Repuestos_Reservados") = True Then
                mstrSQL = "UPDATE Tllr_Repuestos_Reservados"
                mstrSQL = mstrSQL & " SET Solicitado = " & CDbl(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & ", "
                mstrSQL = mstrSQL & " Precio_Unitario = " & CCur(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " Reservado= " & 0 & ","
                mstrSQL = mstrSQL & " Estado = 'V'" & ","
                mstrSQL = mstrSQL & " Tipo = 'Q'"
                mstrSQL = mstrSQL & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
                mstrSQL = mstrSQL & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
                mstrSQL = mstrSQL & " Id_OT = '" & strIdOt & "' AND  "
                mstrSQL = mstrSQL & " Seccion_OT = '" & strSeccion & "' AND "
                mstrSQL = mstrSQL & " Id_Item = '" & .SelectedItem & "' "
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaRepuestosPresupuesto = False
                    Exit Function
                End If
            Else
                '///////////////////////////////////VALIDAR SI EXISTE EN PARENT
                mstrSQL = "INSERT INTO Tllr_Repuestos_Reservados"
                mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal,"
                mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
                mstrSQL = mstrSQL & " Id_Item, "
                mstrSQL = mstrSQL & " Precio_Unitario,Solicitado,"
                mstrSQL = mstrSQL & " Reservado,Estado,Tipo)"
                mstrSQL = mstrSQL & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "
                mstrSQL = mstrSQL & " '" & strIdOt & "', '" & strSeccion & "',"
                mstrSQL = mstrSQL & " '" & .SelectedItem & "', "
                mstrSQL = mstrSQL & " " & CDbl(Val(Format(.SelectedItem.SubItems(3), "#####0.00"))) & "," & CCur(Val(Format(.SelectedItem.SubItems(2), "#####0.00"))) & ","
                mstrSQL = mstrSQL & " " & 0 & ",'V','Q')"
                If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apAbort Then
                    GuardaRepuestosPresupuesto = False
                    Exit Function
                End If
            End If
        Next
    Else
        GuardaRepuestosPresupuesto = True
    End If
End With


End Function

Sub ActualizarSaldoRepuestos(strIdDocumento, strSeccion)
Dim i As Integer
For intIndice = 1 To lvwRepuestos.ListItems.Count
    mstrSQL = "UPDATE Tllr_Repuestos_OT SET Saldo='" & lvwRepuestos.ListItems(intIndice).SubItems(12) & "'"
    mstrSQL = mstrSQL & " WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND  "
    mstrSQL = mstrSQL & " Id_Sucursal = '" & gstrIdSucursal & "' AND "
    mstrSQL = mstrSQL & " Id_OT = '" & strIdDocumento & "' AND  "
    mstrSQL = mstrSQL & " Seccion_OT = '" & strSeccion & "' AND "
    mstrSQL = mstrSQL & " Id_Item = '" & lvwRepuestos.ListItems(intIndice) & "' "
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
Next
End Sub
Sub VerificaCampañas()
Dim adoTemp As New ADODB.Recordset
Dim lstrSQL As String

    lstrSQL = "Select Vin,Id_Item,Servicio from Tllr_Campañas where Vin='" & Me.lblVin & "' And Estado='V' And Fecha_Inicio <='" & Format(Date, "DD/MM/YYYY") & "' And Fecha_Termino>='" & Format(Date, "DD/MM/YYYY") & "'"
    If Conexion.SendHost(lstrSQL, adoTemp, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoTemp
            While Not .EOF
                If MsgBox("Campaña:" & Chr(13) & adoTemp!servicio & Chr(13) & "esta VIGENTE la Realiza ahora ? ", vbInformation + vbYesNo, "Advertencia") = vbYes Then
                    txtComentario = Me.txtComentario & "Campaña :  " & adoTemp!servicio
                    mstrSQL = "Update Tllr_Campañas Set Estado='T' Where Vin='" & Me.lblVin & "' And Id_Item='" & adoTemp!Id_Item & "'"
                    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
                End If
                adoTemp.MoveNext
            Wend
        End With
    End If

End Sub

Private Sub txtSolicita_KeyPress(KeyAscii As Integer)
'kjcv  08-02-12
KeyAscii = UpCaseLetter(KeyAscii)

End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii >= 48 And KeyAscii <= 57) And _
       Not (KeyAscii = 45) And _
       Not (KeyAscii = 43) And _
       Not (KeyAscii = 32 And InStr(1, txtTelefono.Text, " ") = 0) And _
       Not (KeyAscii = 8) Then
        
        ' Si el carácter no es válido, mostrar mensaje de error y cancelar la entrada
        MsgBox "Solo se permiten números enteros, guiones, el signo más y un espacio en blanco.", vbExclamation, "Error de validación"
        KeyAscii = 0 ' Cancelar la entrada del carácter no válido
    End If

End Sub

Private Function IsValidEmailChar(KeyAscii As Integer) As Boolean
   
    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Or _
       (KeyAscii >= 48 And KeyAscii <= 57) Or _
       (KeyAscii = 95) Or _
       (KeyAscii = 46) Or _
       (KeyAscii = 8) Or _
       (KeyAscii = 64) Then
        IsValidEmailChar = True
    Else
        IsValidEmailChar = False
    End If
End Function
