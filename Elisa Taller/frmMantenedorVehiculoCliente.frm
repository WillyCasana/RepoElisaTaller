VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMantenedorVehiculoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehiculos de Clientes"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmMantenedorVehiculoCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro (Ctrl+G)"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
            ImageKey        =   "Cancelar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Registro (Ctrl+D)"
            ImageKey        =   "Borrar"
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
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
            ImageKey        =   "Primero"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
            ImageKey        =   "Siguiente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro (Ctrl+U)"
            ImageKey        =   "Ultimo"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+Q)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   7215
      Begin VB.CheckBox chkProblema 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "Comentario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   4250
         Width           =   4335
      End
      Begin VB.CommandButton cmdComentario 
         Appearance      =   0  'Flat
         Caption         =   "¿?"
         Height          =   375
         Left            =   6960
         TabIndex        =   32
         Top             =   4680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton optBitV 
         Appearance      =   0  'Flat
         Caption         =   "V.I.N."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   930
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   270
         Width           =   780
      End
      Begin VB.OptionButton optBitP 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox txtRutVeh 
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
         Left            =   5160
         TabIndex        =   1
         ToolTipText     =   "ID Vehículo"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1560
         TabIndex        =   7
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtPatente 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtNroChasis 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtNroMotor 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtNroVin 
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
         Left            =   1560
         TabIndex        =   6
         Top             =   3000
         Width           =   2295
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
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
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
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox chkVigencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Activo:"
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
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.Toolbar tlbBusca 
         Height          =   330
         Index           =   0
         Left            =   4560
         TabIndex        =   21
         Top             =   720
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBusca 
         Height          =   330
         Index           =   1
         Left            =   6000
         TabIndex        =   22
         Top             =   1080
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
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "Crear"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBusca 
         Height          =   330
         Index           =   2
         Left            =   6000
         TabIndex        =   23
         Top             =   1800
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBusca 
         Height          =   330
         Index           =   4
         Left            =   3960
         TabIndex        =   24
         Top             =   3480
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
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "Crear"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   6240
         Top             =   3720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
      End
      Begin MSComctlLib.Toolbar tlbBusca 
         Height          =   330
         Index           =   3
         Left            =   4560
         TabIndex        =   31
         Top             =   240
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblComentario 
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
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   28
         Top             =   2670
         Width           =   345
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
         Left            =   1560
         TabIndex        =   27
         Top             =   3840
         Width           =   4335
      End
      Begin VB.Label lblColorExt 
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
         Left            =   1560
         TabIndex        =   8
         Top             =   1800
         Width           =   4335
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
         Left            =   1560
         TabIndex        =   26
         Top             =   1080
         Width           =   4335
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
         Left            =   1560
         TabIndex        =   25
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Nº Motor"
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
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Vin"
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
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Kilometros"
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
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   1480
         Width           =   495
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
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
            Picture         =   "frmMantenedorVehiculoCliente.frx":038A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":049C
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":05AE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":06C0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":07D2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":08E4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":09F6
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":0B08
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":0C1A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":0D2C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":0E3E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":0F50
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":1062
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":1174
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":1286
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":1398
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":14AA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":18FC
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":1D4E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":1E60
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":1FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":2118
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":2274
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":23D0
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":2E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":32F0
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":3454
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":38B0
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":3A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":4D18
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":52B4
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":5410
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":556C
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":58C0
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":5C14
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":5F68
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":62BC
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":6610
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":6B54
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":6C66
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":6FBA
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":730E
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":7422
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":7776
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":7888
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculoCliente.frx":7BDC
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantenedorVehiculoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoPrincipal As New ADODB.Recordset
'Dim apfFormulario As New APFORM1.APFORM
Dim apfFormulario As New APFORM2.APFORM
Dim mstrSQL As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mblnSW As Boolean
Const mcNombreTabla = "Tllr_Vehiculo_Cliente"
Const mcCampoCodigo = "Patente"
Dim mstrBit As String
Dim mstrDigPat As String
Dim mstrRutPat As String
Function ExisteCliente(strIdCliente As String, ByRef pstrNombreCliente As String) As Boolean
mstrSQL = "SELECT Razon_Social FROM GLBL_Cliente_Proveedor Where Id_Cliente_Proveedor='" & strIdCliente & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            ExisteCliente = True
            pstrNombreCliente = IIf(Not IsNull(!Razon_Social), !Razon_Social, "")
        Else
            ExisteCliente = False
        End If
    End With
End If
Conexion.CloseHost AdoPrincipal
End Function

Private Sub chkProblema_Click()
If Me.chkProblema.Value = vbChecked Then
    Me.lblComentario.Visible = True
    Me.txtComentario.Visible = True
Else
    Me.lblComentario.Visible = False
    Me.txtComentario.Visible = False
End If
End Sub

Private Sub Form_Load()
mblnSW = True
Me.optBitP.Caption = gstrNombrePatente
'updAño.Value = Year(Now)
'kjcv 24-01-12
Label(1).Visible = False
txtRutVeh.Visible = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Me.Tag <> "Crear" Then
    If gstrProcedencia = "Movimientos" Then
        With frmRecepcion
            If gstrPresionoEnter = "OK" Then
                 '///NEO
                If Trim$(.txtPatente) <> "" Then
                    .DatosVehiculo .txtPatente
                Else
                    .DatosVehiculo Me.txtPatente
                End If
                '///
            
            End If
               
        End With
    ElseIf gstrProcedencia = "Presupuesto" Then
        With frmPresupuesto
            .DatosVehiculo .txtPatente
        End With
    End If
    Cancel = 0
Else
    If MsgBox("¿Está seguro de querer salir?", vbQuestion + vbYesNo + vbDefaultButton1, "Vehículo Cliente") = vbYes Then
        Cancel = 0
    Else
        Cancel = 1
    End If
    
End If

End Sub

Private Sub optBitP_Click()
If optBitP.Value = True Then
    mstrBit = "P"
    txtPatente.MaxLength = 6
    txtPatente.Text = ""
    txtRutVeh.Text = ""
    If txtPatente.Enabled = True Then txtPatente.SetFocus
    Me.tlbBusca(3).Visible = False
End If
End Sub

Private Sub optBitV_Click()
If optBitV.Value = True Then
    mstrBit = "V"
    txtPatente.MaxLength = 25
    txtPatente.Text = ""
    txtRutVeh.Text = ""
    If txtPatente.Enabled = True Then txtPatente.SetFocus
    Me.tlbBusca(3).Visible = True
End If
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
            ImprimirInforme
        Case "Primero"
            PrimerRegistro
        Case "Anterior"
            RegistroAnterior
        Case "Siguiente"
            RegistroSiguiente
        Case "Ultimo"
            UltimoRegistro
        Case "Renovar"
            Renovar
        Case "Cerrar"
            CerrarSalir
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
Dim adoTemp As New ADODB.Recordset

If gstrProcedencia = "Movimientos" Then
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0130", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        If gapAccion = apcrear Then
            AgregarRegistro
            txtPatente = frmRecepcion.txtPatente
            mstrSQL = "Select CodigoMarcaVehiculo from Auto_Parametros where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoTemp.BOF And Not adoTemp.EOF Then
                    Me.lblMarca.Tag = ValorNulo(adoTemp!CodigoMarcaVehiculo)
                End If
            End If
            Conexion.CloseHost adoTemp

            If Me.lblMarca.Tag <> "" Then
                Me.lblMarca.Caption = MarcaD(Me.lblMarca.Tag)
            End If
            txtPatente.SetFocus
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            
            Me.SetFocus
        End If
        mblnSW = False
        Screen.MousePointer = vbDefault
    End If
ElseIf gstrProcedencia = "Presupuesto" Then
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0130", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        If gapAccion = apcrear Then
            AgregarRegistro
            txtPatente = frmPresupuesto.txtPatente
            mstrSQL = "Select CodigoMarcaVehiculo from Auto_Parametros where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoTemp.BOF And Not adoTemp.EOF Then
                    Me.lblMarca.Tag = ValorNulo(adoTemp!CodigoMarcaVehiculo)
                End If
            End If
            Conexion.CloseHost adoTemp

            If Me.lblMarca.Tag <> "" Then
                Me.lblMarca.Caption = MarcaD(Me.lblMarca.Tag)
            End If
            txtPatente.SetFocus
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            
            Me.SetFocus
        End If
        mblnSW = False
        Screen.MousePointer = vbDefault
    End If
ElseIf gstrProcedencia = "ReservaHora" Then
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0130", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        If gapAccion = apcrear Then
            AgregarRegistro
            txtPatente = frmReservadeHoras.txtPatente
            mstrSQL = "Select CodigoMarcaVehiculo from Auto_Parametros where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoTemp.BOF And Not adoTemp.EOF Then
                    Me.lblMarca.Tag = ValorNulo(adoTemp!CodigoMarcaVehiculo)
                End If
            End If
            Conexion.CloseHost adoTemp

            If Me.lblMarca.Tag <> "" Then
                Me.lblMarca.Caption = MarcaD(Me.lblMarca.Tag)
            End If
            txtPatente.SetFocus
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            
            Me.SetFocus
        End If
        mblnSW = False
        Screen.MousePointer = vbDefault
    End If
ElseIf gstrProcedencia = "MantenedorPropio" Then
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0130", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        If gapAccion = apcrear Then
            AgregarRegistro
            txtPatente = frmMantenedorVehiculosPropios.txtPatente
            mstrSQL = "Select CodigoMarcaVehiculo from Auto_Parametros where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoTemp.BOF And Not adoTemp.EOF Then
                    Me.lblMarca.Tag = ValorNulo(adoTemp!CodigoMarcaVehiculo)
                End If
            End If
            Conexion.CloseHost adoTemp

            If Me.lblMarca.Tag <> "" Then
                Me.lblMarca.Caption = MarcaD(Me.lblMarca.Tag)
            End If
            txtPatente.SetFocus
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            
            Me.SetFocus
        End If
        mblnSW = False
        Screen.MousePointer = vbDefault
    End If

Else
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0130", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        If gapAccion = apcrear Then
           AgregarRegistro
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            
            Me.SetFocus
        End If
        If gapAccion = apninguno Then
           Renovar
        End If
        gapAccion = apninguno
        mblnSW = False
        Screen.MousePointer = vbDefault
    End If
End If
'kjcv 20.11.13
If Me.chkProblema.Value = vbChecked Then
    Me.lblComentario.Visible = True
    Me.txtComentario.Visible = True
Else
    Me.lblComentario.Visible = False
    Me.txtComentario.Visible = False
End If
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
            ImprimirInforme
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
Public Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    optBitP.Value = True
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtPatente & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtPatente & "' order by " & mcCampoCodigo
            If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
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
    txtPatente.Enabled = True
    txtPatente.SetFocus
End Sub
Private Sub GrabarRegistro()
If Not validacion() Then
    Exit Sub
End If

If Me.Tag = "Crear" Then
    mstrSQL = "INSERT INTO " & mcNombreTabla & ""
    mstrSQL = mstrSQL & "(" & mcCampoCodigo & ", Id_Marca,Id_Modelo,"
    mstrSQL = mstrSQL & "Id_Cliente_Proveedor, Id_Color_Exterior,"
    mstrSQL = mstrSQL & "Año, Kilometros_Actuales,"
    mstrSQL = mstrSQL & "Nro_Motor, Nro_Chasis, Vin, "
'    mstrSQL = mstrSQL & "Vigencia, Usr_Id, Usr_Fecha, BitID, RutVehiculo ) "
    'kjcv 15.11.13
    mstrSQL = mstrSQL & "Vigencia, Usr_Id, Usr_Fecha, BitID, RutVehiculo, Cliente_Problema, Comentario ) "
    mstrSQL = mstrSQL & " Values ('" & Trim(txtPatente.Text) & "', '" & Trim(lblMarca.Tag) & "', '" & Trim(lblModelo.Tag) & "', "
    mstrSQL = mstrSQL & " '" & Trim(lblCliente.Tag) & "' ,'" & IIf(Trim(lblColorExt.Tag) = "", "00", lblColorExt.Tag) & "' , "
    mstrSQL = mstrSQL & " " & txtAño & ", " & IIf(txtKilAct <> "", CLng(txtKilAct), 0) & ", "
    mstrSQL = mstrSQL & " '" & IIf(Trim(txtNroMotor.Text) = "", "S/N", UCase(Trim(txtNroMotor.Text))) & "','" & IIf(Trim(txtNroChasis.Text) = "", "S/N", UCase(Trim(txtNroChasis.Text))) & "', '" & IIf(Trim(txtNroVin.Text) = "", "S/VIN", UCase(Trim(txtNroVin.Text))) & "', "
'    mstrSQL = mstrSQL & " '" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', '" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "', '" & mstrBit & "', '" & txtRutVeh & "'  ) "
    'kjcv 15.11.13
    mstrSQL = mstrSQL & " '" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', '" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "', '" & mstrBit & "', '" & txtRutVeh & "', '" & IIf(chkProblema.Value = vbChecked, "S", "N") & "','" & Trim(txtComentario.Text) & "'  ) "
Else
    mstrSQL = "UPDATE " & mcNombreTabla & " "
    mstrSQL = mstrSQL & " SET Id_Marca ='" & Trim(lblMarca.Tag) & "', "
    mstrSQL = mstrSQL & " Id_Modelo='" & Trim(lblModelo.Tag) & "', "
    mstrSQL = mstrSQL & " Id_Cliente_Proveedor= '" & Trim(lblCliente.Tag) & "', "
    mstrSQL = mstrSQL & " Id_Color_Exterior='" & IIf(Trim(lblColorExt.Tag) = "", "00", lblColorExt.Tag) & "' , "
    mstrSQL = mstrSQL & " Año=" & txtAño & ", "
    mstrSQL = mstrSQL & " Kilometros_Actuales=" & CLng(txtKilAct) & ", "
    mstrSQL = mstrSQL & " Nro_Motor='" & IIf(Trim(txtNroMotor.Text) = "", "S/N", UCase(Trim(txtNroMotor.Text))) & "', "
    mstrSQL = mstrSQL & " Nro_Chasis='" & IIf(Trim(txtNroChasis.Text) = "", "S/N", UCase(Trim(txtNroChasis.Text))) & "', "
    mstrSQL = mstrSQL & " VIN='" & IIf(Trim(txtNroVin.Text) = "", "S/VIN", UCase(Trim(txtNroVin.Text))) & "', "
    mstrSQL = mstrSQL & " Vigencia='" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
    mstrSQL = mstrSQL & " Usr_Id= '" & gstrUsuario & "', "
    mstrSQL = mstrSQL & " Usr_Fecha= '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "', "
    mstrSQL = mstrSQL & " BitID= '" & mstrBit & "', "
'    mstrSQL = mstrSQL & " RutVehiculo= '" & txtRutVeh & "' "
    'kjcv 15.11.13
    mstrSQL = mstrSQL & " RutVehiculo= '" & txtRutVeh & "', "
    mstrSQL = mstrSQL & " Cliente_Problema='" & IIf(chkProblema.Value = vbChecked, "S", "N") & "', "
    mstrSQL = mstrSQL & " Comentario='" & Trim(txtComentario.Text) & "'"

    mstrSQL = mstrSQL & " WHERE " & mcCampoCodigo & "='" & Trim(txtPatente.Text) & "'"
End If

If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
    mblnTablaVacia = False
    ActivaBotones
    Me.Tag = ""
    If gstrProcedencia = "Movimientos" Then
        '///NEO
            gstrBusca = Me.txtPatente.Text
            gstrPresionoEnter = "OK"
        '///
        Unload Me
    End If
End If

End Sub
Private Sub BorrarRegistro()
Screen.MousePointer = vbDefault
If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
    mstrSQL = "DELETE FROM " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtPatente & "'"
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtPatente & "' order by " & mcCampoCodigo
        If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                LeerCampos
            Else
                mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtPatente & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
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
gstrProcedencia = "Mantenedor"
frmBuscaVehiculo.Show vbModal
If gstrBusca <> "" Then
    mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End If
Me.SetFocus
End Sub
Private Sub ImprimirInforme()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
Dim OTSeleccionada As String

    'Devuelve la ruta del directorio Windows
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
    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (PATENTE text,MARCA text, MODELO text, KMS text, ANO text, COLOR text, NCHASIS text, NMOTOR text, NVIN text, RUTCLIENTE text, CLIENTE text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
        
    Tabla.AddNew
    Tabla!Patente = Me.txtPatente
    Tabla!Marca = Me.lblMarca
    Tabla!Modelo = Me.lblModelo
    Tabla!KMS = Me.txtKilAct
    Tabla!ANO = Me.txtAño
    Tabla!Color = Me.lblColorExt
    Tabla!NCHASIS = Me.txtNroChasis
    Tabla!NMOTOR = Me.txtNroMotor
    Tabla!NVIN = Me.txtNroVin
    Tabla!RutCliente = Me.txtCodigo
    Tabla!Cliente = Me.lblCliente
    
    Tabla.Update
    
    Tabla.Close
   
    With rptMantenedor
        .ReportFileName = gstrPathReporte & "\APVEHICULOCLIENTE.RPT"
        .Formulas(0) = "Titulo='Listado Vehículo Cliente'"
        .Formulas(1) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(2) = "Ruc='" & gstrIdEmpresa & "'"
        .Formulas(3) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(4) = "Usuario='" & gstrUsuario & "'"
        .Formulas(5) = "Marcamodulo='ElisaTaller'"
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Connect = cnnAux.ConnectionString
        .Action = True
    End With

   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub
Private Sub PrimerRegistro()
mstrSQL = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroAnterior()
mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtPatente & "' order by " & mcCampoCodigo & " DESC"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroSiguiente()
mstrSQL = "SELECT TOP 1 * FROM " & mcNombreTabla & " WHERE PATENTE > '" & txtPatente & "' ORDER BY PATENTE ASC"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost AdoPrincipal
End Sub
Private Sub UltimoRegistro()
mstrSQL = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo & " DESC"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost AdoPrincipal
End Sub
Private Sub Renovar()
Set AdoPrincipal = New ADODB.Recordset
mstrSQL = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
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
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
With tlbBarraHerramientas.Buttons
    .Item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
    .Item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoEditar, True, False))
    .Item("Cancelar").Enabled = False
    .Item("Borrar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoBorrar, True, False))
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
txtPatente.Enabled = True
With tlbBarraHerramientas.Buttons
    .Item("Crear").Enabled = False
    .Item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
    .Item("Cancelar").Enabled = True
    .Item("Borrar").Enabled = False
    .Item("Buscar").Enabled = False
    .Item("Imprimir").Enabled = False
    .Item("Primero").Enabled = False
    .Item("Anterior").Enabled = False
    .Item("Siguiente").Enabled = False
    .Item("Ultimo").Enabled = False
    .Item("Renovar").Enabled = False
    .Item("Cerrar").Enabled = True
End With
End Sub
Private Sub VerificaTablaVacia()
If (Not AdoPrincipal.BOF And Not AdoPrincipal.EOF) And AdoPrincipal.RecordCount > 0 Then
    mblnTablaVacia = False
Else
    mblnTablaVacia = True
    LimpiaCampos
    MsgBox "La tabla no contiene registros...", vbInformation, "Advertencia"
End If
End Sub
Private Sub LeerCampos()
If mblnTablaVacia Then
    LimpiaCampos
    Exit Sub
End If

With AdoPrincipal
    If ValorNulo(!BitID) = "P" Then
        optBitP.Value = True
    Else
        optBitV.Value = True
    End If
    
    txtPatente.Text = !Patente
    
    If IsNull(!vigencia) Then
        chkVigencia.Value = vbUnchecked
    Else
        If !vigencia = "S" Then
            chkVigencia.Value = vbChecked
        Else
            chkVigencia.Value = vbUnchecked
        End If
    End If
    
    txtRutVeh = ValorNulo(!rutVehiculo)
    lblMarca.Tag = ValorNulo(!Id_Marca): lblMarca.Caption = MarcaD(!Id_Marca)
    lblModelo.Tag = ValorNulo(!Id_Modelo): lblModelo.Caption = ModeloD(!Id_Marca, !Id_Modelo)
    lblColorExt.Tag = ValorNulo(!Id_Color_Exterior): lblColorExt.Caption = ValorNulo(ColorExtDes(IIf(Not IsNull(!Id_Color_Exterior), !Id_Color_Exterior, ".")))
    txtAño = ValorNulo(!Año)
    lblCliente.Tag = ValorNulo(!Id_Cliente_Proveedor): lblCliente.Caption = ValorNulo(ClienteDes(IIf(Not IsNull(!Id_Cliente_Proveedor), !Id_Cliente_Proveedor, ".")))
    txtCodigo = ValorNulo(!Id_Cliente_Proveedor)
    txtKilAct.Text = !Kilometros_Actuales
    txtNroChasis.Text = !Nro_Chasis
    txtNroMotor = !Nro_Motor
    txtNroVin = !VIN
    'kjcv 15.11.13
    Me.txtComentario = ValorNulo(!Comentario)
    If IsNull(!Cliente_Problema) Then
        chkProblema.Value = vbUnchecked
    Else
        If !Cliente_Problema = "S" Then
            chkProblema.Value = vbChecked
        Else
            chkProblema.Value = vbUnchecked
        End If
    End If
    
    
End With
End Sub
Private Sub LimpiaCampos()
txtPatente.Text = ""
lblMarca.Caption = "": lblMarca.Tag = ""
lblModelo.Caption = "": lblModelo.Tag = ""
lblColorExt.Caption = "": lblColorExt.Tag = ""
txtRutVeh = ""
'kjcv 30.10.12
txtCodigo.Text = ""
'lblColorInt.Caption = "": lblColorInt.Tag = ""
lblCliente.Caption = "": lblCliente.Tag = ""
'updAño.Value = Year(Now)
'pckFechaVenta.Value = Date
'lblConcesionario.Caption = "": lblConcesionario.Tag = ""
txtKilAct = "0"
txtNroChasis = ""
txtNroMotor = ""
txtNroVin = ""
chkVigencia.Value = vbUnchecked

End Sub
Private Sub ValoresporDefecto()
With AdoPrincipal
    chkVigencia.Value = vbChecked
'    updAño.Value = Year(Now)
'    pckFechaVenta.Value = Date
    txtKilAct = "0"
End With
End Sub
Private Function validacion() As Boolean


validacion = True
If txtPatente.Text = "" Then
    MsgBox "La " & gstrNombrePatente & " debe contener un valor...", vbInformation, "Advertencia"
    txtPatente.SetFocus
    validacion = False
    Exit Function
End If
If lblMarca.Tag = "" Then
    MsgBox "La Marca debe contener un valor...", vbInformation, "Advertencia"
    Me.SetFocus
    validacion = False
    Exit Function
End If

If lblModelo.Tag = "" Then
    MsgBox "El Modelo debe contener un valor...", vbInformation, "Advertencia"
    Me.SetFocus
    validacion = False
    Exit Function
End If
'kjcv 30.10.12 Validacion de Ingreso de Cliente
If Me.txtCodigo.Text = "" Then
    MsgBox "El Cliente debe Especificarse...", vbInformation, "Advertencia"
    Me.txtCodigo.SetFocus
    validacion = False
    Exit Function
End If



If Me.lblCliente.Tag = "" Then
    MsgBox "El Cliente debe Especificarse...", vbInformation, "Advertencia"
    Me.SetFocus
    validacion = False
    Exit Function
End If

If Me.lblColorExt = "" Then
    MsgBox "El Color del Vehículo debe Especificarse...", vbInformation, "Advertencia"
    Me.SetFocus
    validacion = False
    Exit Function
End If

'//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As New ADODB.Recordset
        mstrSQL = "select  " & mcCampoCodigo & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtPatente & "'"
        If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Esta " & gstrNombrePatente & " ya esta registrada, Verifique " 'con la descripción " & Chr(13) & "[" & IIf(IsNull(adoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", adoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                validacion = False
                txtPatente.SetFocus
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set frmMantenedorVehiculoCliente = Nothing

End Sub
Private Sub RevizaAtributos()

mblnAccesoCrear = True
mblnAccesoEditar = True
mblnAccesoBorrar = True
mblnAccesoImprimir = True

End Sub

Private Sub tlbBusca_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim lstrNombre As String

Select Case Index
Case 0  ' MARCA
    Select Case Button.Key
    Case "Nuevo"
        gstrBusca = apfFormulario.Marca(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, "", "", apcrear)
        If gstrBusca <> "" Then
            lblMarca.Tag = gstrBusca
            lblMarca.Caption = MarcaD(gstrBusca)
        End If
    Case "Buscar"
        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Glbl_Marca", "Id_Marca", "Descripcion", "Buscar Marca")
        
        If gstrBusca <> "" Then
            lblMarca.Tag = gstrBusca
            lblMarca.Caption = MarcaD(gstrBusca)
        End If
    End Select
Case 1  'MODELO
    Select Case Button.Key
    Case "Nuevo"
        Dim IdModelo As String
        Dim descModelo As String
        Libreria.ModelosNuevo Conexion, IdModelo, descModelo, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
        lblModelo.Tag = IdModelo
        lblModelo.Caption = descModelo
    Case "Buscar"
        gstrBusca = apfFormulario.BuscarRegistrosModelo(Conexion, "Glbl_Modelo", "Id_Modelo", "Id_marca", "descripcion", "Buscar Modelo", lblMarca.Tag)
        If gstrBusca <> "" Then
            lblModelo.Tag = gstrBusca
            lblModelo.Caption = ModeloD(lblMarca.Tag, gstrBusca)
        End If
    End Select
Case 2  'COLOR EXTERIOR
    Select Case Button.Key
    Case "Nuevo"
    
    Case "Buscar"
        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Glbl_Color_Exterior", "Id_Color_Exterior", "Descripcion", "Buscar Color Exterior")
        If gstrBusca <> "" Then
            lblColorExt.Tag = gstrBusca
            lblColorExt.Caption = ColorExtDes(gstrBusca)
        End If
    End Select
Case 3 ' Busca VIN en Auto Stock
    frmBuscarVehiculo.Show vbModal
Case 4  'CLIENTE
    Select Case Button.Key
    Case "Nuevo"
        gstrBusca = ""
        lstrNombre = ""
        Libreria.ClienteNuevo Conexion, gstrBusca, gstrRazonSocial, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
        If gstrBusca <> "" Then
                lblCliente = lstrNombre
                lblCliente.Tag = gstrBusca
        End If
    Case "Buscar"
        Libreria.ClienteBuscar Conexion, gstrBusca, gstrRazonSocial, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
        If gstrBusca <> "" Then
        'kjcv 06.07.18
            If ValidaCliente(gstrBusca) Then
                lblCliente.Tag = gstrBusca
                lblCliente.Caption = ClienteDes(gstrBusca)
                txtCodigo = gstrBusca
            End If
        End If
    End Select
End Select

End Sub



Private Sub txtAño_GotFocus()
MarcaTexto txtAño
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
Dim mstrCliente As String

If KeyAscii = 13 Then
    If ExisteCliente(txtCodigo, mstrCliente) = True Then
        lblCliente = mstrCliente
        lblCliente.Tag = txtCodigo
    Else
'        gstrBusca = ""
'        mstrCliente = ""
'        gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtCodigo, mstrCliente, apcrear, "Cliente - Proveedor", gstrIdSucursal)
'        lblCliente = mstrCliente
'        lblCliente.Tag = gstrBusca
'        txtCodigo = gstrBusca
        'kjcv 13.11.13
        gstrRutCliente = ""
        gstrNombreCliente = ""
        Libreria.ClienteBuscar Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
         If gstrRutCliente <> "" Then
            Me.lblCliente = gstrNombreCliente
            Me.lblCliente.Tag = gstrRutCliente
        End If
        
    End If
End If
End Sub

Private Sub txtKilAct_GotFocus()
MarcaTexto txtKilAct
End Sub

Private Sub txtNroChasis_GotFocus()
MarcaTexto txtNroChasis
End Sub

Private Sub txtNroMotor_GotFocus()
MarcaTexto txtNroMotor
End Sub

Private Sub txtNroVin_GotFocus()
MarcaTexto txtNroVin
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)

'If gstrValidaPatente = "S" Then
'    If mstrBit <> "V" Then
'        KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'    End If
'End If
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If


End Sub

Private Sub txtPatente_LostFocus()
'If txtPatente <> "" Then
''    CheckPatente txtPatente, mstrDigPat, mstrRutPat
'    txtRutVeh = mstrRutPat
'    mstrRutPat = ""
'End If
End Sub
