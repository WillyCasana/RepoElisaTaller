VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmMargenRep 
   Caption         =   "Margen de Repuestos"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMargenRep.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptOT 
      Left            =   7440
      Top             =   600
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6720
      Begin VB.OptionButton optcarroceria 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Carrocería"
         ForeColor       =   &H80000008&
         Height          =   264
         Left            =   3810
         TabIndex        =   7
         Top             =   525
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.OptionButton optmecanica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Mecánica"
         ForeColor       =   &H80000008&
         Height          =   264
         Left            =   5325
         TabIndex        =   6
         Top             =   525
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1188
      End
      Begin VB.TextBox txtNroOt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         MaxLength       =   15
         TabIndex        =   3
         Top             =   525
         Width           =   2670
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Index           =   0
         Left            =   10485
         Top             =   2730
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
               Picture         =   "frmMargenRep.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMargenRep.frx":3E66
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCliente 
         Height          =   330
         Left            =   2850
         TabIndex        =   2
         Top             =   525
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta(0)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Número de OT"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   225
         Width           =   1740
      End
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   4680
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   8255
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pieza"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Familia"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "% Descto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Descto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Subtotal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Precio Costo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Margen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Margen %"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraNueva"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Crear"
            Object.ToolTipText     =   "Nueva búsqueda"
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar "
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir "
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "CotizaPerdida"
            Object.ToolTipText     =   "Cotización Perdida"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Vender"
            Object.ToolTipText     =   "Pasar a Venta"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Agenda"
            Object.ToolTipText     =   "Agenda Diaria"
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
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl + Q)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":3F78
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":408A
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":419C
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":42AE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":43C0
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":44D2
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":45E4
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":46F6
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":4808
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":491A
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":4A2C
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":4B3E
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":4C50
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":4D62
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":4E74
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":4F86
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":5098
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":54EA
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":593C
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":5A4E
            Key             =   "Foto"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":5B62
            Key             =   "Venta"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":5C5E
            Key             =   "Agenda"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":5DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":5F96
            Key             =   "PASAVENTA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbTotales 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Suma - Totales"
            TextSave        =   "Suma - Totales"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Suma - Costos"
            TextSave        =   "Suma - Costos"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Suma - Margen"
            TextSave        =   "Suma - Margen"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Margen Total"
            TextSave        =   "Margen Total"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
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
   Begin MSComctlLib.ImageList ImgBarraNueva 
      Left            =   10440
      Top             =   600
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
            Picture         =   "frmMargenRep.frx":62B6
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":63C8
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":64DA
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":65EC
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":66FE
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6810
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6922
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6A34
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6B46
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6C58
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6D6A
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6E7C
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":6F8E
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":70A0
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":71B2
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":72C4
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":73D6
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":7828
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":7C7A
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":7D8C
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":7EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":8044
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":81A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":82FC
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":8DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":921C
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":9380
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":97DC
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":9938
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":AC44
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":B1E0
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":B33C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":B498
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":B7EC
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":BB40
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":BE94
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":C1E8
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":C53C
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":CA80
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":CB92
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":CEE6
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":D23A
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":D34E
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":D6A2
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":D7B4
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMargenRep.frx":DB08
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMargenRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
Sub ImprimirConsulta()
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
    
    If lvDetalle.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Pieza text,Familia text,cantidad text,valor text,pdescto text,descto text,subtotal text,preciocosto text,Margen text,Pmargen text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!pieza = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!Familia = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!cantidad = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Valor = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!pdescto = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!descto = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!SubTotal = (IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6)))
        Tabla!preciocosto = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla!margen = IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
        Tabla!pmargen = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
        Tabla.Update
    Next i
   Tabla.Close
   Dbsnueva.Close
   
   With rptOT
        .ReportFileName = gstrPathReporte & "\MARGENREPUESTOS.rpt"
        .WindowTitle = "Margen de Repuestos"
        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='MARGEN DE REPUESTOS'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "SUMSUBTOTAL='" & Me.stbTotales.Panels(2).Text & "'"
        .Formulas(6) = "SUMCOSTO='" & Me.stbTotales.Panels(4).Text & "'"
        .Formulas(7) = "SUMMARGEN='" & Me.stbTotales.Panels(6).Text & "'"
        .Formulas(8) = "OT='" & Me.txtNroOt & "'"
        .Formulas(9) = "SECCION='" & IIf(Me.optCarroceria.Value = True, "CARROCERIA", "MECANICA") & "'"
        .Formulas(10) = "MargenTotal='" & Me.stbTotales.Panels(8).Text & "'"
        .Destination = crptToWindow
        .Action = True
   End With
   
''   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub
Sub cmdBuscarOT_Click()
Dim mstrSql As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem

lvDetalle.ListItems.Clear
mstrWhere = ""
    'procedimiento almacenado
    mstrSql = "Exec Tllr_MargenRepuestos_OT " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & IIf(Me.optCarroceria.Value = True, "C", "M") & "','" & Me.txtNroOt & "'"
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
            'If .Fields(2) <> 0 Then
              Set itmItem = lvDetalle.ListItems.Add(, , .Fields(0))
              itmItem.SubItems(1) = ValorNulo(.Fields(1))
              itmItem.SubItems(2) = ValorNulo(.Fields(2)) 'Cantidad
              itmItem.SubItems(3) = FormatoValor(ValorNulo(.Fields(3)), "", gintDecimalesMoneda) 'Valor
              itmItem.SubItems(4) = ValorNulo(.Fields(4)) '% descto
              itmItem.SubItems(5) = ValorNulo(.Fields(5)) 'Monto descto
              itmItem.SubItems(6) = FormatoValor(ValorNulo(.Fields(6)), "", gintDecimalesMoneda) 'Subtotal
              itmItem.SubItems(7) = FormatoValor(Costo_Promedio_Repuesto(txtNroOt, .Fields(0)), "", gintDecimalesMoneda) 'Precio costo
              itmItem.SubItems(8) = FormatoValor(ValorNulo(.Fields(6)) - CDbl(itmItem.SubItems(7)), "", gintDecimalesMoneda) 'Margen (subtotal - preciocosto * cantidad)
              If .Fields(6) = 0 Then
                itmItem.SubItems(9) = "0" 'Margen Porcentual
              Else
                itmItem.SubItems(9) = Round((CDbl(itmItem.SubItems(8)) * 100) / ValorNulo(.Fields(6)), 2) 'Margen Porcentual
              End If
           ' End If
            adoTemp.MoveNext
          Wend
       End If
    End With
    End If
    
    With Me.stbTotales
        .Panels(2).Text = FormatoValor(TotalSeccion(lvDetalle, 6), "", gintDecimalesMoneda)
        .Panels(4).Text = FormatoValor(TotalSeccion(lvDetalle, 7), "", gintDecimalesMoneda)
        .Panels(6).Text = FormatoValor(TotalSeccion(lvDetalle, 8), "", gintDecimalesMoneda)
        If .Panels(2) > 0 Then
            .Panels(8).Text = FormatoValor(CDbl(.Panels(6).Text) * 100 / CDbl(.Panels(2)), "%", 2)
        Else
            .Panels(8) = 0
        End If
    End With
    Screen.MousePointer = 1
    
    
End Sub
Private Sub cmdImprimir_Click()
If lvDetalle.ListItems.Count > 0 Then
    ImprimirConsulta
Else
    MsgBox "no"
End If
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    If Not Atributos("Glbl", "Tllr_30_0030", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
SW = True
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
frmBuscaOT.Show vbModal
Me.txtNroOt = gstrBusca

If gstrSeccion = "M" Then
    Me.optMecanica.Value = True
    Me.optCarroceria.Value = False
End If

If gstrSeccion = "C" Then
    Me.optCarroceria.Value = True
    Me.optMecanica.Value = False
End If
End Sub


Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
     
        Case "Buscar"
            cmdBuscarOT_Click
        Case "Imprimir"
            ImprimirConsulta
       
        Case "Cerrar"
            CerrarSalir
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
                SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            'CancelarAgregaRegistro
        Case 14 And tlbBarraHerramientas.Buttons.item("Crear").Enabled
            KeyAscii = 0
            'AgregarRegistro
        Case 7 And tlbBarraHerramientas.Buttons.item("Grabar").Enabled
            KeyAscii = 0
            'GrabarRegistro
        Case 4 And tlbBarraHerramientas.Buttons.item("Borrar").Enabled = False
            KeyAscii = 0
            'BorrarRegistro
        Case 2 And tlbBarraHerramientas.Buttons.item("Buscar").Enabled
            KeyAscii = 0
'            BuscarRegistro
        Case 9 And tlbBarraHerramientas.Buttons.item("Imprimir").Enabled
            KeyAscii = 0
'            ImprimirInforme
        Case 16 And tlbBarraHerramientas.Buttons.item("Primero").Enabled
            KeyAscii = 0
            'PrimerRegistro
        Case 1 And tlbBarraHerramientas.Buttons.item("Anterior").Enabled
            KeyAscii = 0
            'RegistroAnterior
        Case 19 And tlbBarraHerramientas.Buttons.item("Siguiente").Enabled
            KeyAscii = 0
            'RegistroSiguiente
        Case 21 And tlbBarraHerramientas.Buttons.item("Ultimo").Enabled
            KeyAscii = 0
            'UltimoRegistro
        Case 18 And tlbBarraHerramientas.Buttons.item("Renovar").Enabled
            KeyAscii = 0
            'Renovar
        Case 17 And tlbBarraHerramientas.Buttons.item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub

Sub CerrarSalir()
Unload Me

End Sub
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

