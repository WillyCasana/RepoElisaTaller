VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmEmisionOrdCom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Compra"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14235
   Icon            =   "frmEmisionOrdCom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   60
      TabIndex        =   7
      Top             =   390
      Width           =   11745
      Begin VB.Frame Frame5 
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
         Height          =   510
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   2565
         Begin VB.OptionButton optTerceros 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Terceros"
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
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optCarroce 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Carroceria"
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
            Height          =   195
            Left            =   1215
            TabIndex        =   35
            Top             =   225
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNroOT 
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
         Height          =   288
         Left            =   7455
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1020
         Width           =   3795
      End
      Begin VB.TextBox txtCondCompra 
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
         Height          =   288
         Left            =   4680
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1050
         Width           =   2475
      End
      Begin VB.TextBox txtContacto 
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
         Height          =   288
         Left            =   7455
         MaxLength       =   50
         TabIndex        =   0
         Top             =   465
         Width           =   4110
      End
      Begin VB.TextBox txtProveedor 
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
         Height          =   288
         Left            =   3240
         TabIndex        =   1
         Top             =   480
         Width           =   4110
      End
      Begin VB.TextBox txtNroOrden 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   420
         Left            =   120
         TabIndex        =   9
         Text            =   "0"
         Top             =   435
         Width           =   1560
      End
      Begin MSComCtl2.DTPicker pckFecha 
         Height          =   315
         Left            =   1845
         TabIndex        =   25
         Top             =   465
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93650945
         CurrentDate     =   36781
      End
      Begin MSComctlLib.Toolbar tlbProveedor 
         Height          =   330
         Left            =   6660
         TabIndex        =   27
         Top             =   150
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
               Object.ToolTipText     =   "Nuevo Proveedor"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Proveedor"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tblOT 
         Height          =   330
         Left            =   11265
         TabIndex        =   33
         Top             =   1005
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         DisabledImageList=   "ImgBarraHerramienta"
         HotImageList    =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Patente"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin VB.Label lblEstado 
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   37
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Nro. O/T"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7455
         TabIndex        =   32
         Top             =   825
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Condición de Pago"
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
         Left            =   4680
         TabIndex        =   28
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1860
         TabIndex        =   26
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblencargado 
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7440
         TabIndex        =   11
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lblproveedor 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3240
         TabIndex        =   10
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label lbl 
         Caption         =   "N° Orden"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   810
      End
   End
   Begin Crystal.CrystalReport rptOrdenCompra 
      Left            =   6720
      Top             =   -15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   3
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame3 
      Height          =   3540
      Left            =   60
      TabIndex        =   24
      Top             =   1785
      Width           =   11745
      Begin MSComctlLib.ListView lvwDetalle 
         Height          =   2865
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   5054
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "nro"
            Text            =   "Item"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   11465
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio Unitario"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Sub - Total"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbItemOrdCom 
         Height          =   330
         Left            =   75
         TabIndex        =   29
         Top             =   3120
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   582
         ButtonWidth     =   1746
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imlOtro"
         DisabledImageList=   "imlOtro"
         HotImageList    =   "imlOtro"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "Agregar"
               Object.ToolTipText     =   "Agrega Servicio Nuevo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "Quitar"
               Object.ToolTipText     =   "Quitar Servicio"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlOtro 
         Left            =   2370
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmisionOrdCom.frx":179A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmisionOrdCom.frx":1BEE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2280
      Left            =   7455
      TabIndex        =   13
      Top             =   5355
      Width           =   4350
      Begin VB.TextBox txtPorDes 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3540
         TabIndex        =   30
         Top             =   480
         Width           =   750
      End
      Begin VB.TextBox txttotal 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1905
         Width           =   1776
      End
      Begin VB.TextBox txtiva 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1470
         Width           =   1776
      End
      Begin VB.TextBox txtneto 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1155
         Width           =   1776
      End
      Begin VB.TextBox txtdescuento 
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
         Left            =   2520
         TabIndex        =   6
         Text            =   "0"
         Top             =   840
         Width           =   1776
      End
      Begin VB.TextBox txtsubtotal 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   135
         Width           =   1776
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "% Descuento"
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
         Left            =   75
         TabIndex        =   31
         Top             =   525
         Width           =   1290
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   4284
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label lbltotal 
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   19
         Top             =   1935
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "IGV"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   18
         Top             =   1500
         Width           =   1020
      End
      Begin VB.Label lblneto 
         Caption         =   "Neto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   17
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label lbldescuento 
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   16
         Top             =   870
         Width           =   1230
      End
      Begin VB.Label lblsubtotal 
         Caption         =   "Sub-Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   75
         TabIndex        =   15
         Top             =   180
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   45
      TabIndex        =   12
      Top             =   5355
      Width           =   7320
      Begin VB.TextBox txtObs 
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
         Height          =   2010
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   195
         Width           =   7152
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   7740
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2042
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2154
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2266
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2378
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":248A
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":259C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":26AE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":27C0
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":28D2
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":29E4
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2AF6
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2C08
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2D1A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2E2C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":2F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":3050
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":3162
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":35B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":3A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmisionOrdCom.frx":3B18
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+Q)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEmisionOrdCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset

Dim mstrSql As String
Dim mblnTablaVacia As Boolean

Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim x As Integer

Dim mblnSW As Boolean
Dim mstrEstado As String
Const mcNombreTabla = "Tllr_Orden_Compra"
Const mcCampoCodigo = "Id_Orden"

Sub DetalleOrden(pstrEmpresa As String, pstrSucursal As String, pstrNroOrden As Long)
lvwDetalle.ListItems.Clear
gstrSql = "SELECT * From TLLR_DETALLE_ORDEN_COMPRA"
gstrSql = gstrSql & " WHERE ID_EMPRESA = '" & pstrEmpresa & "' AND ID_SUCURSAL = '" & pstrSucursal & "' AND ID_ORDEN= " & pstrNroOrden & " Order by Item"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set glsiItem = lvwDetalle.ListItems.Add(, , !item)
                glsiItem.SubItems(1) = !Descripcion
                glsiItem.SubItems(2) = !cantidad
                glsiItem.SubItems(3) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
                glsiItem.SubItems(4) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
                .MoveNext
            Wend
        End If
    End With
End If
Conexion.CloseHost gadoPrincipal
End Sub

Sub Totales()
txtSubTotal = FormatoValor(TotalSeccion(lvwDetalle, 4), "", gintDecimalesMoneda)
'If Val(SacarFormatoValor(txtdescuento, "")) > 0 Then
    
'Else
    txtneto = FormatoValor(Val(SacarFormatoValor(txtSubTotal, "")) - Val(SacarFormatoValor(txtDescuento, "")), "", gintDecimalesMoneda)
    txtiva = FormatoValor(Val(SacarFormatoValor(txtneto, "")) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto), "", gintDecimalesMoneda)
    txtTotal = FormatoValor(Val(SacarFormatoValor(txtneto, "")) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), "", gintDecimalesMoneda)
'End If
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


Sub GuardaDetalle()

mstrSql = " DELETE FROM Tllr_Detalle_Orden_Compra WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_Orden = " & CLng(Val(txtNroOrden)) & " "
Conexion.SendHost mstrSql, , , , gcTiempoEspera

With lvwDetalle
    If .ListItems.Count > 0 Then
        For x = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(x)
            mstrSql = "INSERT INTO Tllr_Detalle_Orden_Compra ( "
            mstrSql = mstrSql & " Id_Empresa,Id_Sucursal, "
            mstrSql = mstrSql & " Id_Orden, "
            mstrSql = mstrSql & " Item, "
            mstrSql = mstrSql & " Descripcion, "
            mstrSql = mstrSql & " Cantidad, "
            mstrSql = mstrSql & " Precio_Unitario, "
            mstrSql = mstrSql & " SubTotal) "
            mstrSql = mstrSql & " values ( "
            mstrSql = mstrSql & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',  "
            mstrSql = mstrSql & " " & CLng(Val(txtNroOrden)) & ", "
            mstrSql = mstrSql & " " & CInt(Val(.SelectedItem)) & ", "
            mstrSql = mstrSql & " '" & .SelectedItem.SubItems(1) & "', "
            mstrSql = mstrSql & " " & CDbl(Val(SacarFormatoValor(.SelectedItem.SubItems(2), ""))) & ", "
            mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))) & ","
            mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(4), ""))) & ") "
            Conexion.SendHost mstrSql, , , , gcTiempoEspera
        Next
    End If
End With
End Sub
Sub GuardaTerceros()

'If gstrEstadoOT = "V" Then

If Me.txtNroOt.Tag <> "" And Me.txtNroOt <> "" Then
    gstrEstadoOT = Retorna_Valor_General("Select Estado from Tllr_Ot WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "'", gcdynamic)
    If gstrEstadoOT = "V" Then
    With lvwDetalle
        mstrSql = " DELETE FROM Tllr_Terceros_OT WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "' And Id_Servicio_Tercero Like 'OC-" & Me.txtNroOrden & "-%' And Id_Proveedor='" & Me.txtProveedor.Tag & "'"
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        If .ListItems.Count > 0 Then
            For x = 1 To .ListItems.Count
                Set .SelectedItem = .ListItems(x)
                mstrSql = "INSERT INTO Tllr_Terceros_OT ("
                mstrSql = mstrSql & " Id_Empresa,Id_Sucursal, "
                mstrSql = mstrSql & " Id_OT, "
                mstrSql = mstrSql & " Seccion_OT, "
                mstrSql = mstrSql & " Id_Servicio_Tercero,"
                mstrSql = mstrSql & " Id_Proveedor,"
                mstrSql = mstrSql & " Id_Tipo_Cargo,"
                mstrSql = mstrSql & " Descripcion, "
                mstrSql = mstrSql & " Cantidad, "
                mstrSql = mstrSql & " Valor, "
                mstrSql = mstrSql & " Subtotal, "
                mstrSql = mstrSql & " Porcentaje_Recargo, "
                mstrSql = mstrSql & " Monto_Recargo, "
                mstrSql = mstrSql & " Precio_Final, "
                mstrSql = mstrSql & " NroFarctura, "
                mstrSql = mstrSql & " Porcentaje_Dscto, "
                mstrSql = mstrSql & " Monto_Dscto, "
                mstrSql = mstrSql & " Facturado) "
                mstrSql = mstrSql & " values ( "
                mstrSql = mstrSql & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
                mstrSql = mstrSql & " '" & txtNroOt & "',"
                mstrSql = mstrSql & " '" & txtNroOt.Tag & "',"
                mstrSql = mstrSql & " '" & "OC-" & txtNroOrden & "-" & .SelectedItem & "',"
                mstrSql = mstrSql & " '" & txtProveedor.Tag & "',"
                mstrSql = mstrSql & " '" & TraeCargoOT(txtNroOt, txtNroOt.Tag) & "'," 'IIf(gstrIdCargo = "", gstrIdCargo, gstrIdCargoDefecto) & "',"
                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(1) & "', "
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(2), ""))) & ","
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))) & ","
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(4), ""))) & ","
                mstrSql = mstrSql & 0 & ","
                mstrSql = mstrSql & 0 & ","
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))) & ","
                mstrSql = mstrSql & "'0',"
                mstrSql = mstrSql & 0 & ","
                mstrSql = mstrSql & 0 & ","
                mstrSql = mstrSql & " 'N')"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
            Next
        End If
    End With
    Else
        MsgBox "El Detalle de la Orden de Compra no fue grabada en la OT " & txtNroOt & ". Esta Ot No Esta VIGENTE...", vbInformation, "Información"
    End If
End If
End Sub
Sub GuardaCarroceria()

'If gstrEstadoOT = "V" Then
mstrSql = " DELETE FROM Tllr_Carroceria_OT WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "' And Id_Servicio_Carroceria Like 'OC-" & Me.txtNroOrden & "-%' And Id_Proveedor='" & Me.txtProveedor.Tag & "'"
'kjcv 04.04.17
'mstrSql = " DELETE FROM Tllr_Carroceria_OT WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "' And Id_Servicio_Carroceria Like 'OC-" & Me.txtNroOrden & "-%' "
Conexion.SendHost mstrSql, , , , gcTiempoEspera

If Me.txtNroOt.Tag <> "" And Me.txtNroOt <> "" Then
    gstrEstadoOT = Retorna_Valor_General("Select Estado from Tllr_Ot WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "'", gcdynamic)
    If gstrEstadoOT = "V" Then
    With lvwDetalle
        If .ListItems.Count > 0 Then
            For x = 1 To .ListItems.Count
                Set .SelectedItem = .ListItems(x)
                mstrSql = "INSERT INTO Tllr_Carroceria_Ot"
                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
                mstrSql = mstrSql & " Id_Compañia_Seguro, "
                mstrSql = mstrSql & " Id_Concepto, "
                mstrSql = mstrSql & " D_P,"
                mstrSql = mstrSql & " Id_Parte_Pieza, "
                mstrSql = mstrSql & " Id_Tipo_Cargo, Mecanico_Designado,"
                mstrSql = mstrSql & " Horas, Valor,Valor_Definido ,"
                mstrSql = mstrSql & " Porcentaje_Descuento,Monto_Descuento,"
                mstrSql = mstrSql & " SubTotal,Facturado,Porcentaje_Recargo,Monto_Recargo,Id_Proveedor,Descripcion,Id_Servicio_Carroceria)"
                mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', "   '///empresa, sucursal
                mstrSql = mstrSql & " '" & Me.txtNroOt.Text & "', '" & Me.txtNroOt.Tag & "',"       '///nro ot, seccion
                mstrSql = mstrSql & " '" & "1" & "', "                                              '///cia seguro
                mstrSql = mstrSql & " '" & "01" & "', "                                             '///concepto
                mstrSql = mstrSql & " '" & "" & "',"
                mstrSql = mstrSql & " '" & "01" & "', "                                             '///parte y pieza
                mstrSql = mstrSql & " '" & TraeCargoOT(txtNroOt, txtNroOt.Tag) & "','01',"           '///mecanico designado
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(2), ""))) & ","
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))) & ","
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(3), ""))) & ","
                mstrSql = mstrSql & "0" & ","
                mstrSql = mstrSql & "0" & ","
                mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(.SelectedItem.SubItems(4), ""))) & ","
                mstrSql = mstrSql & " 'N',"
                mstrSql = mstrSql & "0" & ","
                mstrSql = mstrSql & "0" & ","
                mstrSql = mstrSql & " '" & txtProveedor.Tag & "',"
                mstrSql = mstrSql & " '" & .SelectedItem.SubItems(1) & "', "
                mstrSql = mstrSql & " '" & "OC-" & txtNroOrden & "-" & .SelectedItem & "')"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
            Next
        End If
    End With
    Else
        MsgBox "El Detalle de la Orden de Compra no fue grabada en la OT " & txtNroOt & ". Esta Ot No Esta VIGENTE...", vbInformation, "Información"
    End If
End If
End Sub

Sub ReIndexItem()

For x = 1 To lvwDetalle.ListItems.Count
    Set lvwDetalle.SelectedItem = lvwDetalle.ListItems(x)
    lvwDetalle.SelectedItem = CStr(x)
Next
End Sub

Private Sub Form_Load()
mblnSW = True
Me.Label1.Caption = gstrNombreIva
End Sub



Private Sub tblOT_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    frmBuscaOT.Show vbModal
    txtNroOt = gstrBusca
    txtNroOt.Tag = gstrSeccion
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
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_20_0040", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
''    RevizaAtributos
'        If gapAccion = apcrear Then
'           AgregarRegistro
''           txtCodigo = gstrBusca
'        End If
'        If gapAccion = apeditar Then
'            If gstrBusca <> "" Then
'                mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
'                If Conexion.SendHost(mstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
'                        LeerCampos
'                    End If
'                End If
'                Conexion.CloseHost AdoPrincipal
'            End If
''            txtCodigo.Enabled = False
'            Me.SetFocus
'        End If
'        If gapAccion = apninguno Then
'           Renovar
'        End If
    End If
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    txtNroOrden = "?"   'CorrelativoOrdenCompra(gstrIdEmpresa, gstrIdSucursal)
    txtContacto.SetFocus
'    gapAccion = apninguno
'    mblnSW = False
'    Screen.MousePointer = vbDefault
End Sub
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    txtNroOrden = "?"  'CorrelativoOrdenCompra(gstrIdEmpresa, gstrIdSucursal)
    txtContacto.SetFocus
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones

'    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & " > " & IIf(txtNroOrden = "?", 0, txtNroOrden) & " And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
    mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & " > " & IIf(txtNroOrden = "?", 0, txtNroOrden) & " And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
'            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<" & IIf(txtNroOrden = "?", 0, txtNroOrden) & " And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
            mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<" & IIf(txtNroOrden = "?", 0, txtNroOrden) & " And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
            
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost adoPrincipal

End Sub
Private Sub GrabarRegistro()
Dim curNumeroOrden As Currency
Dim adoOrden As New ADODB.Recordset
Dim Sql As String
Dim flag As Boolean
Dim Rpta As Boolean

If Not Validacion() Then
    Exit Sub
End If
mstrEstado = "V"
flag = False
Sql = "select " & mcCampoCodigo & " from " & mcNombreTabla & " where proveedor='" & txtProveedor.Tag & "' and NroOT='" & txtNroOt & "'"
If Conexion.SendHost(Sql, adoOrden, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoOrden.BOF And Not adoOrden.EOF Then
                flag = True
            End If
End If
Conexion.CloseHost adoOrden
If flag = True Then
    If MsgBox("Existe Orden de Compra con Proveedor a la misma OT.¿Desea generar otra Orden Compra ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then Rpta = True
End If

If flag = False Or Rpta = True Then

If Me.Tag = "Crear" Then
    mstrSql = "INSERT INTO " & mcNombreTabla & " (" & mcCampoCodigo & ", "
    mstrSql = mstrSql & " Id_Empresa,Id_Sucursal, "
    mstrSql = mstrSql & " Proveedor, "
    mstrSql = mstrSql & " Contacto, "
    mstrSql = mstrSql & " Condicion_Pago, "
    mstrSql = mstrSql & " Fecha_Orden, "
    mstrSql = mstrSql & " Observacion, "
    mstrSql = mstrSql & " SubTotal, "
    mstrSql = mstrSql & " Porcentaje_Descuento,"
    mstrSql = mstrSql & " Descuento, "
    mstrSql = mstrSql & " Neto, "
    mstrSql = mstrSql & " Iva, "
    mstrSql = mstrSql & " Total, "
    mstrSql = mstrSql & " Estado, "
    mstrSql = mstrSql & " NroOT, SeccionOT,Terceros) "
    curNumeroOrden = CorrelativoOrdenCompra(gstrIdEmpresa, gstrIdSucursal)  'rescato siguiente número orden
    mstrSql = mstrSql & " values (" & curNumeroOrden & ", "
    mstrSql = mstrSql & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',  "
    mstrSql = mstrSql & " '" & txtProveedor.Tag & "', "
    mstrSql = mstrSql & " '" & txtContacto & "', "
    mstrSql = mstrSql & " '" & txtCondCompra & "', "
    mstrSql = mstrSql & " '" & CDate(pckFecha.Value) & "',"
    mstrSql = mstrSql & " '" & txtObs & "', "
    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtSubTotal, ""))) & ", "
    mstrSql = mstrSql & " " & CDbl(Val(txtPorDes)) & ", "
    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtDescuento, ""))) & ", "
    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtneto, ""))) & ","
    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtiva, ""))) & ","
    mstrSql = mstrSql & " " & CCur(Val(SacarFormatoValor(txtTotal, ""))) & ", "
    mstrSql = mstrSql & " '" & mstrEstado & "',"
    mstrSql = mstrSql & " '" & txtNroOt & "','" & txtNroOt.Tag & "',"
    mstrSql = mstrSql & " '" & IIf(Me.optTerceros.Value = True, "S", "N") & "')"
    txtNroOrden = curNumeroOrden
Else
    mstrSql = "UPDATE " & mcNombreTabla & " "
    mstrSql = mstrSql & " SET Proveedor='" & txtProveedor.Tag & "', "
    mstrSql = mstrSql & " Contacto='" & txtContacto & "', "
    mstrSql = mstrSql & " Condicion_Pago='" & txtCondCompra & "', "
    mstrSql = mstrSql & " Fecha_Orden='" & CDate(pckFecha.Value) & "', "
    mstrSql = mstrSql & " Observacion='" & txtObs & "', "
    mstrSql = mstrSql & " SubTotal=" & CCur(Val(SacarFormatoValor(txtSubTotal, ""))) & ", "
    mstrSql = mstrSql & " Porcentaje_Descuento=" & CDbl(Val(txtPorDes)) & ","
    mstrSql = mstrSql & " Descuento=" & CCur(Val(SacarFormatoValor(txtDescuento, ""))) & ", "
    mstrSql = mstrSql & " Neto=" & CCur(Val(SacarFormatoValor(txtneto, ""))) & ", "
    mstrSql = mstrSql & " Iva=" & CCur(Val(SacarFormatoValor(txtiva, ""))) & ", "
    mstrSql = mstrSql & " Total=" & CCur(Val(SacarFormatoValor(txtTotal, ""))) & ", "
    mstrSql = mstrSql & " Estado='" & mstrEstado & "', NroOT = '" & txtNroOt & "', SeccionOT = '" & txtNroOt.Tag & "',"
    mstrSql = mstrSql & " Terceros='" & IIf(Me.optTerceros.Value = True, "S", "N") & "'"
    mstrSql = mstrSql & " WHERE " & mcCampoCodigo & "=" & CLng(Val(txtNroOrden)) & ""
End If
If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
    GuardaDetalle
    If optTerceros.Value = True Then
        GuardaTerceros
    Else
        GuardaCarroceria
    End If
    mblnTablaVacia = False
    ActivaBotones
    Me.Tag = ""
End If

End If

End Sub
Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea anular este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        
        'valida que lo ot no este liquidada o facturada
        If Me.txtNroOt.Tag <> "" Then
            If Retorna_Valor_General("Select Estado from Tllr_Ot where Seccion_Ot='" & Me.txtNroOt.Tag & "' And Id_Ot='" & Me.txtNroOt & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Empresa='" & gstrIdEmpresa & "'", gcdynamic) <> "V" Then
                MsgBox "No se puede eliminar un OC a Terceros, porque la OT Ya no esta vigente...", vbInformation, "Advertencia"
                Exit Sub
            End If
        End If
            
        If Me.optTerceros.Value = True Then
'            mstrSQL = " DELETE FROM Tllr_Terceros_OT WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "' And Id_Servicio_Tercero Like 'OC%' And Id_Proveedor='" & Me.txtProveedor.Tag & "' And Facturado='N'"
            'kjcv 15.06.15
            mstrSql = " DELETE FROM Tllr_Terceros_OT WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "' And Id_Servicio_Tercero Like 'OC-" & Trim(txtNroOrden) & "%" & "' And Id_Proveedor='" & Me.txtProveedor.Tag & "' And Facturado='N'"
            Conexion.SendHost mstrSql, , , , gcTiempoEspera
        Else
            'mstrSQL = " DELETE FROM Tllr_Carroceria_OT WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "' And Id_Servicio_Carroceria Like 'OC%' And Id_Proveedor='" & Me.txtProveedor.Tag & "' And Facturado='N'"
            'kjcv 15.06.15
            mstrSql = " DELETE FROM Tllr_Carroceria_OT WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Id_Sucursal = '" & gstrIdSucursal & "' AND Id_OT = '" & txtNroOt & "' And Seccion_OT = '" & txtNroOt.Tag & "' And Id_Servicio_Carroceria Like 'OC-" & Trim(txtNroOrden) & "%" & "' And Id_Proveedor='" & Me.txtProveedor.Tag & "' And Facturado='N'"
            Conexion.SendHost mstrSql, , , , gcTiempoEspera
        End If

'        mstrSql = "DELETE FROM Tllr_Detalle_Orden_Compra where " & mcCampoCodigo & "='" & txtNroOrden & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
'        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
            'mstrSql = "DELETE FROM " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtNroOrden & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrSql = "UPDATE " & mcNombreTabla & " SET Estado='N' where " & mcCampoCodigo & "='" & txtNroOrden & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                        
            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
'                mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and " & mcCampoCodigo & ">'" & txtNroOrden & "' order by " & mcCampoCodigo
                mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and " & mcCampoCodigo & ">'" & txtNroOrden & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        LeerCampos
                    Else
                        'mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and " & mcCampoCodigo & "<'" & txtNroOrden & "' order by " & mcCampoCodigo
                        mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and " & mcCampoCodigo & "<'" & txtNroOrden & "' order by " & mcCampoCodigo
                        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                                LeerCampos
                            Else
                                mblnTablaVacia = True
                                LimpiaCampos
                            End If
                        End If
                    End If
                End If
            End If
'            Conexion.CloseHost adoPrincipal
'         End If
    End If
End Sub

Private Sub BuscarRegistro()
    'gstrBusca = BuscarRegistros(mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption)
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption)
'    If gstrBusca <> "" Then
'        mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
'        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
'                LeerCampos
'            End If
'        End If
'        Conexion.CloseHost adoPrincipal
'    End If
'    Me.SetFocus

    Screen.MousePointer = 1
    frmInfOrdCom.Show vbModal
    Screen.MousePointer = 1
    If gstrBusca <> "" Then
        'mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
        mstrSql = "select case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
    Me.SetFocus


End Sub
Private Sub ImprimirInforme()
'kjcv 17.07.15
Dim H As Integer
Dim AdoCabeza As New ADODB.Recordset
Dim AdoDetalle As New ADODB.Recordset
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim TablaDetalle As DAO.Recordset
Dim lstrSQL As String
Dim mstrSql As String

On Error GoTo Cancela_impresion

    Screen.MousePointer = vbHourglass
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(gstrPathReporte + "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb"  ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral)  ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE ( Proveedor  TEXT, Direccion TEXT, Telefono Text, Contacto Text,CondPago Text, Fax Text, Sec Text ,SubTotal float,Dscto float,Neto float,Iva float, Total float,Observaciones Text,Numero double,fecha text)"
    
lstrSQL = lstrSQL & " SELECT Glbl_Cliente_Proveedor.Razon_Social,Glbl_Cliente_Proveedor.Direccion,Glbl_Cliente_Proveedor.Telefono,"
lstrSQL = lstrSQL & " Glbl_Cliente_Proveedor.Fax,Tllr_Orden_Compra.Id_Orden,Tllr_Orden_Compra.Fecha_Orden,Tllr_Orden_Compra.Condicion_Pago,"
lstrSQL = lstrSQL & " Tllr_Orden_Compra.Contacto,Tllr_Orden_Compra.SubTotal,Tllr_Orden_Compra.Descuento, Tllr_Orden_Compra.Neto,"
lstrSQL = lstrSQL & " Tllr_Orden_Compra.IVA , Tllr_Orden_Compra.Total, Tllr_Orden_Compra.SeccionOT, Tllr_Orden_Compra.Observacion"
lstrSQL = lstrSQL & " FROM Tllr_Orden_Compra inner join Glbl_Cliente_Proveedor"
lstrSQL = lstrSQL & " ON Tllr_Orden_Compra.Proveedor=Glbl_Cliente_Proveedor.Id_Cliente_Proveedor"
lstrSQL = lstrSQL & " WHERE Tllr_Orden_Compra.Id_Empresa='" & gstrIdEmpresa & "' AND Tllr_Orden_Compra.Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_Orden_Compra.Id_Orden='" & txtNroOrden & "'"
If Conexion.SendHost(lstrSQL, AdoCabeza, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
     If Not AdoCabeza.BOF And Not AdoCabeza.EOF Then
         Set Tabla = Dbsnueva.OpenRecordset("select * from T_PARAREPORTE")
         Tabla.AddNew
         Tabla!Proveedor = AdoCabeza!Razon_Social
         Tabla!Direccion = AdoCabeza!Direccion
         Tabla!Telefono = AdoCabeza!Telefono
         Tabla!Contacto = AdoCabeza!Contacto
         Tabla!CondPago = AdoCabeza!Condicion_Pago
         Tabla!fax = AdoCabeza!fax
         Tabla!Sec = AdoCabeza!SeccionOT
         Tabla!SubTotal = AdoCabeza!SubTotal
         Tabla!Dscto = AdoCabeza!Descuento
         Tabla!Neto = AdoCabeza!Neto
         Tabla!IVA = AdoCabeza!IVA
         Tabla!Total = AdoCabeza!Total
         Tabla!Observaciones = AdoCabeza!Observacion
         Tabla!NUMERO = AdoCabeza!Id_Orden
         Tabla!Fecha = AdoCabeza!Fecha_Orden
         Tabla.Update
         Tabla.Close
    End If
    
End If
    
Conexion.CloseHost AdoCabeza
    
Dbsnueva.Execute "CREATE TABLE T_PARAREPORTEDETALLE (Codigo  TEXT, Pieza TEXT, Cantidad float,PUnitario float, SUBTOTAL float)"
    
mstrSql = " SELECT Tllr_Detalle_Orden_Compra.Item,Tllr_Detalle_Orden_Compra.Descripcion,Tllr_Detalle_Orden_Compra.Cantidad,"
mstrSql = mstrSql & " Tllr_Detalle_Orden_Compra.Precio_Unitario , Tllr_Detalle_Orden_Compra.SubTotal"
mstrSql = mstrSql & " FROM Tllr_Orden_Compra inner join Tllr_Detalle_Orden_Compra"
mstrSql = mstrSql & " ON Tllr_Orden_Compra.Id_Orden=Tllr_Detalle_Orden_Compra.Id_Orden  and Tllr_Orden_Compra.Id_Sucursal=Tllr_Detalle_Orden_Compra.Id_Sucursal "
mstrSql = mstrSql & " WHERE Tllr_Detalle_Orden_Compra.Id_Empresa='" & gstrIdEmpresa & "' AND Tllr_Detalle_Orden_Compra.Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_Detalle_Orden_Compra.Id_Orden='" & txtNroOrden & "'"
    If Conexion.SendHost(mstrSql, AdoDetalle, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        While Not AdoDetalle.EOF
            Set TablaDetalle = Dbsnueva.OpenRecordset("select * from T_PARAREPORTEDETALLE")
            TablaDetalle.AddNew
            TablaDetalle!Codigo = AdoDetalle!item
            TablaDetalle!pieza = AdoDetalle!Descripcion
            TablaDetalle!cantidad = AdoDetalle!cantidad
            TablaDetalle!PUnitario = AdoDetalle!Precio_Unitario
            TablaDetalle!SubTotal = AdoDetalle!SubTotal
            TablaDetalle.Update
            AdoDetalle.MoveNext
        Wend
        TablaDetalle.Close
   End If
     
      
   Dbsnueva.Close
   
Dim Result As DAO.Recordset


With rptOrdenCompra
    .ReportFileName = gstrPathReporte & "\OrdenCServicio.rpt"
    .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
    .WindowState = crptMaximized
    '.WindowShowPrintSetupBtn = True
    '.WindowShowPrintBtn = True
    .Formulas(0) = "NROORDEN=" & CLng(Val(txtNroOrden))
    .Formulas(1) = "IDSUCURSAL='" & gstrIdSucursal & "'"
    .Formulas(2) = "NombreRut='" & gstrNombreRut & "'"
    .Formulas(3) = "NombreSucursal='" & gstrNombreSucursal & "'"
    .Formulas(4) = "NombreIva='" & gstrNombreIva & "'"
    .Formulas(5) = "Tdecimales=" & gintDecimalesMoneda
    .Formulas(6) = "EditaRut='" & gstrEditaRut & "'"
    .Formulas(7) = "Razonsocial='" & gstrEmpresa & "'"
    .Formulas(8) = "Direccion='" & gstrDirSuc & "'"
    .Formulas(9) = "OT='" & txtNroOt & "'"
    .Formulas(10) = "RUT='" & Me.txtProveedor.Tag & "'"
    '.ProgressDialog = True
    .Destination = crptToWindow

'    .Connect = "Driver={SQL Server};Server=wiracocha;UID=sa;PWD=Llosa1936;Database=elisa;" 'Conexion.ConnectionString
'    .SelectionFormula = "{Tllr_Detalle_Orden_Compra.Id_Empresa}='" & gstrIdEmpresa & "' AND {Tllr_Detalle_Orden_Compra.Id_Sucursal}='" & gstrIdSucursal & "' AND {Tllr_Detalle_Orden_Compra.Id_Orden}=" & CLng(Val(txtNroOrden)) & " AND {Tllr_Orden_Compra.Id_Empresa}='" & gstrIdEmpresa & "' AND {Tllr_Orden_Compra.Id_Sucursal}='" & gstrIdSucursal & "' And {Tllr_Orden_Compra.Id_Orden}=" & CLng(Val(txtNroOrden))
    .Action = True
End With

   Screen.MousePointer = vbDefault

Exit Sub

Cancela_impresion:

  If Err.Number = 32755 Then
      MsgBox "Impresión Cancelada", vbInformation, "Imprimiendo"
  End If
  Screen.MousePointer = vbDefault

End Sub
Private Sub PrimerRegistro()
'mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' order by " & mcCampoCodigo
mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' order by " & mcCampoCodigo
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroAnterior()
'mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  and  " & mcCampoCodigo & "<" & txtNroOrden & " order by " & mcCampoCodigo & " DESC"
mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  and  " & mcCampoCodigo & "<" & txtNroOrden & " order by " & mcCampoCodigo & " DESC"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroSiguiente()
'mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  and " & mcCampoCodigo & ">" & txtNroOrden & " order by " & mcCampoCodigo
mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  and " & mcCampoCodigo & ">" & txtNroOrden & " order by " & mcCampoCodigo
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub UltimoRegistro()
'mstrSql = "select TOP 1 * from " & mcNombreTabla & "  where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  order by " & mcCampoCodigo & " DESC"
mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & "  where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  order by " & mcCampoCodigo & " DESC"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub Renovar()
Set adoPrincipal = New ADODB.Recordset
'mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  order by " & mcCampoCodigo
mstrSql = "select TOP 1 case  Estado when '' then '' when 'V' then 'VIGENTE' when 'N' then 'NULA' end as EstadoDes, * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  order by " & mcCampoCodigo
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    VerificaTablaVacia
    ActivaBotones
    If Not mblnTablaVacia Then
        PrimerRegistro
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
'    txtCodigo.Enabled = False
    With tlbBarraHerramientas.Buttons
        .item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
        .item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoEditar, True, False))
        .item("Cancelar").Enabled = False
        .item("Borrar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoBorrar, True, False))
        .item("Buscar").Enabled = IIf(mblnTablaVacia, False, True)
        .item("Imprimir").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoImprimir, True, False))
        .item("Primero").Enabled = IIf(mblnTablaVacia, False, True)
        .item("Anterior").Enabled = IIf(mblnTablaVacia, False, True)
        .item("Siguiente").Enabled = IIf(mblnTablaVacia, False, True)
        .item("Ultimo").Enabled = IIf(mblnTablaVacia, False, True)
        .item("Renovar").Enabled = True
        .item("Cerrar").Enabled = True
    End With
    
End Sub
Private Sub DesactivaBotones()
'    txtCodigo.Enabled = True
    txtNroOrden = 1     'nro correlativo para la sucursal
    With tlbBarraHerramientas.Buttons
        .item("Crear").Enabled = False
        .item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
        .item("Cancelar").Enabled = True
        .item("Borrar").Enabled = mblnAccesoBorrar 'False
        .item("Buscar").Enabled = False
        .item("Imprimir").Enabled = False
        .item("Primero").Enabled = False
        .item("Anterior").Enabled = False
        .item("Siguiente").Enabled = False
        .item("Ultimo").Enabled = False
        .item("Renovar").Enabled = False
        .item("Cerrar").Enabled = True
    End With
End Sub
Private Sub VerificaTablaVacia()
    If (Not adoPrincipal.BOF And Not adoPrincipal.EOF) And adoPrincipal.RecordCount > 0 Then
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

With adoPrincipal

    txtNroOrden = !Id_Orden
    txtProveedor.Tag = !Proveedor
    txtProveedor = ClienteDes(!Proveedor)
    txtContacto = ValorNulo(!Contacto)
    txtCondCompra = ValorNulo(!Condicion_Pago)
    pckFecha.Value = !Fecha_Orden
    txtObs = ValorNulo(!Observacion)
    txtNroOt = ValorNulo(!NroOT)
    txtNroOt.Tag = ValorNulo(!SeccionOT)
    txtSubTotal = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
    txtPorDes = !porcentaje_descuento
    txtDescuento = FormatoValor(!Descuento, "", gintDecimalesMoneda)
    txtneto = FormatoValor(!Neto, "", gintDecimalesMoneda)
    
    txtiva = FormatoValor(!IVA, "", gintDecimalesMoneda)
    txtTotal = FormatoValor(!Total, "", gintDecimalesMoneda)
    
    If !estado = "N" Then
        tlbBarraHerramientas.Buttons.item("Borrar").Enabled = False
    Else
        tlbBarraHerramientas.Buttons.item("Borrar").Enabled = mblnAccesoBorrar
    End If
    
    lblestado = ValorNulo(!estadoDes)
    
    If IsNull(!Terceros) Then
        Me.optTerceros.Value = True
        Me.optCarroce.Value = True
    ElseIf !Terceros = "N" Then
        optTerceros.Value = False
        Me.optCarroce.Value = True
    Else
        optTerceros.Value = True
        Me.optCarroce.Value = False
    End If
    
    Call DetalleOrden(gstrIdEmpresa, gstrIdSucursal, !Id_Orden)
End With
End Sub


Private Sub LimpiaCampos()
With Me
    .txtContacto = ""
    .txtProveedor = ""
    .txtProveedor.Tag = ""
    .txtCondCompra = ""
    .txtObs = ""
    .txtNroOt = ""
    .txtNroOt.Tag = ""
    .lvwDetalle.ListItems.Clear
    .txtneto = ""
    .txtDescuento = ""
    .txtiva = ""
    .txtSubTotal = ""
    .txtTotal = ""
End With
End Sub
Private Sub ValoresporDefecto()
    With Me
        .pckFecha.Value = CDate(Now)
        .txtSubTotal = "0"
        .txtPorDes = "0"
        .txtDescuento = "0"
        .txtneto = "0"
        .txtiva = "0"
        .txtTotal = "0"
    End With
End Sub
Private Function Validacion() As Boolean


    Validacion = True
    If txtNroOrden = "" Then
        MsgBox "El Nro de Orden Debe Especificarse...", vbInformation, "Advertencia"
        Validacion = False
        Exit Function
    End If
    If Me.txtProveedor.Tag = "" Then
        MsgBox "El Proveedor debe Ser Especificado...", vbInformation, "Advertencia"
        txtProveedor.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtNroOt = "" Then
        MsgBox "El Numero de la OT debe Ser Especificado...", vbInformation, "Advertencia"
        txtNroOt.SetFocus
        Validacion = False
        Exit Function
    End If

    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As New ADODB.Recordset
        mstrSql = "select " & mcCampoCodigo & " from " & mcNombreTabla & " where " & mcCampoCodigo & "=" & IIf(txtNroOrden = "?", 0, txtNroOrden) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este Nro. de Orden  ya esta registrado ", vbInformation, "Advertencia"
                Validacion = False
                'txtCodigo.SetFocus
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
    
    
    
    
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmEmisionOrdCom = Nothing
    gstrBusca = txtNroOrden
End Sub
Private Sub RevizaAtributos()

    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True

End Sub


Private Sub tlbItemOrdCom_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Agregar"
    frmAddItemOrdCom.Show 1
    Totales
Case "Quitar"
    If lvwDetalle.ListItems.Count > 0 Then
        lvwDetalle.ListItems.Remove lvwDetalle.SelectedItem.Index
        ReIndexItem
        Totales
    End If
End Select

End Sub

Private Sub tlbProveedor_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim mstrnombre As String

Select Case Button.Key
Case "Nuevo"
    'gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, "", "", mstrnombre, apcrear)
    gstrBusca = apfFormulario.clientes(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, gstrPathReporte, "", "", apcrear, "Cliente - Proveedor", gstrIdSucursal)
    txtProveedor.Tag = gstrBusca
    txtProveedor = mstrnombre
Case "Buscar"
'    apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre, gstrIdEmpresa
'    'apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre
'    txtProveedor.Tag = gstrBusca
'    txtProveedor = mstrnombre
'    txtContacto.SetFocus
    'kjcv 02-02-2012
    gstrRutCliente = ""
    gstrNombreCliente = ""
    Libreria.ClienteBuscar Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
         If gstrRutCliente <> "" Then
            Me.txtProveedor = gstrNombreCliente
            Me.txtProveedor.Tag = gstrRutCliente
            txtContacto.SetFocus
        End If
End Select
End Sub


Private Sub txtCondCompra_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub


Private Sub txtContacto_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtdescuento_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPorDes, strDot)
End Sub

Private Sub txtDescuento_LostFocus()
txtPorDes = basFunciones.PorcentajeMonto(Val(SacarFormatoValor(txtSubTotal, "")), CSng(Val(txtDescuento)))
Totales
End Sub


Private Sub txtPorDes_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPorDes, strDot)
End Sub


Private Sub txtPorDes_LostFocus()
txtDescuento = ValorPorcentaje(Val(SacarFormatoValor(txtSubTotal, "")), CSng(Val(txtPorDes)))
Totales
End Sub


Private Sub txtProveedor_GotFocus()
MarcaTexto txtProveedor
End Sub

Private Sub txtProveedor_LostFocus()
txtProveedor.Tag = txtProveedor
txtProveedor = ProveedorS(txtProveedor)
If txtProveedor = "" Then
    txtProveedor.Tag = ""
    txtProveedor = ""
End If

End Sub
