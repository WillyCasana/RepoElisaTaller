VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Frmmargenrepuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Margen de Repuestos"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "FrmMargenRepuestos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptOT 
      Left            =   3945
      Top             =   6210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   945
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   6720
      Begin VB.OptionButton optcarroceria 
         Caption         =   "Carrocería"
         Height          =   264
         Left            =   4050
         TabIndex        =   7
         Top             =   525
         Width           =   1104
      End
      Begin VB.OptionButton optmecanica 
         Caption         =   "Mecánica"
         Height          =   264
         Left            =   5325
         TabIndex        =   6
         Top             =   525
         Value           =   -1  'True
         Width           =   1188
      End
      Begin VB.TextBox txtNroOt 
         Height          =   300
         Left            =   105
         MaxLength       =   15
         TabIndex        =   3
         Top             =   525
         Width           =   2670
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
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
               Picture         =   "FrmMargenRepuestos.frx":000C
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":011E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":0576
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":09CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":0E26
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":0F38
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":104A
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":115C
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":126E
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":1380
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":1492
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":15A4
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":16B6
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":17C8
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":18DA
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":19EC
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":1AFE
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":1C10
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":1D22
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":1E34
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":2286
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMargenRepuestos.frx":26D8
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
         ImageList       =   "ImgBarraHerramienta"
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
      Left            =   75
      TabIndex        =   0
      Top             =   1425
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
      Appearance      =   1
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
      Left            =   75
      TabIndex        =   5
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
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Crear"
            Object.ToolTipText     =   "Nueva búsqueda"
            ImageIndex      =   1
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar "
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "CotizaPerdida"
            Object.ToolTipText     =   "Cotización Perdida"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Vender"
            Object.ToolTipText     =   "Pasar a Venta"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Agenda"
            Object.ToolTipText     =   "Agenda Diaria"
            ImageIndex      =   22
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
            Object.ToolTipText     =   "Cerrar "
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   750
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
            Picture         =   "FrmMargenRepuestos.frx":27EA
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":28FC
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":2A0E
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":2B20
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":2C32
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":2D44
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":2E56
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":2F68
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":307A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":318C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":329E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":33B0
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":34C2
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":35D4
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":36E6
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":37F8
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":390A
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":3D5C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":41AE
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":42C0
            Key             =   "Foto"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":43D4
            Key             =   "Venta"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":44D0
            Key             =   "Agenda"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":466C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMargenRepuestos.frx":4808
            Key             =   "PASAVENTA"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbTotales 
      Height          =   315
      Left            =   900
      TabIndex        =   8
      Top             =   6225
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Suma - Totales"
            TextSave        =   "Suma - Totales"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Suma - Costos"
            TextSave        =   "Suma - Costos"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Suma - Margen"
            TextSave        =   "Suma - Margen"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2646
            MinWidth        =   2646
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
Attribute VB_Name = "Frmmargenrepuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
Sub ImprimirConsulta()
Dim Dbsnueva As Database
Dim tabla As DAO.Recordset
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
    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Pieza text,Familia text,cantidad text,valor text,pdescto text,descto text,subtotal text,preciocosto text,Margen text,Pmargen text)"
    Set tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        tabla.AddNew
        tabla!pieza = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        tabla!Familia = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        tabla!cantidad = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        tabla!Valor = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        tabla!pdescto = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        tabla!descto = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        tabla!subtotal = (IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6)))
        tabla!preciocosto = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        tabla!margen = IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
        tabla!pmargen = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
        tabla.Update
    Next i
   tabla.Close
   
   With rptOT
        .ReportFileName = gstrPathReporte & "\MARGENREPUESTOS.rpt"
        .WindowTitle = "Margen de Repuestos"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
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
        .Destination = crptToWindow
        .Action = True
   End With
   
   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub




Sub cmdBuscarOT_Click()
Dim mstrSql As String
Dim mstrWhere As String
Dim adoTemp As ADODB.Recordset
Dim AdoAux As ADODB.Recordset
Dim itmItem As ListItem


    lvDetalle.ListItems.Clear
mstrWhere = ""
'With Me
    
    mstrSql = "SELECT Stck_Item.Prefijo + Stck_Item.Basico + Stck_Item.Sufijo AS Expr1," _
     & "Glbl_Familia.Descripcion, Tllr_Repuestos_OT.Cantidad,Tllr_Repuestos_OT.Valor, " _
     & "Tllr_Repuestos_OT.Porcentaje_Descuento,Tllr_Repuestos_OT.Monto_Descuento, " _
     & "Tllr_Repuestos_OT.SubTotal , Stck_Item.Precio_Costo FROM Tllr_Repuestos_OT INNER JOIN Tllr_OT ON " _
     & "Tllr_Repuestos_OT.Id_Empresa = Tllr_OT.Id_Empresa AND " _
     & "Tllr_Repuestos_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND " _
     & "Tllr_Repuestos_OT.Id_OT = Tllr_OT.Id_OT AND " _
     & "Tllr_Repuestos_OT.Seccion_OT = Tllr_OT.Seccion_OT INNER JOIN " _
     & "Stck_Item ON " _
     & "Tllr_Repuestos_OT.Id_Item = Stck_Item.Id_Item INNER JOIN " _
     & "Glbl_Familia ON Stck_Item.Id_Familia = Glbl_Familia.Id_Familia Where Tllr_Repuestos_Ot.Id_Ot='" & Me.txtNroOt & "' and Tllr_Repuestos_Ot.Seccion_Ot='" & IIf(Me.optCarroceria.Value = True, "C", "M") & "'"
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , .Fields(0))
              itmItem.SubItems(1) = ValorNulo(.Fields(1))
              itmItem.SubItems(2) = ValorNulo(.Fields(2)) 'Cantidad
              itmItem.SubItems(3) = ValorNulo(.Fields(3)) 'Valor
              itmItem.SubItems(4) = ValorNulo(.Fields(4)) '% descto
              itmItem.SubItems(5) = ValorNulo(.Fields(5)) 'Monto descto
              itmItem.SubItems(6) = ValorNulo(.Fields(6)) 'Subtotal
              itmItem.SubItems(7) = ValorNulo(.Fields(7)) 'Precio costo
              itmItem.SubItems(8) = ValorNulo(.Fields(6)) - (ValorNulo(.Fields(7)) * ValorNulo(.Fields(2))) 'Margen (subtotal - preciocosto * cantidad)
              itmItem.SubItems(9) = Round((Val(itmItem.SubItems(8)) * 100) / ValorNulo(.Fields(6)), 2) 'Margen Porcentual
              adoTemp.MoveNext
          Wend
       End If
    End With
    End If
    
    With Me.stbTotales
        .Panels(2).Text = FormatoValor(TotalSeccion(lvDetalle, 6), "", 0)
        .Panels(4).Text = FormatoValor(TotalSeccion(lvDetalle, 7), "", 0)
        .Panels(6).Text = FormatoValor(TotalSeccion(lvDetalle, 8), "", 0)
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
        Case 3 And tlbBarraHerramientas.Buttons.item("Cerrar").Enabled
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

