VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmInfReservaRepuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Reserva de Repuestos"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmInfReservaRepuestos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptPatente 
      Left            =   3960
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir Informe"
      Height          =   360
      Left            =   7995
      TabIndex        =   21
      Top             =   7200
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   60
      TabIndex        =   5
      Top             =   -15
      Width           =   11370
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   525
         Left            =   5640
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   4680
         Begin VB.OptionButton optLiquidada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Liquidada"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1746
            TabIndex        =   27
            Top             =   240
            Width           =   990
         End
         Begin VB.OptionButton optNula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Nula"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   26
            Top             =   240
            Width           =   675
         End
         Begin VB.OptionButton optCerrada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Facturadas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2739
            TabIndex        =   25
            Top             =   240
            Width           =   1110
         End
         Begin VB.OptionButton optTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton optVigente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Vigente"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   888
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Fin)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1800
         TabIndex        =   20
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   6
         TabIndex        =   14
         Top             =   525
         Width           =   1020
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Patente"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Marca "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   12
         Top             =   300
         Width           =   870
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Modelo"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   11
         Top             =   300
         Width           =   840
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1185
         Width           =   5175
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   8
         Top             =   525
         Width           =   3435
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5640
         MaxLength       =   50
         TabIndex        =   7
         Top             =   525
         Width           =   4635
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Ini)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   10680
         Top             =   120
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
               Picture         =   "frmInfReservaRepuestos.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfReservaRepuestos.frx":3E66
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   450
         _ExtentX        =   794
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
         Enabled         =   0   'False
      End
      Begin MSComctlLib.Toolbar tlbModelo 
         Height          =   330
         Left            =   9840
         TabIndex        =   16
         Top             =   240
         Width           =   345
         _ExtentX        =   609
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
         Enabled         =   0   'False
      End
      Begin MSComctlLib.Toolbar tlbCliente 
         Height          =   330
         Left            =   4920
         TabIndex        =   17
         Top             =   900
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
         Enabled         =   0   'False
      End
      Begin MSComCtl2.DTPicker pckFechaDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   91619329
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1800
         TabIndex        =   19
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   91619329
         CurrentDate     =   36776
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   6240
      TabIndex        =   0
      Top             =   7200
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   1
      Top             =   7200
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   2010
      Left            =   75
      TabIndex        =   4
      Top             =   2160
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   3545
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado/N°Documento"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha Emisión"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Patente"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Marca"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Modelo"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Seccion"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView lvDetalleRepuestosOT 
      Height          =   2010
      Left            =   75
      TabIndex        =   28
      Top             =   4680
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   3545
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   7408
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Familia"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbTotalCosto 
      Height          =   315
      Left            =   8400
      TabIndex        =   30
      Top             =   6720
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Suma - Valores"
            TextSave        =   "Suma - Valores"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2469
            MinWidth        =   2469
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
   Begin VB.Label Label3 
      Caption         =   "Repuestos Usados"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1920
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmInfReservaRepuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
Dim AdoTemp As New ADODB.Recordset
Dim mstrSql As String
Dim itmAux As ListItem

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
    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,FechaIngreso Text,Patente Text,Cliente Text,Marca Text,Modelo Text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!FechaIngreso = IIf(lvDetalle.SelectedItem.SubItems(2) = "", "", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Cliente = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6))
        Tabla.Update
        
        mstrSql = "SELECT (Stck_Item.Prefijo + Stck_Item.Basico + Stck_Item.Sufijo) as Item, "
        mstrSql = mstrSql & "Stck_Item.Descripcion, Tllr_Repuestos_Reservados.Solicitado, "
        mstrSql = mstrSql & "Tllr_Repuestos_Reservados.Precio_Unitario, Tllr_Repuestos_Reservados.Estado, "
        mstrSql = mstrSql & "Glbl_Familia.Descripcion AS Familia, "
        mstrSql = mstrSql & "Tllr_Repuestos_Reservados.Id_OT "
        mstrSql = mstrSql & "FROM Tllr_Repuestos_Reservados INNER JOIN "
        mstrSql = mstrSql & "Stck_Item ON Tllr_Repuestos_Reservados.Id_Item = Stck_Item.Id_Item INNER "
        mstrSql = mstrSql & "Join Glbl_Familia ON Stck_Item.Id_Familia = Glbl_Familia.Id_Familia "
        mstrSql = mstrSql & " WHERE (Tllr_Repuestos_Reservados.Id_Empresa = '" & gstrIdEmpresa & "') AND"
        mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Id_Sucursal = '" & gstrIdSucursal & "') AND"
        mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Id_OT = '" & lvDetalle.SelectedItem & "') AND"
        mstrSql = mstrSql & " (Tllr_Repuestos_Reservados.Seccion_OT = '" & Mid(lvDetalle.SelectedItem.SubItems(7), 1, 1) & "')"
        
        If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            With AdoTemp
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Tabla.AddNew
                    Tabla!NroOT = "          " & !Item
                    Tabla!estado = !Descripcion
                    Tabla!FechaIngreso = !Familia
                    Tabla!Patente = FormatoValor(!Solicitado, "", 2)
                    Tabla!Cliente = FormatoValor(!Precio_Unitario, gstrMonedaLocal, gintDecimalesMoneda)
                    Tabla.Update
                    .MoveNext
                Wend
            End If
            End With
        End If
        Conexion.CloseHost AdoTemp
    Next i
    Tabla.Close
   
    With rptPatente
        .ReportFileName = gstrPathReporte & "\ReservaRepuestos.rpt"
        .WindowTitle = "Reserva de Repuestos"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Reserva de Repuestos'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "desde='" & pckFechaDesde & "'"
        .Formulas(6) = "hasta='" & pckFechaHasta & "'"
        .Formulas(7) = "NombrePlaca='" & gstrNombrePatente & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Connect = Conexion.ConnectionString
        .Action = True
    End With
   
    Screen.MousePointer = 1

End Sub


Private Sub cckCriterios_Click(Index As Integer)
Select Case Index
Case 1
    If cckCriterios(Index).Value = 0 Then
        txtPatente.Enabled = False
        txtPatente = ""
        
    Else
        txtPatente.Enabled = True
        txtPatente.SetFocus
    End If
Case 2
    If cckCriterios(Index).Value = 0 Then
        tlbMarca.Enabled = False
        txtMarca.Enabled = False
        txtMarca = ""
    Else
        tlbMarca.Enabled = True
        txtMarca = ""
        txtMarca.Enabled = True
        txtMarca.SetFocus
    End If
Case 3
    If cckCriterios(Index).Value = 0 Then
        txtModelo.Enabled = False
        txtModelo = ""
    Else
        txtModelo.Enabled = True
        txtModelo.SetFocus
    End If
Case 4
    If cckCriterios(Index).Value = 0 Then
        tlbCliente.Enabled = False
        txtCliente.Enabled = False
        txtCliente = ""
    Else
        tlbCliente.Enabled = True
        txtCliente.Enabled = True
        txtCliente.SetFocus
    End If
Case 5
    If cckCriterios(Index).Value = 0 Then
        'tlbRecep.Enabled = False
        'txtRecepcionista.Enabled = False
        'txtRecepcionista = ""
    Else
        'tlbRecep.Enabled = True
        'txtRecepcionista.Enabled = True
        'txtRecepcionista.SetFocus
    End If
Case 6
    If cckCriterios(Index).Value = 0 Then
        pckFechaDesde.Enabled = False
    Else
        pckFechaDesde.Enabled = True
        pckFechaDesde.SetFocus
    End If
Case 7
    If cckCriterios(Index).Value = 0 Then
        pckFechaHasta.Enabled = False
    Else
        pckFechaHasta.Enabled = True
        pckFechaHasta.SetFocus
    End If
End Select
End Sub
Private Sub cmdBuscarOT_Click()
Dim lstrSql As String
Dim mstrWhere As String
Dim itmItem As ListItem
Dim mstrEstado As String
Dim ContLinea As Integer
Dim mdblSumaHoras As Double
Dim mstrNumeroDocumento As String

lvDetalle.ListItems.Clear
lvDetalleRepuestosOT.ListItems.Clear
mstrWhere = "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "'"
With Me
    
    If .cckCriterios(1).Value = 1 Then  '////////// patente
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(2).Value = 1 Then  '////////// marca
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(3).Value = 1 Then  '////////// modelo
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(4).Value = 1 Then  '////////// cliente
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(6).Value = 1 And .cckCriterios(7).Value = 1 Then   '////////// fecha iniciosi y terminosi
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "','" & pckFechaHasta.Value & " 23:59:00" & "'"
    ElseIf .cckCriterios(6).Value = 0 And .cckCriterios(7).Value = 0 Then  '////////// fecha iniciono y terminono
        mstrWhere = mstrWhere & ",'',''"
    ElseIf .cckCriterios(6).Value = 1 And .cckCriterios(7).Value = 0 Then   '////////// fecha iniciosi y terminono
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "',''"
    ElseIf .cckCriterios(6).Value = 0 And .cckCriterios(7).Value = 1 Then  '////////// fecha iniciono y terminosi
        mstrWhere = ",'','" & pckFechaHasta.Value & "'"
    End If
    
    '// Estado
    If optTodas.Value = True Then
        mstrWhere = mstrWhere & ",'T'"
    ElseIf optVigente.Value = True Then
        mstrWhere = mstrWhere & ",'V'"
    ElseIf optLiquidada.Value = True Then
        mstrWhere = mstrWhere & ",'L'"
    ElseIf optCerrada.Value = True Then
        mstrWhere = mstrWhere & ",'F'"
    ElseIf optNula.Value = True Then
        mstrWhere = mstrWhere & ",'N'"
    End If
    
    
End With
'/////////////////////////////////////////////////////////////////////////////////
    
    '/// llama al procedimiento almacenado
    mstrSql = "Exec Tllr_ReservaRepuestos " & mstrWhere
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With AdoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              itmItem.SubItems(1) = ValorNulo(IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "V", "VIGENTE", IIf(!estado = "N", "NULA", IIf(!estado = "C", "CERRADA", IIf(!estado = "F", "FACTURADA", IIf(!estado = "B", "BOLETEADA", "OTRO"))))))) & "(" & ValorNulo(!Nro_Factura_Emitida) & ")"
              itmItem.SubItems(2) = Format(ValorNulo(!Fecha_Emision), "dd/mm/yyyy")
              itmItem.SubItems(3) = ValorNulo(!Patente)
              itmItem.SubItems(4) = ValorNulo(!Razon_Social)
              itmItem.SubItems(5) = ValorNulo(!Marca)
              itmItem.SubItems(6) = ValorNulo(!Modelo)
              itmItem.SubItems(7) = ValorNulo(IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA"))
              AdoTemp.MoveNext
          Wend
       End If
    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    mstrEstado = ""
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

If SW Then
    
    If Not Atributos("Glbl", "Tllr_30_0050", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If

    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    SW = False
End If

End Sub

Private Sub Form_Load()
SW = True
Me.cckCriterios(1).Caption = gstrNombrePatente
Me.lvDetalle.ColumnHeaders(4).Text = gstrNombrePatente
End Sub

Private Sub lvDetalle_Click()
If Me.lvDetalle.ListItems.Count > 0 Then
    FillRepuestosReservados gstrIdEmpresa, gstrIdSucursal, lvDetalle.SelectedItem, Mid(lvDetalle.SelectedItem.SubItems(7), 1, 1)
End If
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
'    gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social", gstrIdEmpresa)
'    'gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social")
'    txtCliente.Tag = gstrBusca
'    txtCliente = ClienteD(gstrBusca)
    'kjcv 02-02-2012
    gstrRutCliente = ""
    gstrNombreCliente = ""
    Libreria.ClienteBuscar Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
         If gstrRutCliente <> "" Then
            Me.txtCliente = gstrNombreCliente
            Me.txtCliente.Tag = gstrRutCliente
        End If
End If
End Sub

Private Sub tlbMarca_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Glbl_Marca", "Id_Marca", "Descripcion", "Busca Marca")
    txtMarca.Tag = gstrBusca
    txtMarca = MarcaD(gstrBusca)
End If
End Sub

Private Sub tlbModelo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistrosModelo(Conexion, "Glbl_Modelo", "Id_Modelo", "Id_Marca", "Descripcion", "Busca Modelo", IIf(Me.txtMarca.Tag <> "", txtMarca.Tag, "01"))
    txtModelo.Tag = gstrBusca
    txtModelo = ModeloD(txtMarca.Tag, gstrBusca)
End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub


Sub FillRepuestosReservados(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)
Dim pdblSumaValores As Double
lvDetalleRepuestosOT.ListItems.Clear

pdblSumaValores = 0

mstrSql = "Exec Tllr_CargaRepuestosReservados " & "'" & strIdEmpresa & "','" & strIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"

If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoTemp
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvDetalleRepuestosOT.ListItems.Add(, , ValorNulo(!Id_Item))
            Set lvDetalleRepuestosOT.SelectedItem = itmAux
            itmAux.SubItems(1) = ValorNulo(!Descripcion)
            itmAux.SubItems(2) = ValorNulo(!Familia)
            itmAux.SubItems(3) = FormatoValor(!Solicitado, "", 1)
            itmAux.SubItems(4) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
            pdblSumaValores = pdblSumaValores + !Precio_Unitario
            If !estado = "S" Then
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ForeColor = &HFF0000
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(1).ForeColor = &HFF0000
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(2).ForeColor = &HFF0000
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(3).ForeColor = &HFF0000
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(4).ForeColor = &HFF0000
            End If
            If !estado = "P" Then
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ForeColor = &HC0&
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
                Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
            End If
            .MoveNext
        Wend
        stbTotalCosto.Panels(2) = FormatoValor(pdblSumaValores, "", gintDecimalesMoneda)
    End If
    End With
End If
Conexion.CloseHost AdoTemp
End Sub
Sub FillRepuestosFaltantes(strIdEmpresa As String, strIdSucursal As String, strIdDocumento As String, strSeccion As String)

mstrSql = "SELECT Tllr_Repuestos_Faltantes.Id_Item, "
mstrSql = mstrSql & "Stck_Item.Descripcion, Tllr_Repuestos_Faltantes.Solicitado, "
mstrSql = mstrSql & "Tllr_Repuestos_Faltantes.Precio_Unitario, "
mstrSql = mstrSql & "Glbl_Familia.Descripcion AS Familia, "
mstrSql = mstrSql & "Tllr_Repuestos_Faltantes.Id_OT "
mstrSql = mstrSql & "FROM Tllr_Repuestos_Faltantes INNER JOIN "
mstrSql = mstrSql & "Stck_Item ON "
mstrSql = mstrSql & "Tllr_Repuestos_Faltantes.Id_Item = Stck_Item.Id_Item INNER "
mstrSql = mstrSql & "Join "
mstrSql = mstrSql & "Glbl_Familia ON "
mstrSql = mstrSql & "Stck_Item.Id_Familia = Glbl_Familia.Id_Familia "
mstrSql = mstrSql & " WHERE (Tllr_Repuestos_Faltantes.Id_Empresa = '" & strIdEmpresa & "') AND"
mstrSql = mstrSql & " (Tllr_Repuestos_Faltantes.Id_Sucursal = '" & strIdSucursal & "') AND"
mstrSql = mstrSql & " (Tllr_Repuestos_Faltantes.Id_OT = '" & strIdDocumento & "') AND"
mstrSql = mstrSql & " (Tllr_Repuestos_Faltantes.Seccion_OT = '" & strSeccion & "')"

If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoTemp
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = Me.lvDetalleRepuestosOT.ListItems.Add(, , ValorNulo(!Id_Item))
            Set Me.lvDetalleRepuestosOT.SelectedItem = itmAux
            itmAux.SubItems(1) = ValorNulo(!Descripcion)
            itmAux.SubItems(2) = ValorNulo(!Familia)
            itmAux.SubItems(3) = FormatoValor(!Solicitado, "", 1)
            itmAux.SubItems(4) = FormatoValor(!Precio_Unitario, "", gintDecimalesMoneda)
            
            Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ForeColor = &HC0&
            Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
            Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
            Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
            Me.lvDetalleRepuestosOT.ListItems(Me.lvDetalleRepuestosOT.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
                
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoTemp
End Sub

