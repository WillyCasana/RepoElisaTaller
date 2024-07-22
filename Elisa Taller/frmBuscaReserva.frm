VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBuscaReserva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Reservas"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   Icon            =   "frmBuscaReserva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   13620
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptOT 
      Left            =   3945
      Top             =   6210
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
      Caption         =   "Excel"
      Height          =   360
      Left            =   6195
      TabIndex        =   29
      Top             =   6240
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   13410
      Begin VB.TextBox txtVin 
         Height          =   315
         Left            =   5040
         TabIndex        =   39
         Top             =   525
         Width           =   1695
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Vin"
         Height          =   195
         Index           =   10
         Left            =   5040
         TabIndex        =   38
         Top             =   330
         Width           =   855
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Reserva (Fin)"
         Height          =   195
         Index           =   9
         Left            =   5445
         TabIndex        =   36
         Top             =   1560
         Width           =   1680
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   525
         Left            =   7290
         TabIndex        =   30
         Top             =   1575
         Width           =   5445
         Begin VB.OptionButton optLiquidada 
            Caption         =   "Confirmada"
            Height          =   195
            Left            =   1980
            TabIndex        =   35
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optNula 
            Caption         =   "Nula"
            Height          =   195
            Left            =   4530
            TabIndex        =   34
            Top             =   225
            Width           =   675
         End
         Begin VB.OptionButton optCerrada 
            Caption         =   "Cancelada"
            Height          =   195
            Left            =   3285
            TabIndex        =   33
            Top             =   225
            Width           =   1110
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   75
            TabIndex        =   32
            Top             =   240
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton optVigente 
            Caption         =   "Vigente"
            Height          =   195
            Left            =   1005
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Emisión (Fin)"
         Height          =   195
         Index           =   7
         Left            =   1770
         TabIndex        =   27
         Top             =   1545
         Width           =   1365
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Reserva (Ini)"
         Height          =   195
         Index           =   8
         Left            =   3840
         TabIndex        =   26
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Recepcionista"
         Height          =   195
         Index           =   5
         Left            =   7365
         TabIndex        =   22
         Top             =   930
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7365
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1170
         Width           =   3675
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Nro Reserva"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox txtNroOt 
         Height          =   300
         Left            =   105
         MaxLength       =   15
         TabIndex        =   19
         Top             =   525
         Width           =   2310
      End
      Begin VB.TextBox txtPatente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   15
         Top             =   525
         Width           =   1020
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Placa"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   14
         Top             =   330
         Width           =   855
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Marca "
         Height          =   195
         Index           =   2
         Left            =   7320
         TabIndex        =   13
         Top             =   315
         Width           =   870
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Modelo"
         Height          =   195
         Index           =   3
         Left            =   10320
         TabIndex        =   12
         Top             =   345
         Width           =   840
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1185
         Width           =   4455
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         MaxLength       =   50
         TabIndex        =   9
         Top             =   540
         Width           =   2835
      End
      Begin VB.TextBox txtModelo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10320
         MaxLength       =   50
         TabIndex        =   8
         Top             =   555
         Width           =   2835
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Emisión (Ini)"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1545
         Width           =   1320
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
            NumListImages   =   22
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":038A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":049C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":08F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":0D4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":11A4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":12B6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":13C8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":14DA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":15EC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":16FE
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1810
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1922
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1A34
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1B46
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1C58
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1D6A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1E7C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":1F8E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":20A0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":21B2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":2604
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaReserva.frx":2A56
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   9600
         TabIndex        =   16
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
         Left            =   12600
         TabIndex        =   17
         Top             =   240
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
      Begin MSComctlLib.Toolbar tlbCliente 
         Height          =   330
         Left            =   4140
         TabIndex        =   18
         Top             =   855
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
      Begin MSComctlLib.Toolbar tlbRecep 
         Height          =   330
         Left            =   10590
         TabIndex        =   23
         Top             =   870
         Width           =   375
         _ExtentX        =   661
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
         TabIndex        =   24
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   87097345
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1770
         TabIndex        =   25
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   87097345
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckLiquidaIni 
         Height          =   315
         Left            =   3840
         TabIndex        =   28
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   87097345
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckLiquidaFin 
         Height          =   315
         Left            =   5445
         TabIndex        =   37
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   87097345
         CurrentDate     =   36776
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   4440
      TabIndex        =   0
      Top             =   6255
      Width           =   1680
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   360
      Left            =   7980
      TabIndex        =   1
      Top             =   6255
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   2
      Top             =   6255
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   60
      TabIndex        =   5
      Top             =   2160
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   6932
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° Reserva"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Placa"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Modelo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fecha Emisión"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Fecha Reserva"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Hora Reserva"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Recepcionista"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Observación"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Taxi Destino"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   2640
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   4
      Top             =   6390
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   6390
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuscaReserva"
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
Dim lintVigentes As Integer
Dim lintNulas As Integer
Dim lintConfirmadas As Integer
Dim lintRecepcionadas As Integer
Dim lintCanceladas As Integer
Dim lintTotalReservas As Integer

lintVigentes = 0
lintNulas = 0
lintConfirmadas = 0
lintRecepcionadas = 0
lintCanceladas = 0
lintTotalReservas = 0

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
'    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
'    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroReserva text,Estado text,Patente text,Cliente text,Marca text,Modelo text,FechaIngreso date,FechaReserva date,HoraReserva text,Recepcionista text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroReserva = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Cliente = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!FechaIngreso = DateValue(IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6)))
        Tabla!FechaReserva = DateValue(IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7)))
        Tabla!HoraReserva = IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
        Tabla!Recepcionista = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
        
        Tabla.Update
        
        'acumula los estados
        If Me.lvDetalle.SelectedItem.SubItems(1) = "VIGENTE" Then
            lintVigentes = lintVigentes + 1
        ElseIf Me.lvDetalle.SelectedItem.SubItems(1) = "NULA" Then
            lintNulas = lintNulas + 1
        ElseIf Me.lvDetalle.SelectedItem.SubItems(1) = "CONFIRMADA" Then
            lintConfirmadas = lintConfirmadas + 1
        ElseIf Me.lvDetalle.SelectedItem.SubItems(1) = "RECEPCIONADA" Then
            lintRecepcionadas = lintRecepcionadas + 1
        ElseIf Me.lvDetalle.SelectedItem.SubItems(1) = "CANCELADA" Then
            lintCanceladas = lintCanceladas + 1
        End If

        
    Next i
   Tabla.Close
   Dbsnueva.Close
   
   With rptOT
        .ReportFileName = gstrPathReporte & "\ListadoReservas.rpt"
        .WindowTitle = "Reporte de Reservas de Hora"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
'        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='LISTADO DE RESERVAS DE HORAS PARA SERVICIO'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "NombrePatente='" & gstrNombrePatente & "'"
        .Formulas(7) = "Vigentes='" & FormatoValor((lintVigentes * 100) / Me.lvDetalle.ListItems.Count, "%", 2) & "'"
        .Formulas(8) = "Nulas='" & FormatoValor((lintNulas * 100) / Me.lvDetalle.ListItems.Count, "%", 2) & "'"
        .Formulas(9) = "Confirmadas='" & FormatoValor((lintConfirmadas * 100) / Me.lvDetalle.ListItems.Count, "%", 2) & "'"
        .Formulas(10) = "Recepcionadas='" & FormatoValor((lintRecepcionadas * 100) / Me.lvDetalle.ListItems.Count, "%", 2) & "'"
        .Formulas(11) = "Canceladas='" & FormatoValor((lintCanceladas * 100) / Me.lvDetalle.ListItems.Count, "%", 2) & "'"

        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = True
   End With
   
   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub


Private Sub cckCriterios_Click(Index As Integer)
Select Case Index
Case 0
    If cckCriterios(Index).Value = 0 Then
        txtNroOt.Enabled = False
        txtNroOt = ""
    Else
        txtNroOt.Enabled = True
        txtNroOt.SetFocus
    End If
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
        tlbRecep.Enabled = False
        txtRecepcionista.Enabled = False
        txtRecepcionista = ""
    Else
        tlbRecep.Enabled = True
        txtRecepcionista.Enabled = True
        txtRecepcionista.SetFocus
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
Case 8
    If cckCriterios(Index).Value = 0 Then
        pckLiquidaIni.Enabled = False
    Else
        pckLiquidaIni.Enabled = True
        pckLiquidaIni.SetFocus
    End If
Case 9
    If cckCriterios(Index).Value = 0 Then
        pckLiquidaFin.Enabled = False
    Else
        pckLiquidaFin.Enabled = True
        pckLiquidaFin.SetFocus
    End If
Case 10
    If cckCriterios(Index).Value = 0 Then
        txtVin.Enabled = False
        txtVin = ""
    Else
        txtVin.Enabled = True
        txtVin.SetFocus
    End If
End Select
End Sub


Private Sub cmdBuscarOT_Click()
Dim mstrSQL As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim mstrEstado As String

    lvDetalle.ListItems.Clear
mstrWhere = ""
With Me
    If .cckCriterios(0).Value = 1 Then  '////////// nro ot
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_ReservaHora.Id_Reserva LIKE '" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Tllr_ReservaHora.Id_Reserva LIKE '" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(1).Value = 1 Then  '////////// patente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_ReservaHora.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Tllr_ReservaHora.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(10).Value = 1 Then  '////////// vin
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Vehiculo_Cliente.Vin LIKE '" & MatchMode(.txtVin, "Cualquier Parte del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Tllr_Vehiculo_Cliente.Vin LIKE '" & MatchMode(.txtVin, "Cualquier Parte del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(2).Value = 1 Then  '////////// marca
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(3).Value = 1 Then  '////////// modelo
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(4).Value = 1 Then  '////////// cliente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(5).Value = 1 Then  '////////// recepcionista
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Mecanicos.Nombre LIKE '" & MatchMode(.txtRecepcionista, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Tllr_Mecanicos.Nombre LIKE '" & MatchMode(.txtRecepcionista, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(6).Value = 1 Then  '////////// fecha inicio
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:59" & "'"
            Else
                mstrWhere = " WHERE fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:59" & "'"
            End If
        Else
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision = '" & pckFechaDesde.Value & "' "
            Else
                mstrWhere = " WHERE fecha_emision = '" & pckFechaDesde.Value & "' "
            End If
        End If
    Else
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = " AND fecha_emision = '" & pckFechaHasta.Value & "' "
            Else
                mstrWhere = " WHERE fecha_emision = '" & pckFechaHasta.Value & "' "
            End If
        End If
    End If
    
    '//////////////////////////////////////////////////////
    If .cckCriterios(8).Value = 1 Then  '////////// fecha liquidacion inicio
        If .cckCriterios(9).Value = 1 Then  '////////// fecha liquidacion termino
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_Reserva between '" & pckLiquidaIni.Value & "' and '" & pckLiquidaFin.Value & "'"
            Else
                mstrWhere = " WHERE fecha_Reserva between '" & pckLiquidaIni.Value & "' and '" & pckLiquidaFin.Value & "'"
            End If
        Else
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_Reserva = '" & pckLiquidaIni.Value & "' "
            Else
                mstrWhere = " WHERE fecha_Reserva = '" & pckLiquidaIni.Value & "' "
            End If
        End If
    Else
        If .cckCriterios(9).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = " AND fecha_Reserva = '" & pckLiquidaFin.Value & "'"
            Else
                mstrWhere = " WHERE fecha_Reserva = '" & pckLiquidaFin.Value & "'"
            End If
        End If
    End If
     '////////// empresa y sucursal
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " AND Tllr_ReservaHora.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_ReservaHora.ID_SUCURSAL='" & gstrIdSucursal & "' "
        Else
            mstrWhere = " WHERE Tllr_ReservaHora.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_ReservaHora.ID_SUCURSAL='" & gstrIdSucursal & "' "
        End If
    '//////////////////estado
        If optTodas.Value = True Then
            mstrEstado = "IN ('V','C','E','N','R')"
        ElseIf optVigente.Value = True Then
            mstrEstado = "IN ('V')"
        ElseIf optLiquidada.Value = True Then
            mstrEstado = "IN ('C')"
        ElseIf optCerrada.Value = True Then
            mstrEstado = "IN ('E')"
        ElseIf optNula.Value = True Then
            mstrEstado = "IN ('N')"
        End If
            
        If mstrEstado <> "" Then
            mstrWhere = mstrWhere & " And Tllr_ReservaHora.Estado  " & mstrEstado
        End If
End With
'/////////////////////////////////////////////////////////////////////////////////
    mstrSQL = "SELECT Tllr_ReservaHora.Id_Reserva, "
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Patente AS PAT,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Marca AS IDMAR,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Vin,"
    mstrSQL = mstrSQL & " Glbl_Marca.Descripcion AS MARCA,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Modelo AS IDMOD,"
    mstrSQL = mstrSQL & " Glbl_Modelo.Descripcion AS MODELO,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor AS IDCLI,"
    mstrSQL = mstrSQL & " case Tllr_ReservaHora.Patente when '' then Tllr_ReservaHora.Nombre else  Glbl_Cliente_Proveedor.Razon_Social end as CLIENTE,"
    mstrSQL = mstrSQL & " case Tllr_ReservaHora.Patente when '' then Tllr_ReservaHora.Telefono else  Glbl_Cliente_Proveedor.Telefono end as FONO,"
'    mstrSQL = mstrSQL & " Glbl_Cliente_Proveedor.Razon_Social AS CLIENTE,"
'    mstrSQL = mstrSQL & " Glbl_Cliente_Proveedor.Telefono AS FONO, "
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Fecha_Emision AS FEC, "
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Fecha_Reserva AS FECRES, "
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Hora_Reserva AS HORARES, "
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Estado AS EST, "
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Recepcionista AS IDREC,"
    mstrSQL = mstrSQL & " Tllr_Mecanicos.Nombre AS RECEP, "
    
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Total_Mecanica AS TMEC,"
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Total_Otros AS TOTR,"
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Total_Repuestos AS TREP,"
    mstrSQL = mstrSQL & " Tllr_ReservaHora.Total_Reserva AS TNETO "
    'kjcv 11.09.14
    mstrSQL = mstrSQL & " ,Tllr_ReservaHora.Taxi_destino AS TDESTINO "
    
    'kjcv 01.08.14 se agrega el campo de Observaciones
    mstrSQL = mstrSQL & " ,Tllr_ReservaHora.Reparacion AS REP "
    
    mstrSQL = mstrSQL & " FROM Tllr_ReservaHora LEFT OUTER JOIN Tllr_Mecanicos "
    mstrSQL = mstrSQL & " ON Tllr_ReservaHora.Recepcionista = Tllr_Mecanicos.Id_Mecanico  "
    mstrSQL = mstrSQL & " AND Tllr_ReservaHora.Id_Empresa = Tllr_Mecanicos.Id_Empresa  "
    mstrSQL = mstrSQL & " AND Tllr_ReservaHora.Id_Sucursal = Tllr_Mecanicos.Id_Sucursal  "
    mstrSQL = mstrSQL & "LEFT OUTER Join Glbl_Modelo LEFT OUTER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca RIGHT OUTER JOIN Tllr_Vehiculo_Cliente ON Glbl_Modelo.Id_Modelo = Tllr_Vehiculo_Cliente.Id_Modelo AND Glbl_Modelo.Id_Marca = Tllr_Vehiculo_Cliente.Id_Marca LEFT OUTER Join Glbl_Cliente_Proveedor ON Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor ON Tllr_ReservaHora.Patente = Tllr_Vehiculo_Cliente.Patente "
    mstrSQL = mstrSQL & mstrWhere
    mstrSQL = mstrSQL & "  ORDER BY ID_Reserva"
    
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_Reserva)
              itmItem.SubItems(1) = ValorNulo(IIf(!est = "C", "CONFIRMADA", IIf(!est = "V", "VIGENTE", IIf(!est = "N", "NULA", IIf(!est = "E", "CANCELADA", IIf(!est = "F", "FACTURADA", IIf(!est = "B", "BOLETEADA", IIf(!est = "R", "RECEPCIONADA", "OTRO"))))))))
              itmItem.SubItems(2) = ValorNulo(!Pat)
              itmItem.SubItems(3) = ValorNulo(!Cliente)
              itmItem.SubItems(4) = ValorNulo(!FONO)
              itmItem.SubItems(5) = ValorNulo(!Modelo)
              itmItem.SubItems(6) = Format(ValorNulo(!FEC), "dd/mm/yyyy")
              itmItem.SubItems(7) = Format(ValorNulo(!FECRES), "dd/mm/yyyy")
              itmItem.SubItems(8) = ValorNulo(!HORARES)
              itmItem.SubItems(9) = ValorNulo(!RECEP)
              'kjcv 01.08.14
              itmItem.SubItems(10) = ValorNulo(!REP)
              'kjcv 11.09.14
              itmItem.SubItems(11) = ValorNulo(!TDESTINO)
              adoTemp.MoveNext
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
'    ImprimirConsulta
    ExportarDatos Me.lvDetalle, Me.cdExportar, Me.hwnd
Else
    MsgBox "No existen datos en la lista"
End If
End Sub

Private Sub cmdResumenOT_Click()
If Not lvDetalle.SelectedItem Is Nothing Then
With frmResumenOT
    .lblIdOT = lvDetalle.SelectedItem
    .lblSeccion = lvDetalle.SelectedItem.SubItems(9)
    .lblestado = lvDetalle.SelectedItem.SubItems(1)
    .lblPatente = lvDetalle.SelectedItem.SubItems(2)
    .lblCliente = lvDetalle.SelectedItem.SubItems(3)
    .lblMarca = lvDetalle.SelectedItem.SubItems(4)
    .lblModelo = lvDetalle.SelectedItem.SubItems(5)
    .lblTotalMec = FormatoValor(lvDetalle.SelectedItem.SubItems(12), "", gintDecimalesMoneda)
    .lblTotalCar = FormatoValor(lvDetalle.SelectedItem.SubItems(13), "", gintDecimalesMoneda)
    .lblTotalOtr = FormatoValor(lvDetalle.SelectedItem.SubItems(14), "", gintDecimalesMoneda)
    .lblTotalTer = FormatoValor(lvDetalle.SelectedItem.SubItems(15), "", gintDecimalesMoneda)
    .lblTotalRep = FormatoValor(lvDetalle.SelectedItem.SubItems(16), "", gintDecimalesMoneda)
    .lblTotalMat = FormatoValor(lvDetalle.SelectedItem.SubItems(17), "", gintDecimalesMoneda)
    .lblTotalIns = FormatoValor(lvDetalle.SelectedItem.SubItems(18), "", gintDecimalesMoneda)
    .lblsubtotal = FormatoValor(lvDetalle.SelectedItem.SubItems(19), "", gintDecimalesMoneda)
    .lblIva = FormatoValor(lvDetalle.SelectedItem.SubItems(20), "", gintDecimalesMoneda)
    .lblTotalOT = FormatoValor(lvDetalle.SelectedItem.SubItems(21), "", gintDecimalesMoneda)
    .ReCalculo
    .Show vbModal
End With
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
If Not lvDetalle.SelectedItem Is Nothing Then
    gstrBusca = lvDetalle.SelectedItem
End If
Unload Me
End Sub




Private Sub Form_Activate()
Me.cckCriterios(1).Caption = gstrNombrePatente
If SW Then
    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    pckLiquidaIni = Date
    pckLiquidaFin = EOM(Date)
    'cmdImprimir.Enabled = Atributos("Glbl", "Tllr_30_0010", True, True, True, True)
    SW = False
End If

End Sub

Private Sub Form_Load()
SW = True
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub lvDetalle_DblClick()
If cmdSeleccionar.Enabled = True Then cmdSeleccionar.Value = True
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
'    gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social", gstrIdEmpresa)
'    'gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social")
'    txtCliente.Tag = gstrBusca
'    txtCliente = ClienteD(gstrBusca)
    'kjcv 13.11.13
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

Private Sub tlbRecep_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Busca Mecanico")
    txtRecepcionista.Tag = gstrBusca
    txtRecepcionista = MecanicoD(gstrBusca)
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

Private Sub txtRecepcionista_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub
