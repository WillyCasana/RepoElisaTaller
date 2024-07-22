VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmInfServiceRateReserva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Rate Reserva de Repuestos"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmInfServiceRateReserva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Totales del Período"
      Height          =   1575
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   11295
      Begin VB.Label Label9 
         Caption         =   "Diferencia"
         Height          =   255
         Left            =   5640
         TabIndex        =   38
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblDiferencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         TabIndex        =   37
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblTotalMontoEntregado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   36
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblTotalCantidadEntregado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1200
         TabIndex        =   35
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Entregado"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Solicitado"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblTotalMontoSolicitado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblTotalCantidadSolicitado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1200
         TabIndex        =   31
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Monto"
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
   End
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
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
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
               Picture         =   "frmInfServiceRateReserva.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInfServiceRateReserva.frx":3E66
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
      Height          =   3210
      Left            =   75
      TabIndex        =   4
      Top             =   2160
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   5662
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
      NumItems        =   16
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
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Familia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Piezas Solicitadas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Piezas Despachadas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "% Piezas Despachadas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Precio Unitario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Monto Solicitado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Monto Despachado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Diferencia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "% Diferencia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Patente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
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
Attribute VB_Name = "frmInfServiceRateReserva"
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
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,FechaIngreso Text,Item Text,Descripcion Text,Familia Text,Solicitado Text,Despachado Text,PiezasDesp Text,PrecioUnitario Text,MontoSol Text,MontoDesp Text,Diferencia Text,Patente Text,Cliente Text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!FechaIngreso = IIf(lvDetalle.SelectedItem.SubItems(2) = "", "", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Item = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Descripcion = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Familia = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!Solicitado = IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6))
        Tabla!Despachado = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla!PiezasDesp = IIf(lvDetalle.SelectedItem.SubItems(8) = "", 0, lvDetalle.SelectedItem.SubItems(8))
        Tabla!PrecioUnitario = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
        Tabla!MontoSol = IIf(lvDetalle.SelectedItem.SubItems(10) = "", " ", lvDetalle.SelectedItem.SubItems(10))
        Tabla!MontoDesp = IIf(lvDetalle.SelectedItem.SubItems(11) = "", " ", lvDetalle.SelectedItem.SubItems(11))
        Tabla!Diferencia = IIf(lvDetalle.SelectedItem.SubItems(12) = "", " ", lvDetalle.SelectedItem.SubItems(12))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(14) = "", " ", lvDetalle.SelectedItem.SubItems(14))
        Tabla!Cliente = IIf(lvDetalle.SelectedItem.SubItems(15) = "", " ", lvDetalle.SelectedItem.SubItems(15))
        Tabla.Update
    Next i
   Tabla.Close
   
   With rptPatente
        .ReportFileName = gstrPathReporte & "\ServiceRate.Rpt"
        .WindowTitle = "Service Rate Reserva de Repuestos"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Service Rate Reserva de Repuestos'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "desde='" & pckFechaDesde & "'"
        .Formulas(6) = "hasta='" & pckFechaHasta & "'"
        .Formulas(7) = "SumaCantSol='" & Me.lblTotalCantidadSolicitado & "'"
        .Formulas(8) = "SumaCantDesp='" & Me.lblTotalCantidadEntregado & "'"
        .Formulas(9) = "SumaMontoSol='" & Me.lblTotalMontoSolicitado & "'"
        .Formulas(10) = "SumaMontoDesp='" & Me.lblTotalMontoEntregado & "'"
        .Formulas(11) = "Diferencia='" & Me.lblDiferencia & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = True
   End With
   
   Dbsnueva.Close
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
Dim mstrNumeroDocumento As String
Dim SumaCantidadSolicitado As Double
Dim SumaCantidadDespachado As Double
Dim SumaMontoSolicitado As Double
Dim SumaMontoDespachado As Double
Dim Diferencia As Double

lvDetalle.ListItems.Clear
mstrWhere = "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "'"
SumaCantidadSolicitado = 0
SumaCantidadDespachado = 0
SumaMontoSolicitado = 0
SumaMontoDespachado = 0
Diferencia = 0
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
        mstrWhere = mstrWhere & ",'','" & pckFechaHasta.Value & "'"
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
    mstrSql = "Exec Tllr_ServiceRate_ReservaRepuestos " & mstrWhere
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With AdoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              itmItem.SubItems(1) = ValorNulo(IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "V", "VIGENTE", IIf(!estado = "N", "NULA", IIf(!estado = "C", "CERRADA", IIf(!estado = "F", "FACTURADA", IIf(!estado = "B", "BOLETEADA", "OTRO")))))))
              itmItem.SubItems(2) = Format(ValorNulo(!Fecha_Emision), "dd/mm/yyyy")
              itmItem.SubItems(3) = ValorNulo(!Item)
              itmItem.SubItems(4) = ValorNulo(!Descripcion)
              itmItem.SubItems(5) = ValorNulo(!Familia)
              itmItem.SubItems(6) = FormatoValor(ValorNulo(!Solicitado), "", 2)
              itmItem.SubItems(7) = FormatoValor(ValorNulo(!Despachado), "", 2)
              itmItem.SubItems(8) = FormatoValor((!Despachado * 100) / !Solicitado, "%", 2)
              itmItem.SubItems(9) = FormatoValor(ValorNulo(!Precio_Unitario), gstrMonedaLocal, gintDecimalesMoneda)
              itmItem.SubItems(10) = FormatoValor(!Solicitado * !Precio_Unitario, gstrMonedaLocal, gintDecimalesMoneda)
              itmItem.SubItems(11) = FormatoValor(!Despachado * !Precio_Unitario, gstrMonedaLocal, gintDecimalesMoneda)
              itmItem.SubItems(12) = FormatoValor((!Solicitado * !Precio_Unitario) - (!Despachado * !Precio_Unitario), gstrMonedaLocal, gintDecimalesMoneda)
              If !Precio_Unitario = 0 Then
                itmItem.SubItems(13) = "0"
              Else
                itmItem.SubItems(13) = ((!Despachado * !Precio_Unitario) * 100) / (!Solicitado * !Precio_Unitario)
              End If
              itmItem.SubItems(14) = ValorNulo(!Patente)
              itmItem.SubItems(15) = ValorNulo(!Razon_Social)
              
              SumaCantidadSolicitado = SumaCantidadSolicitado + !Solicitado
              SumaCantidadDespachado = SumaCantidadDespachado + !Despachado
              SumaMontoSolicitado = SumaMontoSolicitado + (!Solicitado * !Precio_Unitario)
              SumaMontoDespachado = SumaMontoDespachado + (!Despachado * !Precio_Unitario)
              Diferencia = SumaMontoSolicitado - SumaMontoDespachado
              AdoTemp.MoveNext
          Wend
          lblDiferencia = FormatoValor(Diferencia, gstrMonedaLocal, gintDecimalesMoneda)
          lblTotalCantidadSolicitado = FormatoValor(SumaCantidadSolicitado, "", gintDecimalesMoneda)
          lblTotalCantidadEntregado = FormatoValor(SumaCantidadDespachado, "", gintDecimalesMoneda)
          lblTotalMontoSolicitado = FormatoValor(SumaMontoSolicitado, gstrMonedaLocal, gintDecimalesMoneda)
          lblTotalMontoEntregado = FormatoValor(SumaMontoDespachado, gstrMonedaLocal, gintDecimalesMoneda)
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
    
    If Not Atributos("Glbl", "Tllr_30_0060", True, True, True, True) Then
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
Me.lvDetalle.ColumnHeaders(15).Text = gstrNombrePatente
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
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
