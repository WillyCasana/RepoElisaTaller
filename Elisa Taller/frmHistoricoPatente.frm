VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmHistoricoPatente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico Por Placa"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmHistoricoPatente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptPatente 
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
      Appearance      =   0  'Flat
      Caption         =   "Imprimir Informe"
      Height          =   360
      Left            =   7995
      TabIndex        =   27
      Top             =   6255
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   60
      TabIndex        =   5
      Top             =   -15
      Width           =   11370
      Begin VB.CommandButton cmdResumenOT 
         Appearance      =   0  'Flat
         Caption         =   "Ver Resumen"
         Height          =   360
         Left            =   9720
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   525
         Left            =   3720
         TabIndex        =   28
         Top             =   1560
         Width           =   4680
         Begin VB.OptionButton optLiquidada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Liquidada"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1746
            TabIndex        =   33
            Top             =   240
            Width           =   990
         End
         Begin VB.OptionButton optNula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Nula"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3780
            TabIndex        =   32
            Top             =   270
            Width           =   675
         End
         Begin VB.OptionButton optCerrada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Emitidas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2739
            TabIndex        =   31
            Top             =   240
            Width           =   990
         End
         Begin VB.OptionButton optTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   30
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
            TabIndex        =   29
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
         TabIndex        =   26
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recepcionista"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   5640
         TabIndex        =   22
         Top             =   945
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5640
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1185
         Width           =   4635
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   6
         TabIndex        =   15
         Top             =   525
         Width           =   1020
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Placa"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtNroRecord 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10380
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "0"
         Top             =   525
         Width           =   555
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
         Width           =   4155
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
         Left            =   10560
         Top             =   960
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
               Picture         =   "frmHistoricoPatente.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHistoricoPatente.frx":3E66
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.UpDown updNroRecord 
         Height          =   315
         Left            =   10920
         TabIndex        =   16
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtNroRecord"
         BuddyDispid     =   196621
         OrigLeft        =   10950
         OrigTop         =   525
         OrigRight       =   11190
         OrigBottom      =   840
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   4920
         TabIndex        =   17
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
         Left            =   8400
         TabIndex        =   18
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
         Left            =   4920
         TabIndex        =   19
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
         Left            =   8880
         TabIndex        =   23
         Top             =   885
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
         TabIndex        =   24
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95682561
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1800
         TabIndex        =   25
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95682561
         CurrentDate     =   36776
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registros"
         Height          =   195
         Index           =   8
         Left            =   10410
         TabIndex        =   20
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   6240
      TabIndex        =   0
      Top             =   6255
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   1
      Top             =   6255
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   120
      TabIndex        =   4
      Top             =   2175
      Width           =   11355
      _ExtentX        =   20029
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
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha Emisión"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Seccion"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tipo"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Kilometros"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Trabajo Efectuado"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Patente"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   3
      Top             =   6390
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
      Top             =   6390
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmHistoricoPatente"
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
'    Dim rc As Long
'    Dim WinPath As String
'    WinPath = Space$(300)
'    rc = GetWindowsDirectory(WinPath, 300)
'    GcamBaseTem = Trim$(WinPath)
'    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
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
'    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
   
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,FechaIngreso Text,Recepcionista text,Seccion text,Tipo text,Kilometros text,Trabajo text,Valor Double,Patente Text,Cliente Text,Marca Text,Modelo Text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!FechaIngreso = IIf(lvDetalle.SelectedItem.SubItems(2) = "", "", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Recepcionista = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Seccion = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Tipo = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!Kilometros = IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6))
        Tabla!Trabajo = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla!Valor = IIf(lvDetalle.SelectedItem.SubItems(8) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(8), gstrMonedaLocal))
'        Tabla!Patente = IIf(txtPatente = "", " ", txtPatente)   ' IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
        Tabla!Cliente = txtCliente
        Tabla!Marca = txtMarca
        Tabla!Modelo = txtModelo
        
        Tabla.Update
    Next i
'   Tabla.Close
   
   Tabla.Close
   Dbsnueva.Close
   
   With rptPatente
        .ReportFileName = gstrPathReporte & "\HistoricoPatente.Rpt"
        .WindowTitle = "Historico Por " & gstrNombrePatente
        .WindowState = crptMaximized
'        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Historico Por " & gstrNombrePatente & "'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "desde='" & pckFechaDesde & "'"
        .Formulas(6) = "hasta='" & pckFechaHasta & "'"
        .Formulas(7) = "TDecimal=" & gintDecimalesMoneda
        .Formulas(8) = "TSigla='" & gstrMonedaLocal & "'"
        .Formulas(9) = "NombrePatente='" & gstrNombrePatente & "'"
        
        .Destination = crptToWindow
        .Action = True
   End With
   
'   Dbsnueva.Close
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
End Select
End Sub
Private Sub cmdBuscarOT_Click()
Dim mstrSql As String
Dim lstrSQL As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim mstrEstado As String
Dim ContLinea As Integer
Dim mstrNumeroDocumento As String
Dim VinCampaña As String

lvDetalle.ListItems.Clear
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
    
    If .cckCriterios(5).Value = 1 Then  '////////// recepcionista
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtRecepcionista, "Comienzo del Campo", apSqlServer) & "'"
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
    mstrSql = "Exec Tllr_HistoricoPatente " & mstrWhere
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          ContLinea = 0
          txtCliente = ValorNulo(!Cliente)
          txtMarca = ValorNulo(!Marca)
          txtModelo = ValorNulo(!Modelo)
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              If !est = "F" Or !est = "B" Then
                 mstrNumeroDocumento = ValorNulo(!Nro_Factura_Emitida)  'TraeNumeroDocumento(!Sec, !Id_OT, "")
              Else
                mstrNumeroDocumento = "S/N"
              End If
              itmItem.SubItems(1) = ValorNulo(IIf(!est = "L", "LIQUIDADA", IIf(!est = "V", "VIGENTE", IIf(!est = "N", "NULA", IIf(!est = "C", "CERRADA", IIf(!est = "F", "FACTURADA", IIf(!est = "B", "BOLETEADA", "OTRO"))))))) & "(" & mstrNumeroDocumento & ")"
              itmItem.SubItems(2) = Format(ValorNulo(!FEC), "dd/mm/yyyy")
              itmItem.SubItems(3) = ValorNulo(!RECEP)
              itmItem.SubItems(4) = ValorNulo(IIf(!Sec = "M", "MECANICA", "CARROCERIA"))
              itmItem.SubItems(5) = ValorNulo(!GAR)
              itmItem.SubItems(6) = ValorNulo(FormatoValor(!KMS, "", 0))
              itmItem.SubItems(9) = ValorNulo(!Pat)
              
              'campaña   //// rescata vin para verificar en la tabla campaña
              lstrSQL = "Select Vin from Tllr_Vehiculo_Cliente Where Patente='" & !Pat & "'"
              If Conexion.SendHost(lstrSQL, AdoAux, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                With AdoAux
                   If Not .BOF And Not .EOF Then
                        VinCampaña = !VIN
                   End If
                End With
              End If
              
              'pregunta por los servicios del vin
              lstrSQL = "Select Servicio from Tllr_Campañas where Vin='" & VinCampaña & "' And estado='T'"
              If Conexion.SendHost(lstrSQL, AdoAux, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                With AdoAux
                   If Not .BOF And Not .EOF Then
                        If lvDetalle.ListItems(lvDetalle.ListItems.Count).SubItems(7) <> "" Then
                          Set itmItem = lvDetalle.ListItems.Add(, , "")
                          ContLinea = 1
                        End If
                        itmItem.SubItems(7) = "Campaña : " & ValorNulo(!servicio)
                        itmItem.SubItems(8) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                        AdoAux.MoveNext
                        While Not .EOF
                            Set itmItem = lvDetalle.ListItems.Add(, , "")
                            itmItem.SubItems(7) = "Campaña : " & ValorNulo(!Descripcion)
                            itmItem.SubItems(8) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                            AdoAux.MoveNext
                        Wend
                   End If
                End With
              End If
              
              '/// Mecanica
              If gstrServiciosMarca = "S" Then
                lstrSQL = "Exec Tllr_CargaServicios_Mecanica_MM " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & !Sec & "','" & !Id_OT & "'"
              Else
                lstrSQL = "Exec Tllr_CargaServicios_Mecanica " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & !Sec & "','" & !Id_OT & "'"
              End If
              
              If Conexion.SendHost(lstrSQL, AdoAux, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                With AdoAux
                   If Not .BOF And Not .EOF Then
                      If lvDetalle.ListItems(lvDetalle.ListItems.Count).SubItems(7) <> "" Then
                        Set itmItem = lvDetalle.ListItems.Add(, , "")
                        ContLinea = 1
                      End If
                      itmItem.SubItems(7) = ValorNulo(!Descripcion)
                      itmItem.SubItems(8) = FormatoValor(!Total, gstrMonedaLocal, gintDecimalesMoneda)
                      AdoAux.MoveNext
                      While Not .EOF
                          Set itmItem = lvDetalle.ListItems.Add(, , "")
                          itmItem.SubItems(7) = ValorNulo(!Descripcion)
                          itmItem.SubItems(8) = FormatoValor(!Total, gstrMonedaLocal, gintDecimalesMoneda)
                          AdoAux.MoveNext
                      Wend
                   End If
                End With
              End If
              
              '/// Carroceria
              lstrSQL = "Exec Tllr_CargaServicios_Carroceria " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & !Sec & "','" & !Id_OT & "'"
              
              If Conexion.SendHost(lstrSQL, AdoAux, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                With AdoAux
                   If Not .BOF And Not .EOF Then
                      If lvDetalle.ListItems(lvDetalle.ListItems.Count).SubItems(7) <> "" Then
                        Set itmItem = lvDetalle.ListItems.Add(, , "")
                        ContLinea = 1
                      End If
                      itmItem.SubItems(7) = ValorNulo(!DescCarr)
                      itmItem.SubItems(8) = FormatoValor(!SubTotal, gstrMonedaLocal, gintDecimalesMoneda)
                      AdoAux.MoveNext
                      While Not .EOF
                          Set itmItem = lvDetalle.ListItems.Add(, , "")
                          itmItem.SubItems(7) = ValorNulo(!DescCarr)
                          itmItem.SubItems(8) = FormatoValor(!SubTotal, gstrMonedaLocal, gintDecimalesMoneda)
                          AdoAux.MoveNext
                      Wend
                   End If
                End With
              End If
              
              '/// Otros
              lstrSQL = "Exec Tllr_CargaServicios_Otro " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & !Sec & "','" & !Id_OT & "'"
              
              If Conexion.SendHost(lstrSQL, AdoAux, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                With AdoAux
                   If Not .BOF And Not .EOF Then
                      If lvDetalle.ListItems(lvDetalle.ListItems.Count).SubItems(7) <> "" Then
                        Set itmItem = lvDetalle.ListItems.Add(, , "")
                        ContLinea = 1
                      End If
                      itmItem.SubItems(7) = ValorNulo(!Des)
                      itmItem.SubItems(8) = FormatoValor(!SubTotal, gstrMonedaLocal, gintDecimalesMoneda)
                      AdoAux.MoveNext
                      While Not .EOF
                          Set itmItem = lvDetalle.ListItems.Add(, , "")
                          itmItem.SubItems(7) = ValorNulo(!Des)
                          itmItem.SubItems(8) = FormatoValor(!SubTotal, gstrMonedaLocal, gintDecimalesMoneda)
                          AdoAux.MoveNext
                      Wend
                   End If
                End With
              End If
              
              '/// Terceros
              lstrSQL = "Exec Tllr_CargaServicios_Terceros " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & !Sec & "','" & !Id_OT & "'"
              
              If Conexion.SendHost(lstrSQL, AdoAux, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                With AdoAux
                   If Not .BOF And Not .EOF Then
                      If lvDetalle.ListItems(lvDetalle.ListItems.Count).SubItems(7) <> "" Then
                        Set itmItem = lvDetalle.ListItems.Add(, , "")
                        ContLinea = 1
                      End If
                      itmItem.SubItems(7) = ValorNulo(!servicio)
                      itmItem.SubItems(8) = FormatoValor(!STotal, gstrMonedaLocal, gintDecimalesMoneda)
                      AdoAux.MoveNext
                      While Not .EOF
                          Set itmItem = lvDetalle.ListItems.Add(, , "")
                          itmItem.SubItems(7) = ValorNulo(!servicio)
                          itmItem.SubItems(8) = FormatoValor(!STotal, gstrMonedaLocal, gintDecimalesMoneda)
                          AdoAux.MoveNext
                      Wend
                   End If
                End With
              End If
              
              '/// Repuestos
              lstrSQL = "Exec Tllr_CargaServicios_Repuestos " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & !Sec & "','" & !Id_OT & "'"
              
              If Conexion.SendHost(lstrSQL, AdoAux, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                With AdoAux
                   If Not .BOF And Not .EOF Then
                      If lvDetalle.ListItems(lvDetalle.ListItems.Count).SubItems(7) <> "" Then
                        Set itmItem = lvDetalle.ListItems.Add(, , "")
                        ContLinea = 1
                      End If
                      itmItem.SubItems(7) = ValorNulo(!item) & " (" & !CANTY & ")"
                      itmItem.SubItems(8) = FormatoValor(!SubTotal, gstrMonedaLocal, gintDecimalesMoneda)
                      AdoAux.MoveNext
                      While Not .EOF
                          Set itmItem = lvDetalle.ListItems.Add(, , "")
                          itmItem.SubItems(7) = ValorNulo(!item) & " (" & !CANTY & ")"
                          itmItem.SubItems(8) = FormatoValor(!SubTotal, gstrMonedaLocal, gintDecimalesMoneda)
                          AdoAux.MoveNext
                      Wend
                   End If
                End With
              End If
              
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
    ImprimirConsulta
Else
    MsgBox "no"
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
    gstrSeccion = lvDetalle.SelectedItem.SubItems(11)
End If
Unload Me
End Sub




Private Sub Form_Activate()

If SW Then

    If Not Atributos("Glbl", "Tllr_30_0130", True, True, True, True) Then
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
Me.lvDetalle.ColumnHeaders(10).Text = gstrNombrePatente
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social", gstrIdEmpresa)
    'gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social")
    txtCliente.Tag = gstrBusca
    txtCliente = ClienteD(gstrBusca)
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

Private Sub txtNroRecord_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtNroRecord, strComa)
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
