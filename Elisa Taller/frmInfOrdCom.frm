VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmInfOrdCom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ordenes de Compra"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmInfOrdCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSeleccionar 
      Appearance      =   0  'Flat
      Caption         =   "&Seleccionar"
      Height          =   360
      Left            =   6120
      TabIndex        =   16
      Top             =   4515
      Width           =   1215
   End
   Begin Crystal.CrystalReport rptOrdCom 
      Left            =   75
      Top             =   4530
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
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8700
      Begin MSComctlLib.Toolbar tlbProv 
         Height          =   330
         Left            =   4080
         TabIndex        =   15
         Top             =   1005
         Width           =   540
         _ExtentX        =   953
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
      Begin VB.CheckBox cckCriterio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Proveedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   14
         Top             =   840
         Width           =   1140
      End
      Begin VB.TextBox txtProv 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   13
         Top             =   1050
         Width           =   3915
      End
      Begin VB.CheckBox cckCriterio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Fin)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   12
         Top             =   210
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.CheckBox cckCriterio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Ini)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2415
         TabIndex        =   10
         Top             =   210
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin MSComCtl2.DTPicker pckIni 
         Height          =   330
         Left            =   2415
         TabIndex        =   9
         Top             =   405
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   87687169
         CurrentDate     =   36824
      End
      Begin VB.CheckBox cckCriterio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Nro. (Fin)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1245
         TabIndex        =   8
         Top             =   210
         Width           =   945
      End
      Begin VB.CheckBox cckCriterio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Nro. (Ini)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   210
         Width           =   945
      End
      Begin VB.TextBox txtHasta 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         TabIndex        =   6
         Top             =   405
         Width           =   945
      End
      Begin VB.TextBox txtDesde 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Top             =   405
         Width           =   945
      End
      Begin MSComCtl2.DTPicker pckFin 
         Height          =   330
         Left            =   3960
         TabIndex        =   11
         Top             =   405
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   87687169
         CurrentDate     =   36824
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Appearance      =   0  'Flat
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   7455
      TabIndex        =   4
      Top             =   4515
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   4800
      TabIndex        =   3
      Top             =   4515
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   3480
      TabIndex        =   2
      Top             =   4515
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwOrdCom 
      Height          =   2970
      Left            =   15
      TabIndex        =   1
      Top             =   1500
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   5239
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro. Orden"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha "
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Proveedor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nro OT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sección OT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contacto"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Condicion Pago"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Sub - Total"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "% Desc."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "(S/.) Desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Neto"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "I.G.V"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Total"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Observación"
         Object.Width           =   6174
      EndProperty
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
            Picture         =   "frmInfOrdCom.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":18AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":1D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":215C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":25B4
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":26C6
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":27D8
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":28EA
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":29FC
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":2B0E
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":2C20
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":2D32
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":2E44
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":2F56
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":3068
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":317A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":328C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":339E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":34B0
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":35C2
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":3A14
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInfOrdCom.frx":3E66
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInfOrdCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrWhere As String
Dim mstrnombre As String

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
    
    If lvwOrdCom.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    'If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.

    'Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    
    gstrSql = "CREATE TABLE T_REPORTEOC (NroOC INT,"
    gstrSql = gstrSql & " Proveedor text,"
    gstrSql = gstrSql & " NroOT text,"
    gstrSql = gstrSql & " Seccion text,"
    gstrSql = gstrSql & " Contacto text,"
    gstrSql = gstrSql & " Condicion text,"
    gstrSql = gstrSql & " Fecha date,"
    gstrSql = gstrSql & " Subtotal CURRENCY,"
    gstrSql = gstrSql & " PDesc FLOAT,"
    gstrSql = gstrSql & " MDesc CURRENCY,"
    gstrSql = gstrSql & " Neto CURRENCY, "
    gstrSql = gstrSql & " Iva CURRENCY, "
    gstrSql = gstrSql & " Total CURRENCY)"
    Dbsnueva.Execute gstrSql
    
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_REPORTEOC")
    For i = 1 To lvwOrdCom.ListItems.Count
        Set lvwOrdCom.SelectedItem = lvwOrdCom.ListItems(i)
        Tabla.AddNew
        Tabla!NroOC = IIf(lvwOrdCom.SelectedItem = "", " ", CLng(Val(lvwOrdCom.SelectedItem)))
        Tabla!Proveedor = IIf(lvwOrdCom.SelectedItem.SubItems(2) = "", " ", lvwOrdCom.SelectedItem.SubItems(2))
        Tabla!NroOT = IIf(lvwOrdCom.SelectedItem.SubItems(3) = "", " ", lvwOrdCom.SelectedItem.SubItems(3))
        Tabla!Seccion = IIf(lvwOrdCom.SelectedItem.SubItems(4) = "", " ", lvwOrdCom.SelectedItem.SubItems(4))
        Tabla!Contacto = IIf(lvwOrdCom.SelectedItem.SubItems(5) = "", " ", lvwOrdCom.SelectedItem.SubItems(5))
        Tabla!Condicion = IIf(lvwOrdCom.SelectedItem.SubItems(6) = "", " ", lvwOrdCom.SelectedItem.SubItems(6))
        Tabla!Fecha = DateValue(IIf(lvwOrdCom.SelectedItem.SubItems(1) = "", " ", lvwOrdCom.SelectedItem.SubItems(1)))
        Tabla!SubTotal = IIf(lvwOrdCom.SelectedItem.SubItems(7) = "", " ", CCur(SacarFormatoValor(lvwOrdCom.SelectedItem.SubItems(7), "")))
        Tabla!PDesc = IIf(lvwOrdCom.SelectedItem.SubItems(8) = "", " ", lvwOrdCom.SelectedItem.SubItems(8))
        Tabla!MDesc = IIf(lvwOrdCom.SelectedItem.SubItems(9) = "", " ", CCur(SacarFormatoValor(lvwOrdCom.SelectedItem.SubItems(9), "")))
        Tabla!Neto = IIf(lvwOrdCom.SelectedItem.SubItems(10) = "", " ", CCur(SacarFormatoValor(lvwOrdCom.SelectedItem.SubItems(10), "")))
        Tabla!IVA = IIf(lvwOrdCom.SelectedItem.SubItems(11) = "", " ", CCur(SacarFormatoValor(lvwOrdCom.SelectedItem.SubItems(11), "")))
        Tabla!Total = IIf(lvwOrdCom.SelectedItem.SubItems(12) = "", " ", CCur(SacarFormatoValor(lvwOrdCom.SelectedItem.SubItems(12), "")))
        Tabla.Update
    Next i
   Tabla.Close
   Dbsnueva.Close
   With rptOrdCom
        .ReportFileName = gstrPathReporte & "\LOCS.rpt"
        .WindowTitle = "Reporte de Ordenes de Compra"
        '.DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='LISTADO DE ORDENES DE COMPRA'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "Tdecimales=" & gintDecimalesMoneda
        .Formulas(6) = "NombreIva='" & gstrNombreIva & "'"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = True
   End With
   
   
   Screen.MousePointer = 1

End Sub
Private Sub cckCriterio_Click(Index As Integer)
Select Case Index

Case 0                                  '//////     nro inicio
    If cckCriterio(Index).Value = 0 Then
        txtDesde = ""
        txtDesde.Enabled = False
    Else
        txtDesde.Enabled = True
        txtDesde.SetFocus
    End If
Case 1  '//////     nro fINAL
    If cckCriterio(Index).Value = 0 Then
        txtHasta = ""
        txtHasta.Enabled = False
    Else
        txtHasta.Enabled = True
        txtHasta.SetFocus
    End If
Case 2  '//////     FECHA INICIAL
    If cckCriterio(Index).Value = 0 Then
        pckIni.Enabled = False
    Else
        pckIni.Enabled = True
        pckIni.SetFocus
    End If
Case 3  '//////     FECHA FINAL
    If cckCriterio(Index).Value = 0 Then
        pckFin.Enabled = False
    Else
        pckFin.Enabled = True
        pckFin.SetFocus
    End If
Case 4  '//////     PROVEEDOR
    If cckCriterio(Index).Value = 0 Then
        txtProv = ""
        txtProv.Enabled = False
        tlbProv.Enabled = False
    Else
        txtProv.Enabled = True
        tlbProv.Enabled = True
        txtProv.SetFocus
    End If
End Select

End Sub

Private Sub cmdBuscar_Click()
lvwOrdCom.ListItems.Clear
mstrWhere = ""

If cckCriterio(0).Value = 1 Then    'DESDE
    If cckCriterio(1).Value = 1 Then    ' HASTA
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & "  AND Tllr_Orden_Compra.Id_Orden >= " & CLng(Val(txtDesde)) & " AND Tllr_Orden_Compra.Id_Orden <= " & CLng(Val(txtHasta)) & " "
        Else
            mstrWhere = "  WHERE Tllr_Orden_Compra.Id_Orden >= " & CLng(Val(txtDesde)) & " AND Tllr_Orden_Compra.Id_Orden <= " & CLng(Val(txtHasta)) & " "
        End If
    Else
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & "  AND Tllr_Orden_Compra.Id_Orden = " & CLng(Val(txtDesde)) & " "
        Else
            mstrWhere = "  WHERE Tllr_Orden_Compra.Id_Orden = " & CLng(Val(txtDesde)) & " "
        End If
    End If
Else
    If cckCriterio(1).Value = 1 Then
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & "  AND Tllr_Orden_Compra.Id_Orden = " & CLng(Val(txtHasta)) & " "
        Else
            mstrWhere = "  WHERE Tllr_Orden_Compra.Id_Orden = " & CLng(Val(txtHasta)) & " "
        End If
    End If
End If

If cckCriterio(2).Value = 1 Then    'FECHA INICIO
    If cckCriterio(3).Value = 1 Then    ' FECHA TERMINO
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & "  AND Tllr_Orden_Compra.Fecha_Orden >= '" & CDate(pckIni.Value) & "' AND Tllr_Orden_Compra.Fecha_Orden <= '" & CDate(pckFin.Value) & " 23:59:59' "
        Else
            mstrWhere = "  WHERE Tllr_Orden_Compra.Fecha_Orden >= '" & CDate(pckIni.Value) & "' AND Tllr_Orden_Compra.Fecha_Orden <= '" & CDate(pckFin.Value) & " 23:59:59' "
        End If
    Else
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & "  AND Tllr_Orden_Compra.Fecha_Orden = '" & CDate(pckIni.Value) & "' "
        Else
            mstrWhere = "  WHERE Tllr_Orden_Compra.Fecha_Orden = '" & CDate(pckIni.Value) & "' "
        End If
    End If
Else
    If cckCriterio(3).Value = 1 Then
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & "  AND Tllr_Orden_Compra.Fecha_Orden = '" & CDate(pckFin.Value) & " 23:59:59' "
        Else
            mstrWhere = "  WHERE Tllr_Orden_Compra.Fecha_Orden = '" & CDate(pckFin.Value) & " 23:59:59' "
        End If
    End If
End If

If cckCriterio(4).Value = 1 Then
    If mstrWhere <> "" Then
        mstrWhere = mstrWhere & "  AND Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(txtProv, "Comienzo del Campo", apSqlServer) & "' "
    Else
        mstrWhere = "  WHERE Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(txtProv, "Comienzo del Campo", apSqlServer) & "' "
    End If
End If

If mstrWhere <> "" Then
    mstrWhere = mstrWhere & " AND Id_Empresa ='" & gstrIdEmpresa & "' AND  Id_Sucursal='" & gstrIdSucursal & "'"
Else
    mstrWhere = " WHERE Id_Empresa ='" & gstrIdEmpresa & "' AND  Id_Sucursal='" & gstrIdSucursal & "'"
End If

gstrSql = "SELECT Tllr_Orden_Compra.Id_Orden AS NROORDEN,"
gstrSql = gstrSql & " Tllr_Orden_Compra.Proveedor,"
gstrSql = gstrSql & " Glbl_Cliente_Proveedor.Razon_Social as NOMBRE, "
gstrSql = gstrSql & " Tllr_Orden_Compra.Contacto AS CONT,"
gstrSql = gstrSql & " Tllr_Orden_Compra.Condicion_Pago AS CONP,"
gstrSql = gstrSql & " Tllr_Orden_Compra.Fecha_Orden AS FECOR,"
gstrSql = gstrSql & " Tllr_Orden_Compra.Observacion AS OBS,"
gstrSql = gstrSql & " Tllr_Orden_Compra.SubTotal AS STOT,"
gstrSql = gstrSql & " Tllr_Orden_Compra.Porcentaje_Descuento AS PDESC,"
gstrSql = gstrSql & " Tllr_Orden_Compra.Descuento AS MDESC, "
gstrSql = gstrSql & " Tllr_Orden_Compra.Neto AS NETO,"
gstrSql = gstrSql & " Tllr_Orden_Compra.Iva AS IVA, "
gstrSql = gstrSql & " Tllr_Orden_Compra.Total AS TOT,"
gstrSql = gstrSql & " Tllr_Orden_Compra.NroOT AS OT,"
gstrSql = gstrSql & " Tllr_Orden_Compra.SeccionOT AS SEC"
gstrSql = gstrSql & " FROM Tllr_Orden_Compra LEFT OUTER JOIN Glbl_Cliente_Proveedor ON Tllr_Orden_Compra.Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor"
gstrSql = gstrSql & mstrWhere
gstrSql = gstrSql & " Order by Tllr_Orden_Compra.Id_Orden"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set glsiItem = lvwOrdCom.ListItems.Add(, , !NROORDEN)
                glsiItem.SubItems(1) = Format(!FECOR, "dd/mm/yyyy")
                glsiItem.SubItems(2) = !Nombre
                glsiItem.SubItems(3) = IIf(Not IsNull(!OT), !OT, "Sin O/T")
                glsiItem.SubItems(4) = IIf(Not IsNull(!Sec), IIf(!Sec = "M", "MECANICA", "CARROCERIA"), "")
                glsiItem.SubItems(5) = !cont
                glsiItem.SubItems(6) = !CONP
                glsiItem.SubItems(7) = FormatoValor(!STOT, "", gintDecimalesMoneda)
                glsiItem.SubItems(8) = FormatoValor(!PDesc, "", 2)
                glsiItem.SubItems(9) = FormatoValor(!MDesc, "", gintDecimalesMoneda)
                glsiItem.SubItems(10) = FormatoValor(!Neto, "", gintDecimalesMoneda)
                glsiItem.SubItems(11) = FormatoValor(!IVA, "", gintDecimalesMoneda)
                glsiItem.SubItems(12) = FormatoValor(!TOT, "", gintDecimalesMoneda)
                glsiItem.SubItems(13) = !OBS
                .MoveNext
            Wend
        End If
    End With
    
End If

Conexion.CloseHost gadoPrincipal

End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
ImprimirConsulta
End Sub


Private Sub cmdSeleccionar_Click()
If Not Me.lvwOrdCom.SelectedItem Is Nothing Then
    gstrBusca = Me.lvwOrdCom.SelectedItem
End If
Unload Me
End Sub

Private Sub Form_Activate()
    If Not Atributos("Glbl", "Tllr_30_0020", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.pckIni = BOM(Date)
    Me.pckFin = EOM(Date)
    Me.lvwOrdCom.ColumnHeaders(12).Text = gstrNombreIva
End Sub

Private Sub lvwOrdCom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvwOrdCom, ColumnHeader
End Sub


Private Sub lvwOrdCom_DblClick()
If cmdSeleccionar.Enabled = True Then cmdSeleccionar.Value = True
End Sub

Private Sub tlbProv_ButtonClick(ByVal Button As MSComctlLib.Button)

If Button.Key = "Buscar" Then
'    apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre, gstrIdEmpresa
'    'apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre
'    txtProv.Tag = gstrBusca
'    txtProv.Text = mstrnombre
gstrRutCliente = ""
gstrNombreCliente = ""
Libreria.ClienteBuscar Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
     If gstrRutCliente <> "" Then
        txtProv.Text = gstrNombreCliente
        txtProv.Tag = gstrRutCliente
    End If

End If

End Sub


