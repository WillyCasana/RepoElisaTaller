VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTemparios 
   Caption         =   "Temparios"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   Icon            =   "frmTemparios.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8550
   Begin VB.Frame frFiltros 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Width           =   5775
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chlDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frTitulos 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   5835
      Begin MSComctlLib.ListView lvColTitulo 
         Height          =   345
         Index           =   0
         Left            =   4320
         TabIndex        =   10
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
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
         Enabled         =   0   'False
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   2011
         EndProperty
      End
      Begin MSComctlLib.ListView lvFilasTitulo 
         Height          =   345
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   609
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
         Enabled         =   0   'False
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "DESCRIPCION"
            Object.Width           =   5644
         EndProperty
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3135
      Left            =   6000
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.HScrollBar HScrol 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5880
      Width           =   5895
   End
   Begin MSComctlLib.Toolbar BarraHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fila"
            Object.ToolTipText     =   "Agregar Fila"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Columna"
            Object.ToolTipText     =   "Agregar Columna"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EliminaFila"
            Object.ToolTipText     =   "Elimina Fila"
            ImageIndex      =   33
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EliminaCol"
            Object.ToolTipText     =   "Elimina Columna"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Seleccionar"
            Object.ToolTipText     =   "Selecciona Item Chequeados"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   6360
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":18AC
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":19BE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":1AD0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":1BE2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":1CF4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":1E06
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":1F18
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":202A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":213C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":224E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":2360
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":2472
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":2584
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":2696
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":27A8
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":28BA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":2D0C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":315E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":3270
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":33CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":3528
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":3684
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":37E0
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":42AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":4700
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":4864
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":4CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":4F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":53A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":57FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":5D40
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":6284
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":6864
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTemparios.frx":6D3C
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frCabecera 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   675
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   8475
      Begin VB.OptionButton opcValores 
         Caption         =   "Precio Costo"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   18
         Top             =   150
         Width           =   1455
      End
      Begin VB.OptionButton opcValores 
         Caption         =   "Precio Venta"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   17
         Top             =   410
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc adoCiaSeguro 
         Height          =   330
         Left            =   1200
         Top             =   300
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
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
      Begin MSDataListLib.DataCombo dbcCiaSeguro 
         Bindings        =   "frmTemparios.frx":708E
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   300
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cía. Seguro:"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.Frame frParche1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Width           =   5835
   End
   Begin VB.Frame frParche2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   -120
      TabIndex        =   13
      Top             =   6240
      Width           =   5835
   End
   Begin VB.Frame frFondo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   6075
      Begin MSComctlLib.ListView lvCol 
         Height          =   3075
         Index           =   0
         Left            =   4320
         TabIndex        =   2
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   5424
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lvFilas 
         Height          =   3075
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   5424
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
            Text            =   "DESCRIPCION"
            Object.Width           =   5644
         EndProperty
      End
   End
   Begin VB.Label lblTipo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDescripcion 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblCodigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmTemparios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCol As Integer
Dim IntIndiceColumna As Integer
Dim SW As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean


Private Sub BarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
  Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Buscar"
            Buscar
        Case "Fila"
            AgregarFila
        Case "Columna"
            AgregarColumna
        Case "Grabar"
            Grabar
        Case "Imprimir"
           Imprimir
        Case "Cerrar"
            Unload Me
        Case "Seleccionar"
            SeleccionarItem
            AsignaTotal
            TotalFinal
            'Unload Me
        Case "EliminaFila"
            If Me.lvFilas.ListItems.Count > 0 Then
                EliminarFila lvFilas.SelectedItem.Index
            End If
        Case "EliminaCol"
            If intCol > 0 Then
                EliminarColumna
            End If
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Buscar()
    Dim AdoTemp As New ADODB.Recordset
    Dim adoTemp1 As New ADODB.Recordset
    Dim strSql As String
    Dim strWhere As String
    Dim j As Integer
    Dim Item As ListItem
    Dim dblValor As Double
    Dim dblValor1 As Double
    
    If Me.dbcCiaSeguro.BoundText = "" Then
        MsgBox "Debe seleccionar una Compañía de Seguros", vbInformation, "Advertencia"
        Me.dbcCiaSeguro.SetFocus
        Exit Sub
    End If
    
    Me.lvFilas.ListItems.Clear
    '//Descarga las columnas...
    If intCol > 0 Then
        For j = intCol - 1 To 1 Step -1
            Unload Me.lvCol(j)
            Unload Me.lvColTitulo(j)
       Next
    End If
    Me.lvCol(0).ListItems.Clear
    Me.lvColTitulo(0).ListItems.Clear
    
    intCol = 0
    
    '//Carga Columnas...
    strSql = "select * from Tllr_Tempario_Columna where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' order by posicion"
    If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        While Not AdoTemp.EOF
           
            If intCol <> 0 Then
                Load Me.lvCol(intCol)
                Load Me.lvColTitulo(intCol)
                Me.lvCol(intCol).Left = (Me.lvCol(intCol - 1).Left + 1500)
                Me.lvColTitulo(intCol).Left = Me.lvCol(intCol).Left
            End If
            Me.lvColTitulo(intCol).ColumnHeaders(2).Text = ValorNulo(AdoTemp!Descripcion)
            Me.lvColTitulo(intCol).Tag = ValorNulo(AdoTemp!D_P)
            Me.lvColTitulo(intCol).Enabled = False
            
            
            Me.lvCol(intCol).Tag = AdoTemp!id_columna
            
            Me.lvCol(intCol).Visible = True
                       
            
            Me.lvColTitulo(intCol).Visible = True
                        
            
            Me.lvCol(intCol).Refresh
            Me.lvColTitulo(intCol).Refresh
          
            
            
            
            intCol = intCol + 1
            AdoTemp.MoveNext
        Wend
    End If
    
    '//Carga Filas...
    
    strWhere = ""
    If Me.chkCodigo.Value = 1 Then
        strWhere = " And Id_fila Like '%" & Me.txtCodigo & "%'"
    End If
    If Me.chlDescripcion.Value = 1 Then
        strWhere = " And Descripcion Like '%" & Me.txtDescripcion & "%'"
    End If
    
    strSql = "select * from Tllr_Tempario_Fila where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "'" & strWhere & " Order by Descripcion"
    If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        While Not AdoTemp.EOF
            Set Item = Me.lvFilas.ListItems.Add(, , AdoTemp!id_fila)
            Item.SubItems(1) = AdoTemp!Descripcion
            
            '//Crea celdas...
            For j = 0 To intCol - 1
                Set Item = Me.lvCol(j).ListItems.Add(, , "")
                dblValor = 0
                dblValor1 = 0
                strSql = "select isnull(valor,0) as valor, isnull(costo,0) as costo from Tllr_Tempario_Valor where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_fila='" & AdoTemp!id_fila & "' and id_columna='" & Me.lvCol(j).Tag & "'"
                If Conexion.SendHost(strSql, adoTemp1, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                    If Not adoTemp1.BOF And Not AdoTemp.EOF Then
                            dblValor = adoTemp1!Valor
                            dblValor1 = adoTemp1!Costo
                    End If
                End If
                If Me.opcValores(0).Value Then
                    Item.SubItems(1) = FormatoValor(dblValor, gstrMonedaLocal, gintDecimalesMoneda)
                    Item.SubItems(2) = FormatoValor(dblValor1, gstrMonedaLocal, gintDecimalesMoneda)
                Else
                    Item.SubItems(1) = FormatoValor(dblValor1, gstrMonedaLocal, gintDecimalesMoneda)
                    Item.SubItems(2) = FormatoValor(dblValor1, gstrMonedaLocal, gintDecimalesMoneda)
                End If
            Next
            Conexion.CloseHost adoTemp1
            AdoTemp.MoveNext
        Wend
    End If
    Conexion.CloseHost AdoTemp
    AjustaAnchoFondo
End Sub
Private Sub Imprimir()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim x As Integer
Dim j As Integer
Dim GcamBaseTem As String
Dim ContPaginas As Integer
Dim ContColumnas As Integer
Dim ContFilas As Integer


    'Devuelve la ruta del directorio Windows
    Dim rc As Long
    Dim WinPath As String
    WinPath = Space$(300)
    rc = GetWindowsDirectory(WinPath, 300)
    GcamBaseTem = Trim$(WinPath)
    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
    '---------------------------------------
    
    If Me.lvFilas.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (CAMPO1 text,CAMPO2 text,CAMPO3 text,CAMPO4 text,CAMPO5 text,CAMPO6 text,CAMPO7 text,CAMPO8 text,PAGINA text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    
    ContColumnas = Int((intCol + 1) / 8) + 1
    ContFilas = 1
    ContPaginas = 1
    
    For j = 1 To Me.lvFilas.ListItems.Count  'ciclo total de filas
        If ContFilas = 1 Then  'titulo
            Tabla.AddNew
            i = 0
            Tabla!campo1 = Me.lvFilasTitulo.ColumnHeaders(2).Text
            Tabla!campo2 = Me.lvColTitulo(i).ColumnHeaders(2).Text
            i = i + 1
            Tabla!campo3 = Me.lvColTitulo(i).ColumnHeaders(2).Text
            i = i + 1
            Tabla!campo4 = Me.lvColTitulo(i).ColumnHeaders(2).Text
            i = i + 1
            Tabla!campo5 = Me.lvColTitulo(i).ColumnHeaders(2).Text
            i = i + 1
            Tabla!campo6 = Me.lvColTitulo(i).ColumnHeaders(2).Text
            i = i + 1
            Tabla!campo7 = Me.lvColTitulo(i).ColumnHeaders(2).Text
            i = i + 1
            Tabla!campo8 = Me.lvColTitulo(i).ColumnHeaders(2).Text
            Tabla!PAGINA = ContPaginas
            Tabla.Update
        End If
            
        Tabla.AddNew
        i = 0
        Tabla!campo1 = Me.lvFilas.ListItems(j).SubItems(1)
        Tabla!campo2 = Me.lvCol(i).ListItems(j).SubItems(1)
        i = i + 1
        Tabla!campo3 = Me.lvCol(i).ListItems(j).SubItems(1)
        i = i + 1
        Tabla!campo4 = Me.lvCol(i).ListItems(j).SubItems(1)
        i = i + 1
        Tabla!campo5 = Me.lvCol(i).ListItems(j).SubItems(1)
        i = i + 1
        Tabla!campo6 = Me.lvCol(i).ListItems(j).SubItems(1)
        i = i + 1
        Tabla!campo7 = Me.lvCol(i).ListItems(j).SubItems(1)
        i = i + 1
        Tabla!campo8 = Me.lvCol(i).ListItems(j).SubItems(1)
        Tabla!PAGINA = ContPaginas
        Tabla.Update
        ContFilas = ContFilas + 1
        If ContFilas = 31 Then
            ContFilas = 1
            ContPaginas = ContPaginas + 1
        End If
    Next j
    
   Tabla.Close
   
'   With rptOT
'        '"//MODIFICADO POR FDO DIAZ EL 29/11/2000
'        .ReportFileName = gstrPathReporte & "\RESDED.rpt"
'        .WindowTitle = "Informe de Resumen de Deducibles"
'        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
'        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
'        .Formulas(1) = "TITULO='RESUMEN DE DEDUCIBLES'"
'        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
'        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
'        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
'        .Destination = crptToWindow
'        .Action = True
'   End With
'
   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub
Private Sub AgregarFila()
    Dim strCodigo As String
    Dim strDescripcion As String
    Dim Item As ListItem
    Dim i As Integer
    
    If intCol = 0 Then
        MsgBox "Primero debe agregar las COLUMNAS del Tempario", vbExclamation, "Advertencia"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    frmAgregaFilaColTempario.Tag = "Fila"
    frmAgregaFilaColTempario.Show vbModal
    
    If lblCodigo <> "" Then
        If Validacion("Fila") = False Then
            Exit Sub
        End If
        
        '//Agrega Filas...
        Set Item = Me.lvFilas.ListItems.Add(, , lblCodigo)
        Item.SubItems(1) = lblDescripcion
        
        For i = 0 To intCol - 1
            Set Item = Me.lvCol(i).ListItems.Add(, , "")
            Item.SubItems(1) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
        Next
        
        '//grabar en tllr_tempario_fila
        GrabarFilaTempario lblCodigo, lblDescripcion
        Screen.MousePointer = vbDefault
    
        AjustaAnchoFondo
    End If
End Sub
Private Sub AgregarColumna()
    Dim strCodigo As String
    Dim strDescripcion As String
    Dim strValor As String
    Dim strTipo As String
    
    frmAgregaFilaColTempario.Tag = "Columna"
    frmAgregaFilaColTempario.Show vbModal
    
    If lblCodigo <> "" Then
        If Validacion("Columna") = False Then
            Exit Sub
        End If

        '//Agrega columnas...
        Screen.MousePointer = vbHourglass
        If intCol <> 0 Then
            Load Me.lvCol(intCol)
            Load Me.lvColTitulo(intCol)
            Me.lvCol(intCol).Left = (Me.lvCol(intCol - 1).Left + 1500)
            Me.lvColTitulo(intCol).Left = Me.lvCol(intCol).Left
        End If
        Me.lvColTitulo(intCol).ColumnHeaders(2).Text = lblDescripcion
        Me.lvColTitulo(intCol).Enabled = False
        Me.lvColTitulo(intCol).Tag = lblTipo
        Me.lvCol(intCol).Tag = lblCodigo
        Me.lvCol(intCol).Visible = True
        Me.lvColTitulo(intCol).Visible = True
        Me.lvCol(intCol).Refresh
        Me.lvColTitulo(intCol).Refresh
    
    
        '//Inicializa Columna...
        For i = 1 To Me.lvCol(intCol).ListItems.Count
            If SacarFormatoValor(Me.lvCol(intCol).ListItems(i).SubItems(1), gstrMonedaLocal) <> 0 Then
                Me.lvCol(intCol).ListItems(i).SubItems(1) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
            End If
            If Me.lvCol(intCol).ListItems(i).Checked Or Me.lvCol(intCol).ListItems(i).ListSubItems(1).ForeColor = vbRed Then
                Me.lvCol(intCol).ListItems(i).Checked = False
                Me.lvCol(intCol).ListItems(i).ListSubItems(1).ForeColor = Me.lvFilas.ListItems(1).ForeColor
                Me.lvCol(intCol).ListItems(i).ListSubItems(1).Bold = False
            End If
        Next
    
        intCol = intCol + 1
        
        '//FERNANDO...grabar en tllr_tempario_columna
        GrabarColumnaTempario lblCodigo, lblDescripcion, lblTipo, intCol
        Screen.MousePointer = vbDefault
    
        AjustaAnchoFondo
    End If
End Sub
Private Sub CargaCiaSeguro()
    '//Carga Cia.Seguro...
    Dim strSql As String
    Dim AdoTemp As New ADODB.Recordset
    
    strSql = "select Id_Compañia_Seguro as codigo, nombre from Tllr_Compañia_Seguro order by nombre"
    If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        Set adoCiaSeguro.Recordset = AdoTemp
    End If
    Set AdoTemp = New ADODB.Recordset
End Sub

Private Sub chkCodigo_Click()
    If Me.chkCodigo.Value = 1 Then
        chlDescripcion.Value = 0: Me.txtDescripcion.Text = "": Me.txtDescripcion.Enabled = False
        txtCodigo.Enabled = True
        txtCodigo.SetFocus
    Else
        chkCodigo.Value = 0
        txtCodigo.Enabled = False
        txtCodigo.Text = ""
    End If
End Sub

Private Sub chlDescripcion_Click()
    If chlDescripcion.Value = 1 Then
        chkCodigo.Value = 0: Me.txtCodigo.Text = "": Me.txtCodigo.Enabled = False
        txtDescripcion.Enabled = True
        txtDescripcion.SetFocus
    Else
        chlDescripcion.Value = 0
        txtCodigo.Enabled = True
        txtCodigo.SetFocus
    End If
End Sub

Private Sub Form_Activate()

    If Not SW Then
        SW = True
        If Atributos("Glbl", "Tllr_10_0110_0052", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            Me.opcValores(0).Visible = True
            Me.opcValores(1).Visible = True
        Else
            Me.opcValores(0).Visible = False
            Me.opcValores(1).Visible = False
        End If '/////////ojo
        If gstrProcedencia <> "Mantenedor" Then
            Me.dbcCiaSeguro.BoundText = frmRecepcion.lblCompañia.Tag
            Buscar
        End If

    End If

End Sub
Private Sub Form_Load()
    CargaCiaSeguro
    intCol = 0
    Me.Width = 6225
    Me.Height = 5160
    SW = False
End Sub
Private Sub Form_Resize()
    Dim i As Integer
    Me.frFondo.Left = 0
    Me.frFondo.Top = Me.BarraHerramientas.Height + Me.frCabecera.Height + Me.frTitulos.Height + Me.frFiltros.Height
    Me.frFondo.Width = Me.ScaleWidth
    Me.frTitulos.Width = Me.frFondo.Width
    Me.frFondo.Height = Me.ScaleHeight - (Me.BarraHerramientas.Height + Me.HScrol.Height + Me.frCabecera.Height + Me.lvFilasTitulo.Height + Me.frFiltros.Height)
    
    '//Scroll Horizontal...
    Me.HScrol.Top = Me.frFondo.Top + (Me.frFondo.Height - 50)
    Me.HScrol.Width = Me.frFondo.Width - Me.VScroll.Width
    
    '//Scroll Vertical...
    Me.VScroll.Left = Me.ScaleWidth - Me.VScroll.Width
    Me.VScroll.Height = Me.ScaleHeight - (1020 + (Me.HScrol.Height * 2))

    
    AjustaAnchoFondo
End Sub

Private Sub HScrol_Change()
    Me.frFondo.Left = 0 - (HScrol.Value / 100) * Me.HScrol.Tag
    Me.frTitulos.Left = Me.frFondo.Left
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lvCol_Click(Index As Integer)
    Me.BarraHerramientas.Buttons.Item("EliminaFila").Enabled = False
    Me.BarraHerramientas.Buttons.Item("EliminaCol").Enabled = True
    IntIndiceColumna = Index
End Sub

Private Sub lvCol_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.ListSubItems(1).ForeColor = vbRed
        Item.ListSubItems(1).Bold = True
    Else
        Item.ListSubItems(1).ForeColor = Me.lvFilas.ListItems(1).ForeColor
        Item.ListSubItems(1).Bold = False
    End If
    Item.Selected = True
End Sub
Private Sub lvCol_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strValor As String
    Select Case KeyCode
        Case 39 '//Derecha
            If (Index + 1) < intCol Then
                Me.lvCol(Index + 1).ListItems(Me.lvCol(Index).SelectedItem.Index).Selected = True
                Me.lvCol(Index + 1).SetFocus

            End If
        Case 37
            If (Index - 1) >= 0 Then
                Me.lvCol(Index - 1).ListItems(Me.lvCol(Index).SelectedItem.Index).Selected = True
                Me.lvCol(Index - 1).SetFocus
            End If
    End Select

    If KeyCode = 13 Then '//Enter
        If Index + 1 < intCol Then
            lvCol(Index + 1).ListItems(lvCol(Index).SelectedItem.Index).Selected = True
            Me.lvCol(Index + 1).SetFocus
        Else
            If (lvCol(Index).SelectedItem.Index + 1) <= lvCol(Index).ListItems.Count Then
                lvCol(0).ListItems(lvCol(Index).SelectedItem.Index + 1).Selected = True
                lvCol(0).SetFocus
            End If
        End If
'        If lvCol(Index).SelectedItem.Index < lvCol(Index).ListItems.Count Then
'            lvCol(Index).ListItems(lvCol(Index).SelectedItem.Index + 1).Selected = True
'            lvCol(Index).SetFocus
'        End If
    ElseIf KeyCode = 46 Then '//DEL o BAK
        lvCol(Index).SelectedItem.SubItems(1) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    ElseIf KeyCode = 8 Then
        strValor = SacarFormatoValor(lvCol(Index).SelectedItem.SubItems(1), gstrMonedaLocal)
        If Len(strValor) = 1 Then
            lvCol(Index).SelectedItem.SubItems(1) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
        Else
            lvCol(Index).SelectedItem.SubItems(1) = FormatoValor(Left(strValor, Len(strValor) - 1), gstrMonedaLocal, gintDecimalesMoneda)
        End If
    ElseIf KeyCode >= 48 And KeyCode <= 57 Then
        lvCol(Index).SelectedItem.SubItems(1) = FormatoValor(SacarFormatoValor(lvCol(Index).SelectedItem.SubItems(1), gstrMonedaLocal) & Chr(KeyCode), gstrMonedaLocal, gintDecimalesMoneda)
        KeyCode = 0
    ElseIf KeyCode >= 96 And KeyCode <= 105 Then
        lvCol(Index).SelectedItem.SubItems(1) = FormatoValor(SacarFormatoValor(lvCol(Index).SelectedItem.SubItems(1), gstrMonedaLocal) & Chr(KeyCode - 48), gstrMonedaLocal, gintDecimalesMoneda)
        KeyCode = 0
    End If

End Sub
Private Sub AjustaAnchoFondo()
    If intCol > 0 Then
        Me.frFondo.Width = Me.lvFilas.Width + (intCol * lvCol(0).Width)
    Else
        Me.frFondo.Width = Me.lvFilas.Width + (1 * lvCol(0).Width)
    End If
    Me.frCabecera.Width = Me.frFondo.Width
    Me.frFiltros.Width = Me.frFondo.Width
    Me.frTitulos.Width = Me.frFondo.Width
    Me.frParche1.Width = Me.frFondo.Width
    Me.frParche2.Width = Me.frFondo.Width
    
    Me.frParche1.Top = Me.BarraHerramientas.Top + (Me.BarraHerramientas.Height / 2)
    Me.frParche2.Top = Me.HScrol.Top

    If (Me.frFondo.Width - Me.HScrol.Width) >= 0 Then
        Me.HScrol.Tag = Me.frFondo.Width - Me.HScrol.Width
        Me.HScrol.Max = 100
        Me.HScrol.LargeChange = (Me.HScrol.Width * 100) / Me.frFondo.Width
        Me.HScrol.SmallChange = 5
        Me.HScrol.Enabled = True
    Else
        Me.HScrol.Tag = 0
        Me.HScrol.Max = 100
        Me.HScrol.LargeChange = 100
        Me.HScrol.SmallChange = 5
        Me.HScrol.Enabled = False
    End If
    Me.HScrol.Value = 0
    
    
    '//Ajusta largo de columnas...
    If Me.lvFilas.ListItems.Count > 0 Then
        Me.lvFilas.Height = Me.lvFilas.ListItems(1).Height * (Me.lvFilas.ListItems.Count + 5)
        For j = 0 To intCol - 1
            Me.lvCol(j).Height = Me.lvCol(j).ListItems(1).Height * (Me.lvCol(j).ListItems.Count + 5)
        Next
        Me.frFondo.Height = 1100 + Me.lvFilas.Height '420 + Me.lvFilas.Height
    End If
    
    If (Me.frFondo.Height - Me.VScroll.Height) >= 0 Then
        Me.VScroll.Tag = Me.frFondo.Height - Me.VScroll.Height
        Me.VScroll.Max = 100
        Me.VScroll.LargeChange = (Me.VScroll.Height * 100) / Me.frFondo.Height
        Me.VScroll.SmallChange = 5
        Me.VScroll.Enabled = True
    Else
        Me.VScroll.Tag = 0
        Me.VScroll.Max = 100
        Me.VScroll.LargeChange = 100
        Me.VScroll.SmallChange = 5
        Me.VScroll.Enabled = False
    End If
    Me.VScroll.Value = 0

    Me.frFondo.Left = 0
    'Me.frFondo.Top = 1260

End Sub

Private Sub lvCol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'Select Case Button
'        Case vbRightButton  '//BOTON DERECHO
'            MsgBox "boton derecho"
'End Select
End Sub

Private Sub lvFilas_Click()
    Me.BarraHerramientas.Buttons.Item("EliminaFila").Enabled = True
    Me.BarraHerramientas.Buttons.Item("EliminaCol").Enabled = False
End Sub

Private Sub lvFilas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.ListSubItems(1).Bold = True
        Item.ListSubItems(1).ForeColor = vbRed
    Else
        Item.ListSubItems(1).Bold = False
        Item.ListSubItems(1).ForeColor = Me.lvFilas.ListItems(1).ForeColor
    End If
    Item.Selected = True
End Sub

Private Sub opcValores_Click(Index As Integer)
    If opcValores(0).Value Then
        Me.BarraHerramientas.Buttons("Seleccionar").Enabled = True
    Else
        Me.BarraHerramientas.Buttons("Seleccionar").Enabled = False
    End If
    Buscar
End Sub
Private Sub VScroll_Change()
    Me.frFondo.Top = 2400 - ((VScroll.Value / 100) * Me.VScroll.Tag)
End Sub
Private Sub Grabar()
    Dim AdoTemp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Integer
    Dim j As Integer
    Dim SW As Boolean
    
    Screen.MousePointer = vbHourglass
    
    '//Recorre Filas...
    For i = 1 To lvFilas.ListItems.Count
        '//Recorre Columnas...
        For j = 0 To intCol - 1
            SW = False
            strSql = "select * from Tllr_Tempario_Valor where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_fila='" & Me.lvFilas.ListItems(i).Text & "' and id_columna='" & Me.lvCol(j).Tag & "'"
            If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                If Not AdoTemp.BOF And Not AdoTemp.EOF Then
                    SW = True
                End If
            End If
            If SW Then
                'If SacarFormatoValor(lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal) = 0 Then
                    'strSql = "delete from Tllr_Tempario_Valor "
                    'strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_fila='" & Me.lvFilas.ListItems(i).Text & "' and id_columna='" & Me.lvCol(j).Tag & "'"
                'Else
                    If Me.opcValores(0).Value Then
                        strSql = "update Tllr_Tempario_Valor set valor=" & SacarFormatoValor(lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal) & ", usr_id='" & gstrIdUsuario & "', usr_fecha='" & Format(Date, "dd/mm/yyyy") & "', costo=" & SacarFormatoValor(lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal) & " "
                        strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_fila='" & Me.lvFilas.ListItems(i).Text & "' and id_columna='" & Me.lvCol(j).Tag & "'"
                    Else
                        strSql = "update Tllr_Tempario_Valor set usr_id='" & gstrIdUsuario & "', usr_fecha='" & Format(Date, "dd/mm/yyyy") & "', costo=" & SacarFormatoValor(lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal) & " "
                        strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_fila='" & Me.lvFilas.ListItems(i).Text & "' and id_columna='" & Me.lvCol(j).Tag & "'"
                    End If
                'End If
                Conexion.SendHost strSql, , , , 10
            Else
                If SacarFormatoValor(lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal) <> 0 Or SacarFormatoValor(lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal) <> 0 Then
                    If Me.opcValores(0).Value Then
                        strSql = "insert into Tllr_Tempario_Valor (id_empresa, id_ciaseguro, id_fila, id_columna, valor, costo, usr_id, usr_fecha) "
                        strSql = strSql & "values( '" & gstrIdEmpresa & "', '" & Me.dbcCiaSeguro.BoundText & "', '" & Me.lvFilas.ListItems(i).Text & "', '" & Me.lvCol(j).Tag & "', " & SacarFormatoValor(lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal) & ", " & SacarFormatoValor(lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal) & ", '" & gstrIdUsuario & "', '" & Format(Date, "dd/mm/yyyy") & "' )"
                    Else
                        strSql = "insert into Tllr_Tempario_Valor (id_empresa, id_ciaseguro, id_fila, id_columna, costo, usr_id, usr_fecha) "
                        strSql = strSql & "values( '" & gstrIdEmpresa & "', '" & Me.dbcCiaSeguro.BoundText & "', '" & Me.lvFilas.ListItems(i).Text & "', '" & Me.lvCol(j).Tag & "', " & SacarFormatoValor(lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal) & ", '" & gstrIdUsuario & "', '" & Format(Date, "dd/mm/yyyy") & "' )"
                    End If
                    Conexion.SendHost strSql, , , , 10
                End If
            End If
            
            
        Next
    Next
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub SeleccionarItem()
    Dim i As Integer
    Dim j As Integer
    Dim por_recargo As Double
    Dim monto_recargo As Double



    If Me.opcValores(1).Value Then
        Exit Sub
    End If
    For i = 1 To Me.lvFilas.ListItems.Count
        For j = 0 To intCol - 1
            If Me.lvCol(j).ListItems(i).Checked Then
                '//FERNANDO...
                'MsgBox "ACCION " & Me.lvColTitulo(J).ColumnHeaders(2).Text
                'MsgBox "PIEZA " & Me.lvFilas.ListItems(I).SubItems(1)
                'MsgBox "VALOR" & Me.lvCol(J).ListItems(I).SubItems(1)
                
                
                If SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal) = "0" Then
                    monto_recargo = 0
                    por_recargo = 0
                Else
                    monto_recargo = (FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal), "", gintDecimalesMoneda) - FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal), "", gintDecimalesMoneda))
                    por_recargo = Round((monto_recargo / FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal), "", gintDecimalesMoneda)) * 100, 2)
                End If
                
                Set glsiItem = frmRecepcion.lvwServiciosCarroceria.ListItems.Add(, , "Concepto Carroceria")
                glsiItem.SubItems(1) = "01"  'dtcConceptos.BoundText
                glsiItem.SubItems(2) = Me.lvColTitulo(j).ColumnHeaders(2).Text & " " & Me.lvFilas.ListItems(i).SubItems(1)   'txtSeccion
                glsiItem.SubItems(3) = Me.lvColTitulo(j).Tag
                glsiItem.SubItems(4) = "01"   'dtcPartePieza.BoundText
                glsiItem.SubItems(5) = FormatoValor(1, "", 1)
                'glsiItem.SubItems(6) = FormatoValor(0, "", gintDecimalesMoneda)
                glsiItem.SubItems(6) = FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal), "", gintDecimalesMoneda)
                glsiItem.SubItems(7) = por_recargo
                ' glsiItem.SubItems(7) = (FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal), "", gintDecimalesMoneda) - FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(2), gstrMonedaLocal), "", gintDecimalesMoneda))
                
                glsiItem.SubItems(8) = FormatoValor(monto_recargo, "", 0)
                
                glsiItem.SubItems(9) = FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal), "", gintDecimalesMoneda)
                glsiItem.SubItems(10) = FormatoValor(0, "", 2)
                glsiItem.SubItems(11) = FormatoValor(0, "", gintDecimalesMoneda)
                glsiItem.SubItems(12) = TraeCargoDes(gstrIdCargo)
                glsiItem.SubItems(13) = gstrIdCargo
                glsiItem.SubItems(14) = ""
                glsiItem.SubItems(15) = ""
                glsiItem.SubItems(16) = FormatoValor(SacarFormatoValor(Me.lvCol(j).ListItems(i).SubItems(1), gstrMonedaLocal), "", gintDecimalesMoneda)
                glsiItem.SubItems(17) = "N"
                glsiItem.SubItems(18) = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
                IncrementaCorrelativoTrabajosTerceros gstrIdEmpresa, gstrIdSucursal
            End If
        Next
    Next
End Sub
Private Sub GrabarColumnaTempario(Codigo As String, Descripcion As String, Tipo As String, Posicion As Integer)
    
    strSql = "Insert Into Tllr_Tempario_Columna (Id_empresa,Id_CiaSeguro,Id_Columna,Descripcion,Posicion,D_P,Usr_Id,Usr_Fecha) Values ("
    strSql = strSql & "'" & gstrIdEmpresa & "',"
    strSql = strSql & "'" & Me.dbcCiaSeguro.BoundText & "',"
    strSql = strSql & "'" & Codigo & "',"
    strSql = strSql & "'" & Descripcion & "',"
    strSql = strSql & Posicion & ","
    strSql = strSql & "'" & Tipo & "',"
    strSql = strSql & "'" & gstrIdUsuario & "',"
    strSql = strSql & "'" & Format(Now, "dd/mm/yyyy") & "')"
    Conexion.SendHost strSql, , , , 10
    
End Sub
Private Sub GrabarFilaTempario(Codigo As String, Descripcion As String)
    
    strSql = "Insert Into Tllr_Tempario_Fila (Id_empresa,Id_CiaSeguro,Id_Fila,Descripcion,Usr_Id,Usr_Fecha) Values ("
    strSql = strSql & "'" & gstrIdEmpresa & "',"
    strSql = strSql & "'" & Me.dbcCiaSeguro.BoundText & "',"
    strSql = strSql & "'" & Codigo & "',"
    strSql = strSql & "'" & Descripcion & "',"
    strSql = strSql & "'" & gstrIdUsuario & "',"
    strSql = strSql & "'" & Format(Now, "dd/mm/yyyy") & "')"
    Conexion.SendHost strSql, , , , 10
    
End Sub
Sub AsignaTotal()

    frmRecepcion.stbTotalCarroceria.Panels(2).Text = FormatoValor(TotalSeccion(frmRecepcion.lvwServiciosCarroceria, 16), "", gintDecimalesMoneda)
    frmRecepcion.stbTotalDesabolladura.Panels(2).Text = FormatoValor(SubTotalDesabolladura, "", gintDecimalesMoneda)
    frmRecepcion.stbTotalPintura.Panels(2).Text = FormatoValor(SubTotalPintura, "", gintDecimalesMoneda)
    frmRecepcion.stbTotalArmeyDesarme.Panels(2).Text = FormatoValor(SubTotalArmeDesarme, "", gintDecimalesMoneda)
    
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
Function SubTotalDesabolladura() As Double
Dim intS As Integer
Dim dblPreSuma As Double

dblPreSuma = 0
With frmRecepcion.lvwServiciosCarroceria
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
With frmRecepcion.lvwServiciosCarroceria
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
With frmRecepcion.lvwServiciosCarroceria
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        If .SelectedItem.SubItems(3) = "A" Then
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(.SelectedItem.SubItems(16), ""))
        End If
    Next
End With
SubTotalArmeDesarme = dblPreSuma
End Function


Function TotalOT() As Double
Dim dblSemiTotal As Double
With frmRecepcion
    dblSemiTotal = Val(SacarFormatoValor(.stbTotalMec.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalCarroceria.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalOtros.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalTerceros.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalRepuestos.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + IIf(Not IsNull(gcurInsumo), gcurInsumo, 0)
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.stbTotalMateriales.Panels(2).Text, ""))
    dblSemiTotal = dblSemiTotal + Val(SacarFormatoValor(.StbLubricantes.Panels(2).Text, ""))
End With
TotalOT = dblSemiTotal
End Function

Sub TotalFinal()
    frmRecepcion.stbTotalOT.Panels(2).Text = FormatoValor(TotalOT, "", gintDecimalesMoneda)
End Sub

Private Function Validacion(FilaCol As String) As Boolean
Dim Tabla As String
Dim CampoCodigo As String

    If FilaCol = "Fila" Then
        Tabla = "Tllr_Tempario_Fila"
        CampoCodigo = "Id_Fila"
    Else
        Tabla = "Tllr_tempario_Columna"
        CampoCodigo = "Id_Columna"
    End If

    Validacion = True
  
    
    '//Verifica si existe un registro...
    Dim AdoTemp As New ADODB.Recordset
    mstrSql = "Select * from " & Tabla & " Where " & CampoCodigo & "='" & lblCodigo & "' And Id_Empresa='" & gstrIdEmpresa & "' And id_CiaSeguro='" & Me.dbcCiaSeguro.BoundText & "'"
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(AdoTemp!Descripcion), "SIN DESCRIPCION", AdoTemp!Descripcion & "]"), vbInformation, "Advertencia"
            Validacion = False
        End If
    End If
    Conexion.CloseHost AdoTemp
End Function

Sub ImprimeTitulos()
    'titulos
End Sub
Private Sub EliminarFila(intIndice As Integer)
Dim strSql As String
    
    If MsgBox("¿ Desea eliminar la fila " & Me.lvFilas.SelectedItem.SubItems(1), vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        
        strSql = "delete from Tllr_Tempario_Valor "
        strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_fila='" & Me.lvFilas.SelectedItem & "'"
        Conexion.SendHost strSql, , , , 10
        
        strSql = "delete from Tllr_Tempario_Fila "
        strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_fila='" & Me.lvFilas.SelectedItem & "'"
        Conexion.SendHost strSql, , , , 10
        
        Me.lvFilas.ListItems.Remove (intIndice)
        
        For i = 0 To intCol - 1
            Me.lvCol(i).ListItems.Remove (intIndice)
        Next
    End If
End Sub
Private Sub EliminarColumna()
Dim strSql As String
    
    If MsgBox("¿ Desea eliminar la columna " & Me.lvColTitulo(IntIndiceColumna).ColumnHeaders(2).Text, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        
        strSql = "delete from Tllr_Tempario_Valor "
        strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_columna='" & Me.lvCol(IntIndiceColumna).Tag & "'"
        Conexion.SendHost strSql, , , , 10
        
        strSql = "delete from Tllr_Tempario_Columna "
        strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_ciaseguro='" & Me.dbcCiaSeguro.BoundText & "' and id_columna='" & Me.lvCol(IntIndiceColumna).Tag & "'"
        Conexion.SendHost strSql, , , , 10
        
        If IntIndiceColumna <> 0 Then
            If IntIndiceColumna = intCol - 1 Then
                Unload Me.lvCol(IntIndiceColumna)
                Unload Me.lvColTitulo(IntIndiceColumna)
                intCol = intCol - 1
            Else
                Buscar
            End If
        Else
            Buscar
        End If
        
    End If
End Sub

