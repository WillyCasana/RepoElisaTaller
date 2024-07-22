VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHistoricoOT 
   Caption         =   "Histórico de Orden de Trabajo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   Icon            =   "frmHistoricoOT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgTitulo 
      Left            =   720
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":038A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":049C
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":05AE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":06C0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":07D2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":08E4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":09F6
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":0B08
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":0C1A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":0D2C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":0E3E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":0F50
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":1062
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":1174
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":1286
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":1398
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":14AA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":18FC
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":1D4E
            Key             =   "CopiarX"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":1E60
            Key             =   "AgregarSucursal"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":22B4
            Key             =   "VerSucursal"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":2A28
            Key             =   "Abrir"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":2B48
            Key             =   "Horizontal"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":383C
            Key             =   "Resalte"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":3C90
            Key             =   "Cerrar2"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":3DEC
            Key             =   "Reset"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":3F48
            Key             =   "Config_Col"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":42E4
            Key             =   "otro"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":4738
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":4A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":4EA8
            Key             =   "Categorizar"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":52FC
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":541C
            Key             =   "Pegar"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":553C
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":5A80
            Key             =   "Categorias"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":5B94
            Key             =   "UpDown"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":5EE8
            Key             =   "Excel2"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":623C
            Key             =   "malo"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":6590
            Key             =   "malo2"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoricoOT.frx":68E4
            Key             =   "selectcamp"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbTitulo 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imgTitulo"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Vista Previa"
            ImageKey        =   "Preview"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Exportar a Microsoft Excel"
            ImageKey        =   "Excel2"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageKey        =   "Copiar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ordenar"
            Object.ToolTipText     =   "Ordenar"
            ImageKey        =   "SortAsc"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Vertical"
            Object.ToolTipText     =   "Vista Vertical de los Datos"
            ImageKey        =   "Horizontal"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Columnas"
            Object.ToolTipText     =   "Selector de Columnas"
            ImageKey        =   "selectcamp"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvHistorico 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Acción"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Usuario"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   0
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmHistoricoOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
TraeHistorial
End Sub
Private Sub TraeHistorial()
Dim lstrSQL As String
Dim adoPrincipal As New ADODB.Recordset
Dim Item As ListItem

lstrSQL = "SELECT * FROM Tllr_OT_Historico"
lstrSQL = lstrSQL & " WHERE Id_OT ='" & frmRecepcion.lblNroRecepcion & "'"
lstrSQL = lstrSQL & " And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"

Me.lsvHistorico.ListItems.Clear
If Conexion.SendHost(lstrSQL, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    adoPrincipal.MoveFirst
    End If
            Do Until adoPrincipal.EOF
            Set Item = Me.lsvHistorico.ListItems.Add(, , adoPrincipal!ID)
                Item.SubItems(1) = ValorNulo(adoPrincipal!Movto)
                Item.SubItems(2) = ValorNulo(adoPrincipal!Usr_Id)
                Item.SubItems(3) = ValorNulo(adoPrincipal!Usr_Fecha)
                adoPrincipal.MoveNext
            Loop
End If
Conexion.CloseHost adoPrincipal

End Sub

Private Sub tlbTitulo_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lstrSQL As String
Dim lvarPaso As Variant

Select Case Button.Key
    Case "Imprimir"
        
    Case "Preview"
        
    Case "Excel"
        ExportarDatos Me.lsvHistorico, Me.cdExportar, Me.hwnd
    Case "Copiar"
        
    Case "Ordenar"
        
    Case "Vertical"
        
    Case "Columnas"
        Set gfrmFormPadre = Me
        frmColumnas.Show vbModal
End Select

End Sub

