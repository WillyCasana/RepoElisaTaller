VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmordencompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Compra"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmordencompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   10185
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker lblfecha 
      Height          =   264
      Left            =   9744
      TabIndex        =   35
      Top             =   336
      Width           =   1692
      _ExtentX        =   2963
      _ExtentY        =   476
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36781
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6720
      Top             =   336
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   3
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle"
      Height          =   3624
      Left            =   84
      TabIndex        =   33
      Top             =   2352
      Width           =   11352
      Begin MSAdodcLib.Adodc Adomoneda 
         Height          =   312
         Left            =   2016
         Top             =   3276
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   2
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
      Begin MSAdodcLib.Adodc Adosucursal 
         Height          =   312
         Left            =   1092
         Top             =   3276
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   2
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
      Begin MSAdodcLib.Adodc adoencargado 
         Height          =   312
         Left            =   168
         Top             =   3276
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   2
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
      Begin VB.CommandButton cmdeliminardetalle 
         Caption         =   "Eliminar"
         Height          =   348
         Left            =   9996
         TabIndex        =   9
         Top             =   3192
         Width           =   1272
      End
      Begin VB.CommandButton cmdagregardetalle 
         Caption         =   "Agregar"
         Height          =   348
         Left            =   8652
         TabIndex        =   8
         Top             =   3192
         Width           =   1272
      End
      Begin MSComctlLib.ListView lsvdetalle 
         Height          =   2868
         Left            =   84
         TabIndex        =   7
         Top             =   252
         Width           =   11184
         _ExtentX        =   19711
         _ExtentY        =   5054
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "nro"
            Text            =   "Item"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pieza"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio Unitario"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Desc./Carg. (%)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Subtotal"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FechaEntrega"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "familia"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ItemReal"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1776
      Left            =   7476
      TabIndex        =   20
      Top             =   6048
      Width           =   3960
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   288
         Left            =   2016
         TabIndex        =   30
         Top             =   1428
         Width           =   1776
      End
      Begin VB.TextBox txtiva 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   288
         Left            =   2016
         TabIndex        =   29
         Top             =   924
         Width           =   1776
      End
      Begin VB.TextBox txtneto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   288
         Left            =   2016
         TabIndex        =   28
         Top             =   672
         Width           =   1776
      End
      Begin VB.TextBox txtdescuento 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   2016
         TabIndex        =   11
         Text            =   "0"
         Top             =   420
         Width           =   1776
      End
      Begin VB.TextBox txtsubtotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   288
         Left            =   2016
         TabIndex        =   27
         Top             =   168
         Width           =   1776
      End
      Begin VB.Line Line1 
         X1              =   2016
         X2              =   3780
         Y1              =   1344
         Y2              =   1344
      End
      Begin VB.Label lbltotal 
         Caption         =   "TOTAL"
         Height          =   180
         Left            =   84
         TabIndex        =   26
         Top             =   1428
         Width           =   684
      End
      Begin VB.Label Label1 
         Caption         =   "I.V.A."
         Height          =   180
         Left            =   84
         TabIndex        =   25
         Top             =   1008
         Width           =   1020
      End
      Begin VB.Label lblneto 
         Caption         =   "Neto"
         Height          =   264
         Left            =   84
         TabIndex        =   24
         Top             =   756
         Width           =   936
      End
      Begin VB.Label lbldescuento 
         Caption         =   "Descuento"
         Height          =   264
         Left            =   84
         TabIndex        =   23
         Top             =   504
         Width           =   1104
      End
      Begin VB.Label lblsubtotal 
         Caption         =   "Sub-Total"
         Height          =   264
         Left            =   84
         TabIndex        =   22
         Top             =   252
         Width           =   936
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones"
      Height          =   1776
      Left            =   84
      TabIndex        =   19
      Top             =   6048
      Width           =   7320
      Begin VB.TextBox txtobservaciones 
         Height          =   1440
         Left            =   84
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmordencompra.frx":0442
         Top             =   252
         Width           =   7152
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1524
      Left            =   84
      TabIndex        =   12
      Top             =   672
      Width           =   11352
      Begin VB.CommandButton cmdcrearproveedor 
         Height          =   315
         Left            =   6132
         Picture         =   "frmordencompra.frx":0448
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   672
         Width           =   315
      End
      Begin VB.TextBox txtnombreproveedor 
         Height          =   288
         Left            =   1176
         TabIndex        =   1
         Top             =   672
         Width           =   4548
      End
      Begin MSDataListLib.DataCombo cmbsucursal 
         Bindings        =   "frmordencompra.frx":054A
         Height          =   288
         Left            =   1176
         TabIndex        =   0
         Top             =   252
         Width           =   4548
         _ExtentX        =   8043
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Sucursal"
         Text            =   "DataCombo2"
      End
      Begin VB.CommandButton cmdbuscaproveedor 
         Height          =   315
         Left            =   5796
         Picture         =   "frmordencompra.frx":0564
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   672
         Width           =   315
      End
      Begin VB.TextBox txtproveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   228
         Left            =   1260
         TabIndex        =   31
         Top             =   252
         Visible         =   0   'False
         Width           =   1104
      End
      Begin VB.TextBox txttipocambio 
         Height          =   288
         Left            =   7728
         TabIndex        =   6
         Top             =   1092
         Width           =   2196
      End
      Begin MSDataListLib.DataCombo cmbmoneda 
         Bindings        =   "frmordencompra.frx":0666
         Height          =   288
         Left            =   7728
         TabIndex        =   5
         Top             =   672
         Width           =   3456
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Moneda"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo cmbencargado 
         Bindings        =   "frmordencompra.frx":067E
         Height          =   288
         Left            =   1176
         TabIndex        =   2
         Top             =   1092
         Width           =   4548
         _ExtentX        =   8043
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Empleado"
         Text            =   "DataCombo2"
         Object.DataMember      =   ""
      End
      Begin VB.TextBox txtcodigo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   288
         Left            =   7728
         TabIndex        =   14
         Top             =   252
         Width           =   2196
      End
      Begin VB.Label lblsucursal 
         Caption         =   "Sucursal"
         Height          =   264
         Left            =   168
         TabIndex        =   32
         Top             =   252
         Width           =   852
      End
      Begin VB.Label lbltipocambio 
         Caption         =   "Tipo Cambio"
         Height          =   264
         Left            =   6636
         TabIndex        =   18
         Top             =   1092
         Width           =   936
      End
      Begin VB.Label lblmoneda 
         Caption         =   "Moneda"
         Height          =   264
         Left            =   6636
         TabIndex        =   17
         Top             =   672
         Width           =   936
      End
      Begin VB.Label lblencargado 
         Caption         =   "Encargado"
         Height          =   264
         Left            =   168
         TabIndex        =   16
         Top             =   1092
         Width           =   1188
      End
      Begin VB.Label lblproveedor 
         Caption         =   "Proveedor"
         Height          =   264
         Left            =   168
         TabIndex        =   15
         Top             =   672
         Width           =   1356
      End
      Begin VB.Label lblnumeroorden 
         Caption         =   "N° ORDEN"
         ForeColor       =   &H8000000D&
         Height          =   264
         Left            =   6636
         TabIndex        =   13
         Top             =   252
         Width           =   1440
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   252
      Top             =   588
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":0699
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":07AB
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":08BD
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":09CF
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":0AE1
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":0BF3
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":0D05
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":0E17
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":0F29
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":103B
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":114D
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":125F
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":1371
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":1483
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":1595
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":16A7
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":17B9
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":1C0B
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordencompra.frx":205D
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
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
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   180
      Left            =   9072
      TabIndex        =   34
      Top             =   420
      Width           =   516
   End
End
Attribute VB_Name = "frmordencompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strnrolista As String
Dim adoPrincipal As ADODB.Recordset
Dim Adopaso As ADODB.Recordset
Dim Adorecordencargado As ADODB.Recordset
Dim Adorecordmoneda As ADODB.Recordset
Dim adorecordsucursal As ADODB.Recordset


Dim mstrSql As String
Dim mblnTablaVacia As Boolean

Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Dim mblnSW As Boolean

Const mcNombreTabla = "Stck_Orden_Compra"
Const mcCampoCodigo = "Id_Orden_Compra"
Const mcCampoCodigo2 = "Id_Sucursal"
Const mcCampoNombre = "Fecha"

Private Sub cmbmoneda_Click(Area As Integer)
Dim X As String
Me.txttipocambio = Retorna_Valor_General("Select Top 1 Paridad From Glbl_Moneda_Tipo_Cambio Where Id_Moneda = '" & Me.cmbmoneda.BoundText & "' Order By Id_Fecha Desc")
X = "Select Top 1 Paridad From Glbl_Moneda_Tipo_Cambio Where Id_Moneda = '" & Me.cmbmoneda.BoundText & "' Order By Id_Fecha Desc"
End Sub

Private Sub cmbmoneda_LostFocus()
Me.txttipocambio = Retorna_Valor_General("Select Top 1 Paridad From Glbl_Moneda_Tipo_Cambio Where Id_Moneda = '" & Me.cmbmoneda.BoundText & "' Order By Id_Fecha Desc")
End Sub

Private Sub cmbsucursal_Click(Area As Integer)
    'If Me.cmbsucursal <> "" And tlbBarraHerramientas.Buttons.Item("Crear").Enabled = False Then
        
    'End If
End Sub

Private Sub cmbsucursal_LostFocus()
    'If Me.cmbsucursal <> "" And tlbBarraHerramientas.Buttons.Item("Crear").Enabled = False Then
    '    Me.txtCodigo = Retorna_Nro_Registro("SELECT Id_Orden_Compra From Stck_Orden_Compra Where Id_Sucursal ='" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IdEmpresa & "'") + 1
    'End If
End Sub

Private Sub cmdagregardetalle_Click()
'Primero debe ver si hay una recepcion para esta orden de compra
Dim strsql As String
strsql = "SELECT Count(Id_Orden_compra) From Stck_Orden_Compra_Recepcion_Detalle Where  Id_Orden_Compra = '" & Me.txtcodigo & "' and  Id_Sucursal ='" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
If Retorna_Valor_General_Tipo2(strsql) < 1 Then    'Ya no se puede modificar por que existe un registro relacionado en la tabla recepcion
    'gintlinea = Me.lsvdetalle.ListItems.Count + 1
    gintlinea = Me.lsvdetalle.ListItems.Count + 1
    frmDetalleOrdenCompra.Show
    gstrBusca = ""
    frmBuscar.Show 1
    If gstrBusca <> "" Then
        frmDetalleOrdenCompra.txtcodigopieza = gstrBusca
        frmDetalleOrdenCompra.txtpv = Retorna_Valor_General("Select Precio_Costo From Stck_Item Where Id_Item = '" & gstrBusca & "'")
        frmDetalleOrdenCompra.txtdescripcion = Retorna_Valor_General("Select Descripcion From Stck_Item Where Id_Item = '" & gstrBusca & "'")
        frmDetalleOrdenCompra.txtcantidad.SetFocus
        frmDetalleOrdenCompra.txtcodigo = Mid(ValorNulo(gstrBusca), InStr(1, ValorNulo(gstrBusca), "°") + 1, Len(ValorNulo(gstrBusca)))
    End If
Else
    MsgBox "Esta orden de compra no puede ser modificada, por haber sido recepcionados uno o más articulos", vbInformation, "Orden de Compra"
End If
End Sub

Private Sub cmdbuscaproveedor_Click()
Dim str1 As String
Dim str2 As String
Form1.BuscarRegistroClientes Conexion, str1, str2
'Form1.BuscarRegistroClientes Conexion, str1, str2, ""
Me.txtproveedor = str1
Me.txtnombreproveedor = str2

'Form1.BuscarRegistroClientes Conexion, Me.txtproveedor, Me.txtnombreproveedor
'Me.txtproveedor = Form1.BuscarRegistros(Conexion, "Glbl_Cliente_Proveedor", "Id_Cliente_Proveedor", "Razon_Social", Me.Caption)
'Me.txtnombreproveedor = Retorna_Valor_General("Select Razon_Social From Glbl_Cliente_Proveedor Where Id_Cliente_Proveedor = '" & Me.txtproveedor & "'")
End Sub
Private Sub cmdcrearproveedor_Click()
    Me.txtproveedor = Form1.clientes(Conexion, USRID, "", "", IDEMPRESA, "", Me.txtproveedor, Me.txtnombreproveedor, apcrear)
End Sub

Private Sub cmdeliminardetalle_Click()
Dim item As ListItem
Dim i As Integer
'Primero debe ver si hay una recepcion para esta orden de compra
Dim strsql As String
strsql = "SELECT Count(Id_Orden_compra) From Stck_Orden_Compra_Recepcion_Detalle Where  Id_Orden_Compra = '" & Me.txtcodigo & "' and  Id_Sucursal ='" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
If Retorna_Valor_General_Tipo2(strsql) < 1 Then    'Ya no se puede modificar por que existe un registro relacionado en la tabla recepcion

        If strnrolista <> "" Then
            If MsgBox("¿Desea Sacar esta linea del detalle. Será eliminada definitivamente?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
                'mstrSql = "Delete From Stck_Orden_compra_Detalle Where Id_Linea = '" & strnrolista & "' and Id_Orden_Compra = '" & Me.txtcodigo & "' and id_sucursal = '" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
                'Ejecuta_Consulta_Ado (mstrSql)
        
            Me.lsvdetalle.ListItems.Remove (Val(strnrolista))
            'Ahora recalcula indices de los numero de lineas
            For i = 1 To Me.lsvdetalle.ListItems.Count
                Me.lsvdetalle.ListItems.item(i) = LTrim(Str(i))
            Next i
        
        '**********************************************************************************************
                'Ahora limpia y carga valores reales
            '    Me.lsvdetalle.ListItems.Clear
            '    mstrSql = "Select * From Stck_Orden_Compra_Detalle Where Id_Orden_Compra = '" & Me.txtcodigo & "' and id_sucursal = '" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
            '    If Conexion.SendHost(mstrSql, adoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            '        If Not (adoPaso.EOF = True And adoPaso.BOF = True) Then
            '            adoPaso.MoveFirst
            '        End If
            '        Do Until adoPaso.EOF
            '            Set Item = frmordencompra.lsvdetalle.ListItems.Add(, , ValorNulo(adoPaso.Fields!id_Linea))
            '            Item.SubItems(1) = Mid(ValorNulo(adoPaso.Fields!Id_Item), InStr(1, ValorNulo(adoPaso.Fields!Id_Item), "°") + 1, Len(ValorNulo(adoPaso.Fields!Id_Item)))
            '            Item.SubItems(2) = Retorna_Valor_General("Select Descripcion From Stck_Item Where Id_Item = '" & ValorNulo(adoPaso.Fields!Id_Item) & "'")
            '            Item.SubItems(3) = ValorNulo(adoPaso.Fields!cantidad)
            '            Item.SubItems(4) = FormatoValor(ValorNulo(adoPaso.Fields!Precio_Unitario), Gsigla, 0)
            '            Item.SubItems(5) = ValorNulo(adoPaso.Fields!descto_recgo)
            '            Item.SubItems(6) = FormatoValor(ValorNulo(adoPaso.Fields!subtotal), Gsigla, 0)
             '           Item.SubItems(7) = ValorNulo(adoPaso.Fields!fecha_entrega)
             '           Item.SubItems(8) = Mid(ValorNulo(adoPaso.Fields!Id_Item), 1, InStr(1, ValorNulo(adoPaso.Fields!Id_Item), "°") - 1)
             '           Item.SubItems(9) = ValorNulo(adoPaso.Fields!Id_Item)
             '           adoPaso.MoveNext
             '       Loop
             '   If CloseHost(adoPaso) = apOk Then
             '   End If
             '   End If
                
                
        '**********************************************************************************************
            'Ahora recalcula indices de los numero de lineas
          '  For i = 1 To Me.lsvdetalle.ListItems.Count
          '      mstrSql = "UPDATE Stck_Orden_Compra_Detalle SET Id_Linea = " + Str(i) _
          '      & " Where Id_Linea = " & Me.lsvdetalle.ListItems(i) & " and Id_Orden_Compra = '" & Me.txtcodigo & "' and id_sucursal = '" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
          '      Ejecuta_Consulta_Ado (mstrSql)
          '  Next i
       '
        
        '**********************************************************************************************
                'Ahora limpia y carga valores reales
        '        Me.lsvdetalle.ListItems.Clear
        '        mstrSql = "Select * From Stck_Orden_Compra_Detalle Where Id_Orden_Compra = '" & Me.txtcodigo & "' and id_sucursal = '" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
        '        If Conexion.SendHost(mstrSql, adoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        '            If Not (adoPaso.EOF = True And adoPaso.BOF = True) Then
        '                adoPaso.MoveFirst
        '            End If
        '            Do Until adoPaso.EOF
        '                Set Item = frmordencompra.lsvdetalle.ListItems.Add(, , ValorNulo(adoPaso.Fields!id_Linea))
        '                Item.SubItems(1) = Mid(ValorNulo(adoPaso.Fields!Id_Item), InStr(1, ValorNulo(adoPaso.Fields!Id_Item), "°") + 1, Len(ValorNulo(adoPaso.Fields!Id_Item)))
        '                Item.SubItems(2) = Retorna_Valor_General("Select Descripcion From Stck_Item Where Id_Item = '" & ValorNulo(adoPaso.Fields!Id_Item) & "'")
        '                Item.SubItems(3) = ValorNulo(adoPaso.Fields!cantidad)
        '                Item.SubItems(4) = FormatoValor(ValorNulo(adoPaso.Fields!Precio_Unitario), Gsigla, 0)
         '               Item.SubItems(5) = ValorNulo(adoPaso.Fields!descto_recgo)
         '               Item.SubItems(6) = FormatoValor(ValorNulo(adoPaso.Fields!subtotal), Gsigla, 0)
        '                Item.SubItems(7) = ValorNulo(adoPaso.Fields!fecha_entrega)
        '                Item.SubItems(8) = Mid(ValorNulo(adoPaso.Fields!Id_Item), 1, InStr(1, ValorNulo(adoPaso.Fields!Id_Item), "°") - 1)
        '                Item.SubItems(9) = ValorNulo(adoPaso.Fields!Id_Item)
        '                adoPaso.MoveNext
        '            Loop
        '        If CloseHost(adoPaso) = apOk Then
        '        End If
        '        End If
        
        
        
        '***********************************************************************************************
                'Ahora recalcula el subtotal del detalle
                frmordencompra.txtsubtotal = FormatoValor(0, Gsigla, 0)
                For i = 1 To Me.lsvdetalle.ListItems.Count
                    Me.txtsubtotal = FormatoValor((Round(Val(SacarFormatoValor(Me.txtsubtotal, Gsigla)) + Val(SacarFormatoValor(Me.lsvdetalle.ListItems(i).ListSubItems(6).Text, Gsigla)))), Gsigla, 0)
                Next i
                
        '***********************************************************************************************
                'Ahora recalcula el neto, iva, total
                  frmordencompra.txtneto = FormatoValor(Val(SacarFormatoValor(frmordencompra.txtsubtotal, Gsigla)) - Val(frmordencompra.txtdescuento), Gsigla, 0)
                  frmordencompra.txtiva = FormatoValor(Round(((18 * Val(SacarFormatoValor(frmordencompra.txtneto, Gsigla)))) / 100), Gsigla, 0)
                  frmordencompra.txttotal = FormatoValor(Val(SacarFormatoValor(frmordencompra.txtneto, Gsigla)) + Val(SacarFormatoValor(frmordencompra.txtiva, Gsigla)), Gsigla, 0)

                
                
                'frmordencompra.txtiva = FormatoValor((Round(((Val(SacarFormatoValor(frmordencompra.txtsubtotal, Gsigla)) - Val(frmordencompra.txtdescuento)) * valoriva) / 100)), Gsigla, 0)
                'frmordencompra.txtneto = FormatoValor((Round(((Val(SacarFormatoValor(frmordencompra.txtsubtotal, Gsigla)) - Val(frmordencompra.txtdescuento))) - Val(SacarFormatoValor(frmordencompra.txtiva, Gsigla)))), Gsigla, 0)
                'frmordencompra.txttotal = FormatoValor((Round(Val(SacarFormatoValor(frmordencompra.txtsubtotal, Gsigla)) - Val(frmordencompra.txtdescuento))), Gsigla, 0)
                
        '***********************************************************************************************
                'Ahora debe acutalizar los siguietes datos en la tabla maestra
                'mstrSql = "UPDATE Stck_Orden_Compra SET  Subtotal = " & SacarFormatoValor(Me.txtsubtotal, Gsigla) _
                '& ", Descto = " & Me.txtdescuento & ", Neto = " _
                '& SacarFormatoValor(Me.txtneto, Gsigla) & ", pje_iva =" & valoriva & ", Iva = " & SacarFormatoValor(Me.txtiva, Gsigla) & ", Total = " _
                '& SacarFormatoValor(Me.txttotal, Gsigla) & ",usr_id='" & USRID & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") _
                '& " " & Format(Time, "HH:MM:SS") & "'" _
                '& " Where Id_Orden_Compra='" & Me.txtcodigo & "' and id_sucursal = '" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
                'If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                '    mblnTablaVacia = False
                '    ActivaBotones
                '    Me.Tag = ""
                'End If
            End If
        strnrolista = ""
        gintlinea = Me.lsvdetalle.ListItems.Count + 1
        Else
            MsgBox "Debe seleccionar una linea del Detalle.", vbInformation, "Item"
        End If

Else
    MsgBox "Esta orden de compra no puede ser modificada, por haber sido recepcionados uno o más articulos", vbInformation, "Orden de Compra"
End If

  
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Set Adorecordencargado = New ADODB.Recordset
Set adorecordsucursal = New ADODB.Recordset
Set Adorecordmoneda = New ADODB.Recordset
mblnSW = True
          
'Recupera el valor IVA
valoriva = Retorna_Valor_General_Dinamico("Select Pje_Iva_Compras From Cont_Parametros Where Id_Empresa='" & IDEMPRESA & "'")
Me.Label1 = Me.Label1 + " (" + Str(valoriva) + "%)"
'Llena el combo Encargado
 mstrSql = "Select Id_Empleado, Nombre From Remu_Empleado Order By Nombre"
 If Conexion.SendHost(mstrSql, Adorecordencargado, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    Set Me.adoencargado.Recordset = Adorecordencargado
 End If

'Llena el combo Sucursal
 mstrSql = "Select Id_Sucursal, Descripcion From Glbl_Sucursal Where Id_Empresa ='" + IDEMPRESA + "' Order By Descripcion"
 If Conexion.SendHost(mstrSql, adorecordsucursal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    Set Me.Adosucursal.Recordset = adorecordsucursal
 End If

'Llena el combo Moneda
 mstrSql = "Select Id_Moneda, Descripcion From Glbl_Moneda Order By Descripcion"
 If Conexion.SendHost(mstrSql, Adorecordmoneda, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    Set Me.Adomoneda.Recordset = Adorecordmoneda
 End If

Me.lblfecha.Value = Date
End Sub



Private Sub lsvdetalle_ItemClick(ByVal item As MSComctlLib.ListItem)
 'Esta selecciona el Numero de Linea
  strnrolista = Me.lsvdetalle.ListItems(Me.lsvdetalle.SelectedItem.Index)
  
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
            Imprimirinforme
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
        If Not Atributos("Stck_20_0010", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If

        If gapAccion = apcrear Then
           AgregarRegistro
           txtcodigo = gstrBusca
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost adoPrincipal
            End If
            txtcodigo.Enabled = False
            Me.SetFocus
        End If
        If gapAccion = apninguno Then
           Renovar
        End If
    End If
    gapAccion = apninguno
    mblnSW = False
  '  txtnombre.SetFocus
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
                SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            CancelarAgregaRegistro
        Case 14 And tlbBarraHerramientas.Buttons.item("Crear").Enabled
            KeyAscii = 0
            AgregarRegistro
        Case 7 And tlbBarraHerramientas.Buttons.item("Grabar").Enabled
            KeyAscii = 0
            GrabarRegistro
        Case 4 And tlbBarraHerramientas.Buttons.item("Borrar").Enabled = False
            KeyAscii = 0
            BorrarRegistro
        Case 2 And tlbBarraHerramientas.Buttons.item("Buscar").Enabled
            KeyAscii = 0
            BuscarRegistro
        Case 9 And tlbBarraHerramientas.Buttons.item("Imprimir").Enabled
            KeyAscii = 0
            Imprimirinforme
        Case 16 And tlbBarraHerramientas.Buttons.item("Primero").Enabled
            KeyAscii = 0
            PrimerRegistro
        Case 1 And tlbBarraHerramientas.Buttons.item("Anterior").Enabled
            KeyAscii = 0
            RegistroAnterior
        Case 19 And tlbBarraHerramientas.Buttons.item("Siguiente").Enabled
            KeyAscii = 0
            RegistroSiguiente
        Case 21 And tlbBarraHerramientas.Buttons.item("Ultimo").Enabled
            KeyAscii = 0
            UltimoRegistro
        Case 18 And tlbBarraHerramientas.Buttons.item("Renovar").Enabled
            KeyAscii = 0
            Renovar
        Case 3 And tlbBarraHerramientas.Buttons.item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    limpiacampos
    ValoresporDefecto
    gintlinea = 1  'Deje en 1 la linea del detalle para la próxima vez
    Me.lblfecha.Value = Date
    Me.txtnombreproveedor.SetFocus
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtcodigo & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtcodigo & "' order by " & mcCampoCodigo
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    limpiacampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost adoPrincipal
    Me.txtnombreproveedor.SetFocus
End Sub
Private Sub GrabarRegistro()
Dim mstrSql As String
Dim i As Integer
'Primero debe ver si hay una recepcion para esta orden de compra
Dim strsql As String
strsql = "SELECT Count(Id_Orden_compra) From Stck_Orden_Compra_Recepcion_Detalle Where  Id_Orden_Compra = '" & Me.txtcodigo & "' and  Id_Sucursal ='" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
If Retorna_Valor_General_Tipo2(strsql) < 1 Then    'Ya no se puede modificar por que existe un registro relacionado en la tabla recepcion
            If Not validacion() Then
                Exit Sub
            End If
            'Ahora actualiza los totales
            'Me.txttotal = FormatoValor(Round(Val(SacarFormatoValor(Me.txtsubtotal, Gsigla)) - Val(Me.txtdescuento)), Gsigla, 0)
            'Me.txtiva = FormatoValor(Round(((Val(SacarFormatoValor(Me.txtsubtotal, Gsigla)) - Val(Me.txtdescuento)) * valoriva) / 100), Gsigla, 0)
            'Me.txtneto = FormatoValor(Round((Val(SacarFormatoValor(Me.txtsubtotal, Gsigla)) - Val(Me.txtdescuento)) - Val(Me.txtiva)), Gsigla, 0)
            frmordencompra.txtneto = FormatoValor(Val(SacarFormatoValor(frmordencompra.txtsubtotal, Gsigla)) - Val(frmordencompra.txtdescuento), Gsigla, 0)
            frmordencompra.txtiva = FormatoValor(Round(((18 * Val(SacarFormatoValor(frmordencompra.txtneto, Gsigla)))) / 100), Gsigla, 0)
            frmordencompra.txttotal = FormatoValor(Val(SacarFormatoValor(frmordencompra.txtneto, Gsigla)) + Val(SacarFormatoValor(frmordencompra.txtiva, Gsigla)), Gsigla, 0)
            
            If Me.txtdescuento = "" Then
                Me.txtdescuento = "0"
            End If
        
        
        
        
            If Me.Tag = "Crear" Then
            'Genera el numero de la orden
            Me.txtcodigo = Val(Retorna_Valor_General_Tipo2("SELECT Cast(Max(Cast(Id_Orden_Compra as integer)) as varchar) From Stck_Orden_Compra Where Id_Sucursal ='" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'")) + 1
            'Fin genera el numero
                
                mstrSql = "INSERT INTO Stck_Orden_Compra (Id_Orden_Compra, Id_Sucursal, Id_Empresa, Id_cliente_Proveedor, Id_Empleado, Id_Moneda, Tipo_Cambio, Fecha, Observaciones, Subtotal, Descto, Neto, Pje_Iva, Iva, Total, Usr_Id, Usr_Fecha) Values ('" _
                & Me.txtcodigo & "', '" & Me.cmbsucursal.BoundText & "', '" & IDEMPRESA _
                & "', '" & Me.txtproveedor & "','" & Me.cmbencargado.BoundText & "','" & Me.cmbmoneda.BoundText & "'," & Me.txttipocambio & ",'" & Me.lblfecha & "','" & Me.txtobservaciones & "'," & SacarFormatoValor(Me.txtsubtotal, Gsigla) & "," & Me.txtdescuento & "," & SacarFormatoValor(Me.txtneto, Gsigla) & "," & valoriva & "," & SacarFormatoValor(Me.txtiva, Gsigla) & "," & SacarFormatoValor(Me.txttotal, Gsigla) & ",'" & USRID & "','" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "')"
                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                    mblnTablaVacia = False
                    ActivaBotones
                    Me.Tag = ""
                End If
                'Ahora llena la tabla intermedia
                For i = 1 To Me.lsvdetalle.ListItems.Count
                    mstrSql = "INSERT INTO Stck_Orden_Compra_Detalle (Id_Sucursal, Id_Empresa, Id_Orden_compra, Id_Linea, Id_Item, Cantidad, Precio_Unitario, Descto_Recgo, Subtotal, Fecha_Entrega, Usr_Id, Usr_Fecha) VALUES ('"
                    mstrSql = mstrSql & Me.cmbsucursal.BoundText & "','" & IDEMPRESA & "','" & Me.txtcodigo & "'," & i & ",'" & Me.lsvdetalle.ListItems(i).ListSubItems(9) & "'," & Me.lsvdetalle.ListItems(i).ListSubItems(3)
                    mstrSql = mstrSql & "," & SacarFormatoValor(Me.lsvdetalle.ListItems(i).ListSubItems(4), Gsigla) & "," & Me.lsvdetalle.ListItems(i).ListSubItems(5) & "," & SacarFormatoValor(Me.lsvdetalle.ListItems(i).ListSubItems(6), Gsigla) & ",'" & Format(Me.lsvdetalle.ListItems(i).ListSubItems(7), "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "','" & USRID & "','" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "')"
                    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                        mblnTablaVacia = False
                        ActivaBotones
                        Me.Tag = ""
                    End If
                Next i
                
            Else
                'Actualiza la tabla maestra
                mstrSql = "UPDATE Stck_Orden_Compra SET Id_Cliente_Proveedor='" & Me.txtproveedor & "', Id_Empleado = '" _
                & Me.cmbencargado.BoundText & "', Id_Moneda = '" & Me.cmbmoneda.BoundText _
                & "', Tipo_Cambio = " & Me.txttipocambio & ", Fecha = '" & Format(Me.lblfecha, "DD/MM/YYYY") _
                & " " & Format(Time, "HH:MM:SS") & "', Observaciones = '" & Me.txtobservaciones _
                & "', Subtotal = " & SacarFormatoValor(Me.txtsubtotal, Gsigla) & ", Descto = " & Me.txtdescuento & ", Neto = " _
                & SacarFormatoValor(Me.txtneto, Gsigla) & ", pje_iva =" & valoriva & ", Iva = " & SacarFormatoValor(Me.txtiva, Gsigla) & ", Total = " _
                & SacarFormatoValor(Me.txttotal, Gsigla) & ",usr_id='" & USRID & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "'" _
                & " Where Id_Orden_Compra='" & Me.txtcodigo & "' and Id_Sucursal='" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                    mblnTablaVacia = False
                    ActivaBotones
                    Me.Tag = ""
                End If
                'Ahora actualiza la tabla del medio, primero borrar el contenido para actualizacion
                mstrSql = "Delete  From Stck_Orden_Compra_Detalle Where Id_Orden_Compra = '" & Me.txtcodigo & "' and Id_Sucursal='" & Me.cmbsucursal.BoundText & "' and Id_Empresa = '" & IDEMPRESA & "'"
                 If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                    mblnTablaVacia = False
                    ActivaBotones
                    Me.Tag = ""
                End If
                'Ahora llena la tabla intermedia
                For i = 1 To Me.lsvdetalle.ListItems.Count
                    mstrSql = "INSERT INTO Stck_Orden_Compra_Detalle (Id_Sucursal, Id_Empresa, Id_Orden_compra, Id_Linea, Id_Item, Cantidad, Precio_Unitario, Descto_Recgo, Subtotal, Fecha_Entrega, Usr_Id, Usr_Fecha) VALUES ('" & Me.cmbsucursal.BoundText & "','" & IDEMPRESA & "','"
                    mstrSql = mstrSql & Me.txtcodigo & "'," & i & ",'" & Me.lsvdetalle.ListItems(i).ListSubItems(9) & "'," & Me.lsvdetalle.ListItems(i).ListSubItems(3)
                    mstrSql = mstrSql & "," & SacarFormatoValor(Me.lsvdetalle.ListItems(i).ListSubItems(4), Gsigla) & "," & Me.lsvdetalle.ListItems(i).ListSubItems(5) & "," & SacarFormatoValor(Me.lsvdetalle.ListItems(i).ListSubItems(6), Gsigla) & ",'" & Format(Me.lsvdetalle.ListItems(i).ListSubItems(7), "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "','" & USRID & "','" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "')"
                    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                        mblnTablaVacia = False
                        ActivaBotones
                        Me.Tag = ""
                    End If
                Next i
                
            End If
Else
    MsgBox "Esta orden de compra no puede ser modificada, por haber sido recepcionados uno o más articulos", vbInformation, "Orden de Compra"
End If
            
End Sub
Private Sub BorrarRegistro()
    MsgBox "Esta opción no está disponible.", vbInformation, "Advertencia"
End Sub
Private Sub BuscarRegistro()
'    gstrBusca = Form1.BuscarRegistros(Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption)
    gstrBusca = ""
    gstrsucursal = ""
    'gstrempresa = ""
    frmbuscaordencompra.Show 1
    If gstrBusca <> "" And gstrsucursal <> "" Then
        mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' and ID_empresa = '" & IDEMPRESA & "' and Id_sucursal = '" & gstrsucursal & "' order by " & mcCampoCodigo
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost adoPrincipal
   End If
    Me.SetFocus
End Sub
Private Sub Imprimirinforme()
Me.CrystalReport1.ReportFileName = ReporteRuta + "OrdenCompra.rpt" 'Deja el nombre de la ruta del Reporte
Me.CrystalReport1.ParameterFields(0) = "Id_Orden;" & Me.txtcodigo & " ;TRUE"
Me.CrystalReport1.ParameterFields(1) = "Id_Empresa;" & IDEMPRESA & " ;TRUE"
Me.CrystalReport1.ParameterFields(2) = "Id_Sucursal;" & Me.cmbsucursal.BoundText & " ;TRUE"
Me.CrystalReport1.ParameterFields(3) = "Caracter;" & Guion & " ;TRUE"
Me.CrystalReport1.ParameterFields(4) = "sigla;" & Gsigla & " ;TRUE"
Me.CrystalReport1.ParameterFields(5) = "ObsLin1;" & Mid(Me.txtobservaciones, 1, 79) & " ;TRUE"
Me.CrystalReport1.ParameterFields(6) = "ObsLin2;" & Mid(Me.txtobservaciones, 80, 159) & " ;TRUE"

Me.CrystalReport1.Action = True
End Sub
Private Sub PrimerRegistro()
    Set Form1 = New APFORM
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by (Id_Sucursal + Id_Orden_compra)  "
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
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE (Id_Sucursal+Id_Orden_Compra) <'" & Me.cmbsucursal.BoundText + Me.txtcodigo & "' order by (Id_Sucursal + Id_Orden_Compra)  DESC"
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

    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE (Id_Sucursal+Id_Orden_compra) >'" & Me.cmbsucursal.BoundText + Me.txtcodigo & "' order by (Id_sucursal + Id_Orden_Compra)"
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
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by (Id_Sucursal + Id_Orden_compra) DESC"
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
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by (Id_Sucursal + Id_Orden_compra)"
    
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
    txtcodigo.Enabled = False
    With tlbBarraHerramientas.Buttons
        .item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
        .item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoEditar, True, False))
        .item("Cancelar").Enabled = False
        .item("Borrar").Enabled = False
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
    txtcodigo.Enabled = True
    With tlbBarraHerramientas.Buttons
        .item("Crear").Enabled = False
        .item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
        .item("Cancelar").Enabled = True
        .item("Borrar").Enabled = False
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
        limpiacampos
        MsgBox "La tabla no contiene registros...", vbInformation, "Advertencia"
    End If
End Sub
Private Sub LeerCampos()
Dim item As ListItem
    If mblnTablaVacia Then
        limpiacampos
        Exit Sub
    End If

    With adoPrincipal
        Me.txtcodigo.Text = ValorNulo(.Fields!Id_Orden_Compra)
        Me.txtdescuento.Text = ValorNulo(.Fields!descto)
        Me.txtiva.Text = FormatoValor(ValorNulo(.Fields!IVA), Gsigla, 0)
        Me.txtneto.Text = FormatoValor(ValorNulo(.Fields!neto), Gsigla, 0)
        Me.txtobservaciones.Text = ValorNulo(!observaciones)
        Me.txtsubtotal.Text = FormatoValor(ValorNulo(!subtotal), Gsigla, 0)
        Me.txttipocambio.Text = ValorNulo(!tipo_cambio)
        Me.txttotal.Text = FormatoValor(ValorNulo(!TOTAL), Gsigla, 0)
        Me.lblfecha = ValorNulo(!Fecha)
        Me.cmbencargado = Retorna_Valor_General("Select Nombre From Remu_Empleado Where Id_Empleado = '" & ValorNulo(!Id_Empleado) & "'")
        Me.txtnombreproveedor = Retorna_Valor_General("Select Razon_Social From Glbl_Cliente_Proveedor Where Id_Cliente_Proveedor = '" & ValorNulo(!Id_cliente_Proveedor) & "'")
        Me.cmbmoneda = Retorna_Valor_General("Select Descripcion From Glbl_Moneda Where Id_Moneda = '" & ValorNulo(!Id_Moneda) & "'")
        Me.cmbsucursal = Retorna_Valor_General("Select Descripcion From Glbl_Sucursal Where Id_Sucursal = '" & ValorNulo(!Id_Sucursal) & "' and Id_Empresa = '" & IDEMPRESA & "'")
        Me.txtproveedor = ValorNulo(!Id_cliente_Proveedor)
        'Ahora muestra el detalle
        Me.lsvdetalle.ListItems.Clear
        mstrSql = "Select * From Stck_Orden_Compra_Detalle Where Id_Orden_Compra = '" & Me.txtcodigo & "' and Id_empresa='" & IDEMPRESA & "' and Id_sucursal = '" & Me.cmbsucursal.BoundText & "'"
        If Conexion.SendHost(mstrSql, Adopaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not (Adopaso.EOF = True And Adopaso.BOF = True) Then
                Adopaso.MoveFirst
            End If
            Do Until Adopaso.EOF
              'Ahora agreaga al detalle
                Set item = frmordencompra.lsvdetalle.ListItems.Add(, , ValorNulo(Adopaso.Fields!id_Linea))
                item.SubItems(1) = Mid(ValorNulo(Adopaso.Fields!id_item), InStr(1, ValorNulo(Adopaso.Fields!id_item), "°") + 1, Len(ValorNulo(Adopaso.Fields!id_item)))
                item.SubItems(2) = Retorna_Valor_General("Select Descripcion From Stck_Item Where Id_Item = '" & ValorNulo(Adopaso.Fields!id_item) & "'")
                item.SubItems(3) = ValorNulo(Adopaso.Fields!cantidad)
                item.SubItems(4) = FormatoValor(ValorNulo(Adopaso.Fields!Precio_Unitario), Gsigla, 0)
                item.SubItems(5) = ValorNulo(Adopaso.Fields!descto_recgo)
                item.SubItems(6) = FormatoValor(ValorNulo(Adopaso.Fields!subtotal), Gsigla, 0)
                item.SubItems(7) = ValorNulo(Adopaso.Fields!fecha_entrega)
                item.SubItems(8) = Mid(ValorNulo(Adopaso.Fields!id_item), 1, InStr(1, ValorNulo(Adopaso.Fields!id_item), "°") - 1)
                item.SubItems(9) = ValorNulo(Adopaso.Fields!id_item)
                Adopaso.MoveNext
            Loop
        End If
    End With
    If CloseHost(Adopaso) = apOk Then
    End If
        
        
End Sub
Private Sub limpiacampos()
    Me.txtcodigo.Text = ""
    Me.txtdescuento.Text = "0"
    Me.txtiva.Text = ""
    Me.txtneto.Text = ""
    Me.txtobservaciones.Text = ""
    Me.txtnombreproveedor.Text = ""
    Me.txtsubtotal.Text = ""
    Me.txttipocambio.Text = ""
    Me.txttotal.Text = ""
    Me.cmbencargado.Text = ""
    Me.cmbmoneda.Text = ""
    Me.cmbsucursal.Text = ""
    Me.lsvdetalle.ListItems.Clear
    Me.lblfecha = Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS")
    'txtnombre.Text = ""
End Sub
Private Sub ValoresporDefecto()
    With adoPrincipal
        'chkVigencia.Value = vbChecked
    End With
End Sub
Private Function validacion() As Boolean
    validacion = True
    'If txtCodigo = "" Then
    '    MsgBox "El código de Pieza debe contener un valor...", vbInformation, "Advertencia"
    '    txtCodigo.SetFocus
    '    Validacion = False
    '    Exit Function
    'End If
    If IsNumeric(Me.txtdescuento) = False Or Me.txtdescuento = "" Then
        MsgBox "El porcentaje de descuento debe contener un valor Numerico...", vbInformation, "Advertencia"
        txtdescuento.SetFocus
        validacion = False
        Exit Function
    Else
        If Me.txtdescuento = "" Then
            Me.txtdescuento = "0"
        End If
    End If
    If Me.txtiva = "" Or IsNumeric(Me.txtiva) = False Then
        MsgBox "El campo IVA debe contener un valor Numerico...", vbInformation, "Advertencia"
        'Me.txtiva.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.txtneto = "" Or IsNumeric(Me.txtneto) = False Then
        MsgBox "El Valor Neto debe contener un valor Numerico...", vbInformation, "Advertencia"
        txtneto.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.txtproveedor = "" Then
        MsgBox "El Proveedor debe tener algun valor...", vbInformation, "Advertencia"
        txtdescuento.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.txtsubtotal = "" Or IsNumeric(Me.txtsubtotal) = False Then
        MsgBox "El Subtotal debe contener un valor Numerico...", vbInformation, "Advertencia"
        txtsubtotal.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.txttipocambio = "" Or IsNumeric(Me.txttipocambio) = False Then
        MsgBox "El Tipo de Cambio debe contener un valor Numerico...", vbInformation, "Advertencia"
        txttipocambio.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.txttotal = "" Or IsNumeric(Me.txttotal) = False Then
        MsgBox "El Total debe contener un valor Numerico...", vbInformation, "Advertencia"
        txttotal.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.cmbencargado = "" Then
        MsgBox "El Encargado debe tener algun valor...", vbInformation, "Advertencia"
        Me.cmbencargado.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.cmbmoneda = "" Then
        MsgBox "El Tipo Moneda debe contener un valor...", vbInformation, "Advertencia"
        Me.cmbmoneda.SetFocus
        validacion = False
        Exit Function
    End If
    If Me.cmbsucursal = "" Then
        MsgBox "La Sucursal debe contener un valor...", vbInformation, "Advertencia"
        Me.cmbsucursal.SetFocus
        validacion = False
        Exit Function
    End If
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As ADODB.Recordset
        mstrSql = "Select " & mcCampoCodigo & ", " & mcCampoNombre & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtcodigo & "' and Id_Empresa='" & IDEMPRESA & "' and Id_Sucursal = '" & Me.cmbsucursal.BoundText & "'"
        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(adoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", adoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                validacion = False
                txtcodigo.SetFocus
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gstrBusca = txtcodigo.Text
    Set frmordencompra = Nothing
End Sub
Private Sub RevizaAtributos()
'    mblnAccesoCrear = rsUsuarios!OPC_AUXILIAR_CREAR
'    mblnAccesoEditar = rsUsuarios!OPC_AUXILIAR_EDITAR
'    mblnAccesoBorrar = rsUsuarios!OPC_AUXILIAR_BORRAR
'    mblnAccesoImprimir = rsUsuarios!OPC_AUXILIAR_IMPRIMIR

    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True

End Sub
Sub Ejecuta_Consulta_Ado(strsql)
    If Not Conexion.SendHost(strsql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    End If
End Sub
Sub Ejecuta_Consulta_Ado_Dinamica(strsql)
    If Not Conexion.SendHost(strsql, adoPrincipal, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    End If
End Sub

Function Retorna_Valor_General(strsql)
'Esta funcion me retorna una valor solicitado desde una consulta SQL
'a la tabla General
    Ejecuta_Consulta_Ado (strsql)
    If Not (adoPrincipal.EOF = True And adoPrincipal.BOF = True) Then
        Do Until adoPrincipal.EOF
            Retorna_Valor_General = ValorNulo(adoPrincipal.Fields(0))
            adoPrincipal.MoveNext
        Loop
    End If
    adoPrincipal.Close
End Function

Function Retorna_Valor_General_Dinamico(strsql)
'Esta funcion me retorna una valor solicitado desde una consulta SQL
'a la tabla General
    Ejecuta_Consulta_Ado_Dinamica (strsql)
    If Not (adoPrincipal.EOF = True And adoPrincipal.BOF = True) Then
        Do Until adoPrincipal.EOF
            Retorna_Valor_General_Dinamico = ValorNulo(adoPrincipal.Fields(0))
            adoPrincipal.MoveNext
        Loop
    End If
    adoPrincipal.Close
End Function

Function Retorna_Nro_Registro(strsql)
Dim lintcontador As Integer
'Esta funcion me retorna una valor solicitado desde una consulta SQL
'a la tabla General
    lintcontador = 0
    Ejecuta_Consulta_Ado (strsql)
    If Not (adoPrincipal.EOF = True And adoPrincipal.BOF = True) Then
        Do Until adoPrincipal.EOF
            lintcontador = lintcontador + 1
            adoPrincipal.MoveNext
        Loop
    End If
    adoPrincipal.Close
    Retorna_Nro_Registro = lintcontador
End Function

Private Sub txtdescuento_LostFocus()
'Ahora actualiza los totales
'Me.txttotal = FormatoValor(Round(Val(SacarFormatoValor(Me.txtsubtotal, Gsigla)) - Val(Me.txtdescuento)), Gsigla, 0)
'Me.txtiva = FormatoValor(Round(((Val(SacarFormatoValor(Me.txtsubtotal, Gsigla)) - Val(Me.txtdescuento)) * valoriva) / 100), Gsigla, 0)
'Me.txtneto = FormatoValor(Round((Val(SacarFormatoValor(Me.txtsubtotal, Gsigla)) - Val(Me.txtdescuento)) - Val(Me.txtiva)), Gsigla, 0)
'frmordencompra.txtsubtotal = FormatoValor((Val(SacarFormatoValor(frmordencompra.txtsubtotal, Gsigla)) + (Round((Val(frmordencompra.txtsubtotal)) + (Val(Me.txtpv) * Val(Me.txtcantidad)) + lintrecargodescuento))), Gsigla, 0)
frmordencompra.txtneto = FormatoValor((Val(SacarFormatoValor(frmordencompra.txtsubtotal, Gsigla)) - Val(frmordencompra.txtdescuento)), Gsigla, 0)
frmordencompra.txtiva = FormatoValor(Round(((18 * Val(SacarFormatoValor(frmordencompra.txtneto, Gsigla)))) / 100), Gsigla, 0)
frmordencompra.txttotal = FormatoValor(Val(SacarFormatoValor(frmordencompra.txtneto, Gsigla)) + Val(SacarFormatoValor(frmordencompra.txtiva, Gsigla)), Gsigla, 0)

If Me.txtdescuento = "" Then
    Me.txtdescuento = "0"
End If
End Sub


Function Retorna_Valor_General_Tipo2(strsql)
'Esta funcion me retorna una valor solicitado desde una consulta SQL
'a la tabla General
    Ejecuta_Consulta_Ado_Tipo2 (strsql)
    If Not (adoPrincipal.EOF = True And adoPrincipal.BOF = True) Then
        Do Until adoPrincipal.EOF
            Retorna_Valor_General_Tipo2 = ValorNulo(adoPrincipal.Fields(0))
            adoPrincipal.MoveNext
        Loop
    End If
    adoPrincipal.Close
End Function
Sub Ejecuta_Consulta_Ado_Tipo2(strsql)
    If Not Conexion.SendHost(strsql, adoPrincipal, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    End If
End Sub

