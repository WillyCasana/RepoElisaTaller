VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPresupuestoMantenciones 
   Caption         =   "Presupuesto Automático Mantenciones"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   Icon            =   "frmPresupuestoMantenciones.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdConsultaStock 
      Appearance      =   0  'Flat
      Caption         =   "Consulta Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones"
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   6720
      Width           =   11500
      Begin VB.TextBox txtCometario 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   240
         Width           =   11295
      End
   End
   Begin Crystal.CrystalReport rptPresMantencion 
      Left            =   1440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame fmeRepuesto 
      Caption         =   "Repuestos asociados a la Actividad"
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   11500
      Begin MSComctlLib.ListView lvwRepuestos 
         Height          =   1560
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   2752
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
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
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10019
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Familia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "IDFAM"
            Object.Width           =   18
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "completar lista"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Saldo"
            Object.Width           =   882
         EndProperty
      End
   End
   Begin VB.Frame fmeActividades 
      Caption         =   "Actividades asociadas al Servicio"
      Height          =   2085
      Left            =   45
      TabIndex        =   1
      Top             =   1680
      Width           =   11520
      Begin MSComctlLib.ListView lvwActividades 
         Height          =   1725
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   3043
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10019
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Nº Horas"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Especialidad"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Codigo Especialidad"
            Object.Width           =   18
         EndProperty
      End
   End
   Begin VB.Frame fmeServicios 
      Height          =   1320
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   11500
      Begin VB.TextBox txtTotalMantencion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txtInsumos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9990
         TabIndex        =   20
         Text            =   "0"
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox txtSubtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   120
         Width           =   1365
      End
      Begin VB.TextBox txtCodigoServicio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   855
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo dtcModelo 
         Bindings        =   "frmPresupuestoMantenciones.frx":179A
         Height          =   315
         Left            =   4050
         TabIndex        =   5
         Top             =   315
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcMarca 
         Bindings        =   "frmPresupuestoMantenciones.frx":17B3
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   320
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datModelos 
         Height          =   330
         Left            =   6525
         Top             =   315
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc datMarcas 
         Height          =   330
         Left            =   1320
         Top             =   315
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
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
      Begin MSComctlLib.Toolbar tlbOpciones 
         Height          =   330
         Index           =   3
         Left            =   3915
         TabIndex        =   12
         Top             =   855
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   582
         ButtonWidth     =   1138
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Busca Servicio"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   8520
         X2              =   8520
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label Label8 
         Caption         =   "Tot. Mantención"
         Height          =   240
         Left            =   8640
         TabIndex        =   24
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Insumos"
         Height          =   240
         Left            =   8685
         TabIndex        =   23
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label6 
         Caption         =   "M.O + Repuestos"
         Height          =   240
         Left            =   8640
         TabIndex        =   22
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label lblValorServicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6930
         TabIndex        =   18
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label lblHorasServicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5130
         TabIndex        =   17
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label5 
         Caption         =   "Valor"
         Height          =   285
         Left            =   6300
         TabIndex        =   16
         Top             =   900
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Horas"
         Height          =   240
         Left            =   4545
         TabIndex        =   15
         Top             =   900
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Servicio :"
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Modelo :"
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir "
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbTotalActividades 
      Height          =   405
      Left            =   6480
      TabIndex        =   13
      Top             =   3840
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Total Mano Obra"
            TextSave        =   "Total Mano Obra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   5292
            MinWidth        =   5292
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
   Begin MSComctlLib.StatusBar stbTotalRepuestos 
      Height          =   405
      Left            =   6480
      TabIndex        =   14
      Top             =   6240
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Total Repuestos"
            TextSave        =   "Total Repuestos"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   5292
            MinWidth        =   5292
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
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   3960
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":17CB
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":18DD
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":19EF
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":1B01
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":1C13
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":1D25
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":1E37
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":1F49
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":205B
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":216D
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":227F
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":2391
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":24A3
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":25B5
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":26C7
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":27D9
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":28EB
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":2D3D
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":318F
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":32A1
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":33FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":3559
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":36B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":3811
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":42DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":4731
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":4895
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":4CF1
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":4E4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":6159
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":66F5
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":6851
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":69AD
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":6D01
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":7055
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":73A9
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":76FD
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":7A51
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":7F95
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":80A7
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":83FB
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":874F
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":8863
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":8BB7
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":8CC9
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoMantenciones.frx":901D
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPresupuestoMantenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnSW As Boolean
Dim adoPrincipal As New ADODB.Recordset
Dim mstrSql As String
Dim lsiItemSelected As Boolean
Dim lsiItem As ListItem
Const mcintHeight As Integer = 7700
Const mcintWidth As Integer = 11700
Const mcstrMensaje As String = "Confirma Eliminar El Item Seleccionado desde "

Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Sub Repuestos_de_la_Actividad(strMarca As String, strModelo As String, strServicio As String, strActividad As String)
    
lvwRepuestos.ListItems.Clear
mstrSql = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
mstrSql = mstrSql & " Stck_Item.Descripcion AS NOMBRE, "
mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Valor AS VLR, "
mstrSql = mstrSql & " Stck_Item.Id_Familia AS IDFAM, "
mstrSql = mstrSql & " Stck_Item.Precio_Venta as Precio,"
mstrSql = mstrSql & " Glbl_Familia.Descripcion AS FAMILIA "
mstrSql = mstrSql & " FROM Glbl_Familia RIGHT OUTER JOIN Stck_Item ON  Glbl_Familia.Id_Familia = Stck_Item.Id_Familia RIGHT OUTER JOIN Tllr_Actividad_Repuesto ON Stck_Item.Id_Item = Tllr_Actividad_Repuesto.Id_Item"
mstrSql = mstrSql & " WHERE Tllr_Actividad_Repuesto.Id_Marca = '" & strMarca & "' AND Tllr_Actividad_Repuesto.Id_Modelo = '" & strModelo & "' AND Tllr_Actividad_Repuesto.Id_Servicio = '" & strServicio & "' AND Tllr_Actividad_Repuesto.Id_Actividad = '" & strActividad & "' "
    
    
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwRepuestos.ListItems.Add(, , !Codigo)
                    lsiItem.SubItems(1) = !Nombre
                    lsiItem.SubItems(2) = FormatoValor(!CANTY, "", 1)
                    lsiItem.SubItems(3) = Format(!Precio, "###,##0")
                    lsiItem.SubItems(4) = !Familia
                    lsiItem.SubItems(5) = !IDFAM
                    .MoveNext
                Wend
            End If
        End With
    End If
    
End Sub

Sub Actividades_del_Servicio(strMarca As String, strModelo As String, strServicio As String)

    mstrSql = " SELECT Tllr_Actividad_Servicio_Modelo.Id_Actividad AS CODIGO,"
    mstrSql = mstrSql & " Tllr_Actividad.Descripcion AS NOMBRE,"
    mstrSql = mstrSql & " Tllr_Actividad_Servicio_Modelo.Horas AS TIEMPO,"
    mstrSql = mstrSql & " Tllr_Actividad_Servicio_Modelo.Valor AS VALOR,"
    mstrSql = mstrSql & " Tllr_Actividad.Id_Especialidad AS IDESPE,"
    mstrSql = mstrSql & " Tllr_Especialidad.Descripcion AS ESPECIAL"
    mstrSql = mstrSql & " FROM Tllr_Actividad LEFT OUTER JOIN Tllr_Especialidad ON"
    mstrSql = mstrSql & " Tllr_Actividad.Id_Especialidad = Tllr_Especialidad.Id_Especialidad"
    mstrSql = mstrSql & " RIGHT OUTER JOIN Tllr_Actividad_Servicio_Modelo ON"
    mstrSql = mstrSql & " Tllr_Actividad.Id_Actividad = Tllr_Actividad_Servicio_Modelo.Id_Actividad"
    mstrSql = mstrSql & " WHERE Tllr_Actividad_Servicio_Modelo.Id_Marca = '" & strMarca & "' AND"
    mstrSql = mstrSql & " Tllr_Actividad_Servicio_Modelo.Id_Modelo = '" & strModelo & "' AND"
    mstrSql = mstrSql & " Tllr_Actividad_Servicio_Modelo.Id_Servicio = '" & strServicio & "' "

    lvwActividades.ListItems.Clear
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwActividades.ListItems.Add(, , !Codigo)
                    lsiItem.SubItems(1) = !Nombre
                    lsiItem.SubItems(2) = !TIEMPO
                    lsiItem.SubItems(3) = Format(!Valor, "###,###")
                    lsiItem.SubItems(4) = !ESPECIAL
                    lsiItem.SubItems(5) = !IDESPE
                    .MoveNext
                Wend
            End If
        End With
    End If

End Sub



'Sub Servicios_del_Modelo(strMarca As String, strModelo As String)
'
'    lvwServicios.ListItems.Clear
'
'    mstrsql = "SELECT Tllr_Servicio_Modelo.Id_Servicio AS CODIGO,"
'    mstrsql = mstrsql & " Tllr_Servicio.Descripcion AS NOMBRE, "
'    mstrsql = mstrsql & " Tllr_Servicio_Modelo.Horas AS HORAS,"
'    mstrsql = mstrsql & " Tllr_Servicio.Seccion AS OBJETO,"
'    mstrsql = mstrsql & " Tllr_Servicio_Modelo.Valor AS VALOR"
'    mstrsql = mstrsql & " FROM Tllr_Servicio RIGHT OUTER JOIN"
'    mstrsql = mstrsql & " Tllr_Servicio_Modelo ON"
'    mstrsql = mstrsql & " Tllr_Servicio.Id_Servicio = Tllr_Servicio_Modelo.Id_Servicio"
'    mstrsql = mstrsql & " WHERE Tllr_Servicio_Modelo.Id_Marca = '" & strMarca & "' AND"
'    mstrsql = mstrsql & " Tllr_Servicio_Modelo.Id_Modelo = '" & strModelo & "' "
'
'    If Conexion.SendHost(mstrsql, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
'        With AdoPrincipal
'            If Not .BOF And Not .EOF Then
'                .MoveFirst
'                While Not .EOF
'                    Set lsiItem = lvwServicios.ListItems.Add(, , !Codigo)
'                    lsiItem.SubItems(1) = !Nombre
'                    lsiItem.SubItems(2) = !HORAS
'                    lsiItem.SubItems(3) = Format(!Valor, "###,##0")
'                    lsiItem.SubItems(4) = IIf(!Objeto = "M", "MECANICA", "CARROCERIA")
'                    .MoveNext
'                Wend
'            End If
'        End With
'    End If
'
'End Sub

Sub FillMarcas()
    dtcMarca.Enabled = True
    mstrSql = "Select Id_marca as CODIGO, Descripcion as Nombre from Glbl_Marca where VIGENCIA = 'S' order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datMarcas
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcMarca.ListField = "Nombre"
                dtcMarca.BoundColumn = "Codigo"
                dtcMarca.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Sub FillModelos(strMarca As String)
    dtcModelo.Enabled = True
    mstrSql = "Select Id_modelo as CODIGO, id_Modelo + '////' + Descripcion as Nombre from Glbl_Modelo where VIGENCIA = 'S' and Id_marca = '" & strMarca & "'  order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datModelos
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcModelo.ListField = "Nombre"
                dtcModelo.BoundColumn = "Codigo"
                dtcModelo.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Private Sub cmdConsultaStock_Click()
    If Me.lvwRepuestos.ListItems.Count > 0 Then
        'Levanta listview con los repuestos del presupuesto
        frmRepuestosReservados.Show vbModal
    End If
End Sub

Private Sub dtcMarca_Change()
'lvwServicios.ListItems.Clear
lvwActividades.ListItems.Clear
lvwRepuestos.ListItems.Clear
If dtcMarca.BoundText <> "" Then
    dtcModelo.Text = ""
    FillModelos dtcMarca.BoundText
End If
End Sub

Private Sub dtcModelo_Change()
'lvwServicios.ListItems.Clear
'lvwActividades.ListItems.Clear
'lvwRepuestos.ListItems.Clear
'If dtcModelo.BoundText <> "" Then
'    fmeServicios.Enabled = True
'    Servicios_del_Modelo dtcMarca.BoundText, dtcModelo.BoundText
'    If lvwServicios.ListItems.Count > 0 Then
'        Actividades_del_Servicio dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem
'        If lvwActividades.ListItems.Count > 0 Then
'            Repuestos_de_la_Actividad dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem
'        End If
'    End If
'Else
'    fmeServicios.Enabled = False
'    Servicios_del_Modelo "", ""
'    Actividades_del_Servicio "", "", ""
'End If
End Sub

Private Sub Form_Activate()
If mblnSW Then
    If Not Atributos("Glbl", "Tllr_20_0050", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If
    FillMarcas
    mblnSW = False
End If

End Sub

Private Sub Form_Load()
mblnSW = True
gstrProcedencia = "Presupuesto Mantencion"
End Sub

Private Sub Form_Resize()
With Me
If .WindowState = 0 Then
    .Height = mcintHeight
    .Width = mcintWidth
    .Top = 0
    .Left = 0
End If
End With
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Imprimir"
            ImprimirInforme
        Case "Cerrar"
            Unload Me
    End Select
    Screen.MousePointer = vbDefault

End Sub

Private Sub tlbOpciones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    If Trim(Me.dtcMarca.Text) <> "" Then
        If Trim(Me.dtcModelo.Text) <> "" Then
            frmBuscaServicioMarcaModelo.Show 1
            If Me.lvwActividades.ListItems.Count > 0 Then
                stbTotalActividades.Panels(2) = FormatoValor(TotalSeccion(lvwActividades, 3, 1), "", gintDecimalesMoneda)
                stbTotalRepuestos.Panels(2) = FormatoValor(TotalSeccion(lvwRepuestos, 3, 2), "", gintDecimalesMoneda)
                TotalFinal
            End If
        Else
            MsgBox "Falta el Modelo", vbExclamation, "Parametros de Busqueda"
        End If
    Else
        MsgBox "Falta la Marca", vbExclamation, "Parametros de Busqueda"
    End If
End If

End Sub

Private Sub txtCodigoServicio_Change()
If txtCodigoServicio <> "" Then
    Actividades_del_Servicio dtcMarca.BoundText, dtcModelo.BoundText, txtCodigoServicio.Tag
    If lvwActividades.ListItems.Count > 0 Then
        Repuestos_de_la_Actividad dtcMarca.BoundText, dtcModelo.BoundText, txtCodigoServicio.Tag, lvwActividades.SelectedItem
    Else
        lvwRepuestos.ListItems.Clear
    End If
Else
    lvwActividades.ListItems.Clear
    lvwRepuestos.ListItems.Clear
End If

End Sub
Function TotalSeccion(lvwObjeto As ListView, IndiceSubItem As Integer, ItemServicio As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwObjeto
    If ItemServicio = 1 Then  '// suma actividades
        For intS = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intS)
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
        Next
    End If
    If ItemServicio = 2 Then  '// suma repuestos
        For intS = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intS)
            dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), "")) * CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(2) = "", 0, .SelectedItem.SubItems(2)), ""))
        Next
    End If
End With
TotalSeccion = dblPreSuma
End Function

Sub ImprimirInforme()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim tabla2 As DAO.Recordset
Dim tabla3 As DAO.Recordset
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
    
    If Me.lvwActividades.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(gstrPathReporte + "\Tllr_PresupMantenciones.mdb") = "" Then
        Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\Tllr_PresupMantenciones.mdb", dbLangGeneral) ' Crea a una base de datos nueva
        Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (MARCA text,MODELO text,SERVICIO text,HORAS text,VALOR text,REPUESTOS text,INSUMOS text,NETO text,IVA text,TOTAL text,Comentario Memo,USUARIORED TEXT, INSTANCIA TEXT)"
        Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE2 (REPUESTO text,CANTIDAD text,PRECIOREPUESTO text,FAMILIA text,SUBTOTAL double, SALDO text,USUARIORED TEXT, INSTANCIA TEXT)"
        Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE3 (ACTIVIDAD text,HORASACTIVIDAD text,VALORACTIVIDAD text,USUARIORED TEXT, INSTANCIA TEXT)"
    Else
       Set Dbsnueva = wrkPredeterminado.OpenDatabase(gstrPathReporte & "\Tllr_PresupMantenciones.mdb")
    End If
    
    Dbsnueva.Execute "Delete * from T_PARAREPORTE WHERE USUARIORED = '" & gstrIdUsuario & "' AND INSTANCIA = '" & frmMain.hwnd & "'"
    Dbsnueva.Execute "Delete * from T_PARAREPORTE2 WHERE USUARIORED = '" & gstrIdUsuario & "' AND INSTANCIA = '" & frmMain.hwnd & "'"
    Dbsnueva.Execute "Delete * from T_PARAREPORTE3 WHERE USUARIORED = '" & gstrIdUsuario & "' AND INSTANCIA = '" & frmMain.hwnd & "'"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")

    'Encabezado
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    Tabla.AddNew
    Tabla!Marca = dtcMarca.Text
    Tabla!Modelo = dtcModelo.Text
    Tabla!servicio = txtCodigoServicio
    Tabla!Horas = lblHorasServicio
    Tabla!Valor = lblValorServicio
    Tabla!Repuestos = stbTotalRepuestos.Panels(2)
    Tabla!Insumos = txtInsumos
    Tabla!Neto = txtTotalMantencion
    Tabla!IVA = FormatoValor(CDbl(txtTotalMantencion) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto), "", gintDecimalesMoneda)
    Tabla!Total = FormatoValor(CDbl(txtTotalMantencion) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto) + CDbl(txtTotalMantencion), "", gintDecimalesMoneda)
    Tabla!Comentario = txtCometario
    Tabla!USUARIORED = gstrIdUsuario
    Tabla!INSTANCIA = frmMain.hwnd 'Numero unico
    Tabla.Update
    Tabla.Close
    
    Set tabla2 = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE2")
    'Repuestos
    For i = 1 To Me.lvwRepuestos.ListItems.Count
        Set lvwRepuestos.SelectedItem = lvwRepuestos.ListItems(i)
        tabla2.AddNew
        tabla2!Repuesto = IIf(Me.lvwRepuestos.SelectedItem.SubItems(1) = "", " ", Me.lvwRepuestos.SelectedItem.SubItems(1))
        tabla2!cantidad = IIf(Me.lvwRepuestos.SelectedItem.SubItems(2) = "", " ", Me.lvwRepuestos.SelectedItem.SubItems(2))
        tabla2!PrecioRepuesto = IIf(Me.lvwRepuestos.SelectedItem.SubItems(3) = "", " ", Me.lvwRepuestos.SelectedItem.SubItems(3))
        tabla2!Familia = IIf(Me.lvwRepuestos.SelectedItem.SubItems(5) = "", " ", Me.lvwRepuestos.SelectedItem.SubItems(5))
        tabla2!SubTotal = CDbl(lvwRepuestos.SelectedItem.SubItems(2)) * CDbl(lvwRepuestos.SelectedItem.SubItems(3)) 'Me.stbTotalRepuestos.Panels(2)
        tabla2!Saldo = Me.lvwRepuestos.SelectedItem.SubItems(7)
        tabla2!USUARIORED = gstrIdUsuario
        tabla2!INSTANCIA = frmMain.hwnd 'Numero unico
        tabla2.Update
    Next i
    tabla2.Close
    
    'Actividades
    
    Set tabla3 = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE3")
    For i = 1 To Me.lvwActividades.ListItems.Count
        Set lvwActividades.SelectedItem = lvwActividades.ListItems(i)
        tabla3.AddNew
        tabla3!actividad = IIf(Me.lvwActividades.SelectedItem.SubItems(1) = "", " ", Me.lvwActividades.SelectedItem.SubItems(1))
        tabla3!HorasActividad = IIf(Me.lvwActividades.SelectedItem.SubItems(2) = "", " ", Me.lvwActividades.SelectedItem.SubItems(2))
        tabla3!ValorActividad = IIf(Me.lvwActividades.SelectedItem.SubItems(3) = "", " ", Me.lvwActividades.SelectedItem.SubItems(3))
        tabla3!USUARIORED = gstrIdUsuario
        tabla3!INSTANCIA = frmMain.hwnd 'Numero unico
        tabla3.Update
    Next i
    tabla3.Close
   
   With rptPresMantencion
        .ReportFileName = gstrPathReporte & "\PresupuestoMantenciones.rpt"
        .WindowTitle = "Reporte de Presupuesto de Mantenciones"
        '.DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='PRESUPUESTOS DE MANTENCION'"
        .Formulas(2) = "Empresa='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "NombreIgv='" & gstrNombreIva & "'"
        .Formulas(6) = "Instancia='" & frmMain.hwnd & "'"
        .Formulas(7) = "UsuarioInforme='" & gstrIdUsuario & "'"
        .Formulas(8) = "FamiliaInsumos='" & gstrCodigoInsumos & "'"
        .Formulas(9) = "FamiliaLubricantes='" & gstrCodigoLubricantes & "'"
        .Formulas(10) = "FamiliaMateriales='" & gstrCodigoMateriales & "'"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = True
   End With
   
   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub

Function TotalFinal()
    Me.txtSubtotal = FormatoValor(CDbl(IIf(Me.lblValorServicio = "", 0, Me.lblValorServicio)) + CDbl(Me.stbTotalRepuestos.Panels(2)), "", gintDecimalesMoneda)
    If gcurMaterialesMO <> 0 Then
        Me.txtInsumos = FormatoValor((CDbl(Me.lblValorServicio) * gcurMaterialesMO) / 100, "", gintDecimalesMoneda)
    Else
        Me.txtInsumos = FormatoValor(gcurInsumo, "", gintDecimalesMoneda)
    End If
    Me.txtTotalMantencion = FormatoValor(CDbl(Me.txtSubtotal) + CDbl(Me.txtInsumos), "", gintDecimalesMoneda)
End Function

Private Sub txtInsumos_GotFocus()
txtInsumos = CDbl(txtInsumos)
MarcaTexto txtInsumos
End Sub

Private Sub txtInsumos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{tab}"
End If
End Sub

Private Sub txtInsumos_LostFocus()
    txtInsumos = FormatoValor(txtInsumos, "", gintDecimalesMoneda)
    txtTotalMantencion = FormatoValor(CDbl(txtSubtotal) + CDbl(txtInsumos), "", gintDecimalesMoneda)
End Sub
