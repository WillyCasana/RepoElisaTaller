VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTempServiciosMarMod 
   Caption         =   "Temparios  por Modelo"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19740
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTempServiciosMarMod.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   19740
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeRepuesto 
      Caption         =   "Repuestos asociados a la Actividad"
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   11535
      Begin MSComctlLib.ListView lvwRepuestos 
         Height          =   2775
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4895
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Text            =   "Cantidad"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   3528
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
      End
      Begin MSComctlLib.Toolbar tlbOpciones 
         Height          =   660
         Index           =   2
         Left            =   10440
         TabIndex        =   12
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1164
         ButtonWidth     =   1746
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "Agregar"
               Object.ToolTipText     =   "Agrega Servicio Nuevo"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "Quitar"
               Object.ToolTipText     =   "Quitar Servicio"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Copiar"
               Key             =   "Copiar"
               ImageIndex      =   22
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fmeActividades 
      Caption         =   "Actividades asociadas al Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   5
      Top             =   7560
      Width           =   11520
      Begin MSComctlLib.ListView lvwActividades 
         Height          =   2085
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   3678
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
      Begin MSComctlLib.Toolbar tlbOpciones 
         Height          =   660
         Index           =   1
         Left            =   10470
         TabIndex        =   11
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1164
         ButtonWidth     =   1746
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "Agregar"
               Object.ToolTipText     =   "Agrega Servicio Nuevo"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "Quitar"
               Object.ToolTipText     =   "Quitar Servicio"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Copiar"
               Key             =   "Copiar"
               ImageIndex      =   22
            EndProperty
         EndProperty
      End
   End
   Begin MSDataListLib.DataCombo dtcModelo 
      Bindings        =   "frmTempServiciosMarMod.frx":038A
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcMarca 
      Bindings        =   "frmTempServiciosMarMod.frx":03A3
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fmeServicios 
      Caption         =   "Servicios asociados al Modelo"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11535
      Begin MSComctlLib.Toolbar tlbOpciones 
         Height          =   990
         Index           =   0
         Left            =   10440
         TabIndex        =   10
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1746
         ButtonWidth     =   1746
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "Agregar"
               Object.ToolTipText     =   "Agrega Servicio Nuevo"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "Quitar"
               Object.ToolTipText     =   "Quitar Servicio"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copiar"
               Key             =   "Copiar"
               ImageIndex      =   22
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwServicios 
         Height          =   2775
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4895
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Codigo"
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Des"
            Text            =   "Descripción"
            Object.Width           =   10019
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "NroHoras"
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
            Text            =   "Mecanica/Carroceria"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   10800
         Top             =   1920
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
               Picture         =   "frmTempServiciosMarMod.frx":03BB
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":04CD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":0925
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":0D7D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":11D5
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":12E7
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":13F9
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":150B
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":161D
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":172F
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1841
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1953
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1A65
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1B77
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1C89
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1D9B
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1EAD
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":1FBF
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":20D1
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":21E3
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":2635
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTempServiciosMarMod.frx":2A87
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc datModelos 
      Height          =   330
      Left            =   9120
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   2595
      Top             =   15
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Label Label2 
      Caption         =   "Modelo :"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Marca :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   160
      Width           =   735
   End
End
Attribute VB_Name = "frmTempServiciosMarMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnSW As Boolean
Dim AdoPrincipal As New ADODB.Recordset
Dim mstrSQL As String
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
mstrSQL = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
mstrSQL = mstrSQL & " Stck_Item.Descripcion AS NOMBRE, "
mstrSQL = mstrSQL & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
mstrSQL = mstrSQL & " Tllr_Actividad_Repuesto.Valor AS VLR, "
mstrSQL = mstrSQL & " Stck_Item.Precio_Venta as Precio,"
mstrSQL = mstrSQL & " Stck_Item.Id_Familia AS IDFAM, "
mstrSQL = mstrSQL & " Glbl_Familia.Descripcion AS FAMILIA "
mstrSQL = mstrSQL & " FROM Glbl_Familia RIGHT OUTER JOIN Stck_Item ON  Glbl_Familia.Id_Familia = Stck_Item.Id_Familia RIGHT OUTER JOIN Tllr_Actividad_Repuesto ON Stck_Item.Id_Item = Tllr_Actividad_Repuesto.Id_Item"
mstrSQL = mstrSQL & " WHERE Tllr_Actividad_Repuesto.Id_Marca = '" & strMarca & "' AND Tllr_Actividad_Repuesto.Id_Modelo = '" & strModelo & "' AND Tllr_Actividad_Repuesto.Id_Servicio = '" & strServicio & "' AND Tllr_Actividad_Repuesto.Id_Actividad = '" & strActividad & "' "
    
    
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwRepuestos.ListItems.Add(, , !Codigo)
                    lsiItem.SubItems(1) = !Nombre
                    lsiItem.SubItems(2) = Format(!CANTY, "###,##0.0")
                    lsiItem.SubItems(3) = Format(!Precio, "###,##0")
                    lsiItem.SubItems(4) = !Familia
                    lsiItem.SubItems(5) = !IDFAM
                    .MoveNext
                Wend
            End If
        End With
    End If
    
End Sub

Sub EliminarItem(intTipo As Integer, strMarca As String, strModelo As String, Optional strServicio As String, Optional strActividad As String, Optional strRepuesto As String)
Dim strSql As String
Select Case intTipo
Case 0 '////////elimina servicio
'    If MsgBox(mcstrMensaje & "Servicios por Modelos", 4 + 32) = vbYes Then
        strSql = "SELECT COUNT(*) AS CUANTOS FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
        If Conexion.SendHost(strSql, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
            With AdoPrincipal
                .MoveFirst
                If !CUANTOS > 0 Then
                    '////////////////// TIENE ACTIVIDADES RELACIONADAS
                    MsgBox "TIENE ACTIVIDADES RELACIONADAS"
                    '//////////////////
                    strSql = "DELETE FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSql, , , , gcTiempoEspera
                    '//////////////////
                    strSql = "DELETE FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSql, , , , gcTiempoEspera
                    '//////////////////
                    strSql = "DELETE FROM Tllr_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSql, , , , gcTiempoEspera
                    lvwServicios.ListItems.Remove lvwServicios.SelectedItem.Index
                Else
                    '////////////////// NO TIENE ACTIVIDADES RELACIONADAS
                    strSql = "DELETE FROM Tllr_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSql, , , , gcTiempoEspera
                    lvwServicios.ListItems.Remove lvwServicios.SelectedItem.Index
                End If
            End With
        End If
    'End If
Case 1 '/////////elimina actividad
    If MsgBox(mcstrMensaje & "Actividades por Servicio", 4 + 32) = vbYes Then
        strSql = "SELECT count(*) AS CUANTOS FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
        If Conexion.SendHost(strSql, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
            With AdoPrincipal
                .MoveFirst
                If !CUANTOS > 0 Then '////////////////// TIENE ACTIVIDADES RELACIONADAS
                    MsgBox "TIENE REPUESTOS RELACIONADOS"
                    '//////////////////
                    strSql = "DELETE FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
                    Conexion.SendHost strSql, , , , gcTiempoEspera
                    '//////////////////
                    strSql = "DELETE FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
                    Conexion.SendHost strSql, , , , gcTiempoEspera
                    lvwActividades.ListItems.Remove lvwActividades.SelectedItem.Index
                    lvwRepuestos.ListItems.Clear
                Else '////////////////// NO TIENE ACTIVIDADES RELACIONADAS
                    strSql = "DELETE FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
                    Conexion.SendHost strSql, , , , gcTiempoEspera
                    lvwActividades.ListItems.Remove lvwActividades.SelectedItem.Index
                    
                End If
            End With
        End If
        
    End If
Case 2 '/////////elimina repuesto
    If MsgBox(mcstrMensaje & "Repuestos de la Actividad", 4 + 32) = vbYes Then
        strSql = "DELETE FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' AND Id_Item = '" & strRepuesto & "'"
        Conexion.SendHost strSql, , , , gcTiempoEspera
        lvwRepuestos.ListItems.Remove lvwRepuestos.SelectedItem.Index
'        If lvwActividades.ListItems.Count > 0 Then
'            Repuestos_de_la_Actividad dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem
'        End If
    End If
End Select

End Sub

Sub Actividades_del_Servicio(strMarca As String, strModelo As String, strServicio As String)

    mstrSQL = " SELECT Tllr_Actividad_Servicio_Modelo.Id_Actividad AS CODIGO,"
    mstrSQL = mstrSQL & " Tllr_Actividad.Descripcion AS NOMBRE,"
    mstrSQL = mstrSQL & " Tllr_Actividad_Servicio_Modelo.Horas AS TIEMPO,"
    mstrSQL = mstrSQL & " Tllr_Actividad_Servicio_Modelo.Valor AS VALOR,"
    mstrSQL = mstrSQL & " Tllr_Actividad.Id_Especialidad AS IDESPE,"
    mstrSQL = mstrSQL & " Tllr_Especialidad.Descripcion AS ESPECIAL"
    mstrSQL = mstrSQL & " FROM Tllr_Actividad LEFT OUTER JOIN Tllr_Especialidad ON"
    mstrSQL = mstrSQL & " Tllr_Actividad.Id_Especialidad = Tllr_Especialidad.Id_Especialidad"
    mstrSQL = mstrSQL & " RIGHT OUTER JOIN Tllr_Actividad_Servicio_Modelo ON"
    mstrSQL = mstrSQL & " Tllr_Actividad.Id_Actividad = Tllr_Actividad_Servicio_Modelo.Id_Actividad"
    mstrSQL = mstrSQL & " WHERE Tllr_Actividad_Servicio_Modelo.Id_Marca = '" & strMarca & "' AND"
    mstrSQL = mstrSQL & " Tllr_Actividad_Servicio_Modelo.Id_Modelo = '" & strModelo & "' AND"
    mstrSQL = mstrSQL & " Tllr_Actividad_Servicio_Modelo.Id_Servicio = '" & strServicio & "' "

    lvwActividades.ListItems.Clear
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwActividades.ListItems.Add(, , !Codigo)
                    lsiItem.SubItems(1) = !Nombre
                    lsiItem.SubItems(2) = !TIEMPO
                    'lsiItem.SubItems(3) = Format(!Valor, "###,###")
                    lsiItem.SubItems(3) = FormatoValor(!TIEMPO * gcurPrecioManoObra, "", gintDecimalesMoneda)
                    lsiItem.SubItems(4) = !ESPECIAL
                    lsiItem.SubItems(5) = !IDESPE
                    .MoveNext
                Wend
            End If
        End With
    End If

End Sub

Sub Servicios_del_Modelo(strMarca As String, strModelo As String)
    
    lvwServicios.ListItems.Clear
    
    If gstrServiciosMarca = "S" Then
        mstrSQL = "SELECT Tllr_Servicio_Modelo.Id_Servicio AS CODIGO,"
        mstrSQL = mstrSQL & " Tllr_Servicio.Descripcion AS NOMBRE, "
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo.Horas AS HORAS,"
        mstrSQL = mstrSQL & " Tllr_Servicio.Seccion AS OBJETO,"
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo.Valor AS VALOR"
        mstrSQL = mstrSQL & " FROM Tllr_Servicio LEFT OUTER JOIN"
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo ON"
        mstrSQL = mstrSQL & " Tllr_Servicio.Id_Servicio = Tllr_Servicio_Modelo.Id_Servicio"
        mstrSQL = mstrSQL & " And Tllr_Servicio.Id_Marca = Tllr_Servicio_Modelo.Id_Marca"
        mstrSQL = mstrSQL & " WHERE Tllr_Servicio_Modelo.Id_Marca = '" & strMarca & "' AND"
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo.Id_Modelo = '" & strModelo & "' Order by Tllr_Servicio_Modelo.Id_servicio"
    Else
        mstrSQL = "SELECT Tllr_Servicio_Modelo.Id_Servicio AS CODIGO,"
        mstrSQL = mstrSQL & " Tllr_Servicio.Descripcion AS NOMBRE, "
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo.Horas AS HORAS,"
        mstrSQL = mstrSQL & " Tllr_Servicio.Seccion AS OBJETO,"
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo.Valor AS VALOR"
        mstrSQL = mstrSQL & " FROM Tllr_Servicio RIGHT OUTER JOIN"
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo ON"
        mstrSQL = mstrSQL & " Tllr_Servicio.Id_Servicio = Tllr_Servicio_Modelo.Id_Servicio"
        mstrSQL = mstrSQL & " WHERE Tllr_Servicio_Modelo.Id_Marca = '" & strMarca & "' AND"
        mstrSQL = mstrSQL & " Tllr_Servicio_Modelo.Id_Modelo = '" & strModelo & "' Order by Tllr_Servicio_Modelo.Id_servicio"
    End If
    
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwServicios.ListItems.Add(, , !Codigo)
                    lsiItem.SubItems(1) = ValorNulo(!Nombre)
                    lsiItem.SubItems(2) = ValorNulo(!Horas)
                    'lsiItem.SubItems(3) = Format(!Valor, "###,##0")
                    lsiItem.SubItems(3) = Format(!Horas * gcurPrecioManoObra, "###,###.#0")
                    lsiItem.SubItems(4) = IIf(!Objeto = "M", "MECANICA", "CARROCERIA")
                    .MoveNext
                Wend
            End If
        End With
    End If

End Sub

Sub FillMarcas()
    dtcMarca.Enabled = True
'    mstrSQL = "Select Id_marca as CODIGO, Descripcion as Nombre from Glbl_Marca  where VIGENCIA = 'S' order by Descripcion"
    mstrSQL = "Select Id_marca as CODIGO, Descripcion as Nombre from Glbl_Marca  where VIGENCIA = 'S'"
    mstrSQL = mstrSQL & " AND Glbl_Marca.Id_Marca= '" & strIdMarcaDefecto & "' order by Descripcion"
    
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datMarcas
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcMarca.ListField = "Nombre"
                dtcMarca.BoundColumn = "Codigo"
                dtcMarca.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub

Sub FillModelos(strMarca As String)
    dtcModelo.Enabled = True
    'mstrSql = "Select Id_modelo as CODIGO, Id_modelo+' //// '+Descripcion as Nombre from Glbl_Modelo where VIGENCIA = 'S' and Id_marca = '" & strMarca & "'  order by Descripcion"
    mstrSQL = "Select Id_modelo as CODIGO, Descripcion as Nombre from Glbl_Modelo where VIGENCIA = 'S' and Id_marca = '" & strMarca & "'  order by Descripcion"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datModelos
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcModelo.ListField = "Nombre"
                dtcModelo.BoundColumn = "Codigo"
                dtcModelo.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub

Private Sub dtcMarca_Change()
lvwServicios.ListItems.Clear
lvwActividades.ListItems.Clear
lvwRepuestos.ListItems.Clear
If dtcMarca.BoundText <> "" Then
    dtcModelo.Text = ""
    FillModelos dtcMarca.BoundText
End If
End Sub

Private Sub dtcModelo_Change()
lvwServicios.ListItems.Clear
lvwActividades.ListItems.Clear
lvwRepuestos.ListItems.Clear
If dtcModelo.BoundText <> "" Then
    fmeServicios.Enabled = True
    Servicios_del_Modelo dtcMarca.BoundText, dtcModelo.BoundText
    If lvwServicios.ListItems.Count > 0 Then
        Actividades_del_Servicio dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem
        If lvwActividades.ListItems.Count > 0 Then
            Repuestos_de_la_Actividad dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem
        End If
    End If
Else
    fmeServicios.Enabled = False
    Servicios_del_Modelo "", ""
    Actividades_del_Servicio "", "", ""
End If
End Sub



Private Sub Form_Activate()
If mblnSW Then
    If Not Atributos("Glbl", "Tllr_10_0110_0050", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
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

Private Sub lvwActividades_DblClick()
If lvwActividades.ListItems.Count > 0 Then
    strMode = "Edit"
    Set lsiItem = lvwActividades.SelectedItem
    With frmEditaTempActividad
        .Caption = "Editar Actividad"
        .txtMarca = frmTempServiciosMarMod.dtcMarca.Text
        .txtModelo = frmTempServiciosMarMod.dtcModelo.Text
        .txtServicio = frmTempServiciosMarMod.lvwServicios.SelectedItem.SubItems(1): .txtServicio.Enabled = False
        .txtCodigo = lsiItem: .txtCodigo.Enabled = False
        .txtNombre = lsiItem.SubItems(1): .txtNombre.Enabled = False
        .txtHoras = lsiItem.SubItems(2)
        .txtValor = Format$(lsiItem.SubItems(3), "#######")
        .txtEspecialidad.Text = lsiItem.SubItems(4)
        .Show 1
    End With
End If
End Sub

Private Sub lvwActividades_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvwActividades.ListItems.Count > 0 Then
    Repuestos_de_la_Actividad dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem
End If

End Sub

Private Sub lvwRepuestos_DblClick()
If lvwRepuestos.ListItems.Count > 0 Then
strMode = "Edit"
Set lsiItem = lvwRepuestos.SelectedItem
With frmEditaTempRepuesto
    .Caption = "Editar Repuesto"
    .txtMarca = frmTempServiciosMarMod.dtcMarca.Text
    .txtModelo = frmTempServiciosMarMod.dtcModelo.Text
    .txtServicio = frmTempServiciosMarMod.lvwServicios.SelectedItem.SubItems(1)
    .txtActividad = frmTempServiciosMarMod.lvwActividades.SelectedItem.SubItems(1)
    .txtCodigo = lsiItem
    .txtDescripcion = lsiItem.SubItems(1)
    .txtValor = lsiItem.SubItems(3)
    .txtCantidad = 0
    .Show 1
End With
End If
End Sub

Private Sub lvwServicios_DblClick()

If lvwServicios.ListItems.Count > 0 Then
    strMode = "Edit"
    Set lsiItem = lvwServicios.SelectedItem
    With frmEditaTempServicio
        .Caption = "Editar Servicio"
        .txtMarca = frmTempServiciosMarMod.dtcMarca.Text
        .txtModelo = frmTempServiciosMarMod.dtcModelo.Text
        .txtCodigo = lsiItem: .txtCodigo.Enabled = False
        .txtDescripcion = lsiItem.SubItems(1)
        .txtHoras = lsiItem.SubItems(2)
        .txtValor = Format$(lsiItem.SubItems(3), "######0")
        If lsiItem.SubItems(4) = "MECANICA" Then
            .optObjeto(0).Value = True
        Else
            .optObjeto(1).Value = True
        End If
        .Show 1
    End With

End If
End Sub

Private Sub lvwServicios_ItemClick(ByVal Item As MSComctlLib.ListItem)

If lvwServicios.ListItems.Count > 0 Then
    Actividades_del_Servicio dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem
    If lvwActividades.ListItems.Count > 0 Then
        Repuestos_de_la_Actividad dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem
    Else
        lvwRepuestos.ListItems.Clear
    End If
Else
    lvwActividades.ListItems.Clear
    lvwRepuestos.ListItems.Clear
End If

End Sub

Private Sub tlbOpciones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Select Case Index
    Case 0 '//////////////////////////////// SERVICIOS
        Select Case Button.Key
        Case "Agregar" '////nuevo
            gstrProcedencia = "Temparios"
            frmSelTempServicios.Show 1
        Case "Quitar" '////quitar
            If lvwServicios.ListItems.Count > 0 Then
                If lvwServicios.SelectedItem <> "" Then
                    EliminarItem 0, dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem
                End If
            End If
        Case "Copiar" '/////buscar
            If Me.lvwServicios.ListItems.Count > 0 Then
                frmCopiaServiciosMarMod.Show vbModal
            End If
        End Select
    Case 1 '//////////////////////////////// ACTIVIDADES
        Select Case Button.Key
        Case "Agregar"
            frmSelTempActividades.Show 1
        Case "Quitar" '////quitar
            If lvwActividades.ListItems.Count > 0 Then
                If lvwServicios.SelectedItem <> "" And lvwActividades.SelectedItem <> "" Then
                    EliminarItem 1, dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem
                End If
            End If
        Case "Buscar" '/////buscar
        End Select
    Case 2  '//////////////////////////////// REPUESTOS
        Select Case Button.Key
        Case "Agregar"
            frmSelTempRepuestos.Show 1
        Case "Quitar" '////quitar
            If lvwRepuestos.ListItems.Count > 0 Then
                If lvwServicios.SelectedItem <> "" Then
                    If lvwActividades.SelectedItem <> "" Then
                        If lvwRepuestos.SelectedItem <> "" Then
                            EliminarItem 2, dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem, lvwRepuestos.SelectedItem
                        End If
                    End If
                End If
            End If
        End Select

End Select

End Sub
