VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTemparioServiciosMarMod 
   Caption         =   "Servicios, Actividades y Repuestos  por Modelo"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   Icon            =   "frmTemparioServiciosMarMod.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeRepuesto 
      Caption         =   "Repuestos asociados a la Actividad"
      Height          =   2250
      Left            =   75
      TabIndex        =   6
      Top             =   4980
      Width           =   11500
      Begin MSComctlLib.ListView lvwRepuestos 
         Height          =   1920
         Left            =   60
         TabIndex        =   9
         Top             =   225
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   3387
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
         Appearance      =   1
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
         Height          =   990
         Index           =   2
         Left            =   10410
         TabIndex        =   12
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1746
         ButtonWidth     =   1693
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
   End
   Begin VB.Frame fmeActividades 
      Caption         =   "Actividades asociadas al Servicio"
      Height          =   2370
      Left            =   60
      TabIndex        =   5
      Top             =   2610
      Width           =   11520
      Begin MSComctlLib.ListView lvwActividades 
         Height          =   2085
         Left            =   60
         TabIndex        =   8
         Top             =   195
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
         Appearance      =   1
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
         Height          =   990
         Index           =   1
         Left            =   10470
         TabIndex        =   11
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1746
         ButtonWidth     =   1693
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
   End
   Begin MSDataListLib.DataCombo dtcModelo 
      Bindings        =   "frmTemparioServiciosMarMod.frx":0442
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   30
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcMarca 
      Bindings        =   "frmTemparioServiciosMarMod.frx":045B
      Height          =   315
      Left            =   675
      TabIndex        =   1
      Top             =   30
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin VB.Frame fmeServicios 
      Caption         =   "Servicios asociados al Modelo"
      Height          =   2205
      Left            =   75
      TabIndex        =   0
      Top             =   405
      Width           =   11500
      Begin MSComctlLib.Toolbar tlbOpciones 
         Height          =   990
         Index           =   0
         Left            =   10455
         TabIndex        =   10
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1746
         ButtonWidth     =   1693
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
         Height          =   1980
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   10290
         _ExtentX        =   18150
         _ExtentY        =   3493
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
         Appearance      =   1
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
         Left            =   10815
         Top             =   1305
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
               Picture         =   "frmTemparioServiciosMarMod.frx":0473
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":0585
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":09DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":0E35
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":128D
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":139F
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":14B1
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":15C3
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":16D5
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":17E7
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":18F9
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":1A0B
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":1B1D
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":1C2F
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":1D41
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":1E53
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":1F65
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":2077
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":2189
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":229B
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":26ED
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTemparioServiciosMarMod.frx":2B3F
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc datModelos 
      Height          =   330
      Left            =   4380
      Top             =   0
      Visible         =   0   'False
      Width           =   2970
      _ExtentX        =   5239
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
      Left            =   675
      Top             =   15
      Visible         =   0   'False
      Width           =   2970
      _ExtentX        =   5239
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
      AutoSize        =   -1  'True
      Caption         =   "Modelo :"
      Height          =   195
      Left            =   3735
      TabIndex        =   4
      Top             =   30
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Marca :"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   540
   End
End
Attribute VB_Name = "frmTemparioServiciosMarMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnSw As Boolean
Dim adoPrincipal As New ADODB.Recordset
Dim mstrSql As String
Dim lsiItemSelected As Boolean
Dim lsiItem As ListItem
Const mcintHeight As Integer = 7700
Const mcintWidth As Integer = 11700
Const mcstrMensaje As String = "Confirma Eliminar El Item Seleccionado desde "
Sub Repuestos_de_la_Actividad(strMarca As String, strModelo As String, strServicio As String, strActividad As String)
    
lvwRepuestos.ListItems.Clear
mstrSql = " SELECT Tllr_Actividad_Repuesto.Id_Item AS CODIGO, "
mstrSql = mstrSql & " Stck_Item.Descripcion AS NOMBRE, "
mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Cantidad AS CANTY, "
mstrSql = mstrSql & " Tllr_Actividad_Repuesto.Valor AS VLR, "
mstrSql = mstrSql & " Stck_Item.Id_Familia AS IDFAM, "
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
                    lsiItem.SubItems(2) = Format(!CanTY, "###,##0")
                    lsiItem.SubItems(3) = Format(!Vlr, "###,##0")
                    lsiItem.SubItems(4) = !FAMILIA
                    lsiItem.SubItems(5) = !IDFAM
                    .MoveNext
                Wend
            End If
        End With
    End If
    
End Sub

Sub EliminarItem(intTipo As Integer, strMarca As String, strModelo As String, Optional strServicio As String, Optional strActividad As String, Optional strRepuesto As String)
Dim strSQL As String
Select Case intTipo
    
Case 0 '////////elimina servicio
    If MsgBox(mcstrMensaje & "Servicios por Modelos", 4 + 32) = vbYes Then
        strSQL = "SELECT COUNT(*) AS CUANTOS FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
        If Conexion.SendHost(strSQL, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
            With adoPrincipal
                .MoveFirst
                If !CUANTOS > 0 Then
                    '////////////////// TIENE ACTIVIDADES RELACIONADAS
                    MsgBox "TIENE ACTIVIDADES RELACIONADAS"
                    '//////////////////
                    strSQL = "DELETE FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSQL, , , , gcTiempoEspera
                    '//////////////////
                    strSQL = "DELETE FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSQL, , , , gcTiempoEspera
                    '//////////////////
                    strSQL = "DELETE FROM Tllr_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSQL, , , , gcTiempoEspera
                    lvwServicios.ListItems.Remove lvwServicios.SelectedItem.Index
                Else
                    '////////////////// NO TIENE ACTIVIDADES RELACIONADAS
                    strSQL = "DELETE FROM Tllr_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                    Conexion.SendHost strSQL, , , , gcTiempoEspera
                    lvwServicios.ListItems.Remove lvwServicios.SelectedItem.Index
                End If
            End With
        End If
        
    End If
    
Case 1 '/////////elimina actividad
    
    If MsgBox(mcstrMensaje & "Actividades por Servicio", 4 + 32) = vbYes Then
        strSQL = "SELECT count(*) AS CUANTOS FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
        If Conexion.SendHost(strSQL, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            With adoPrincipal
                .MoveFirst
                If !CUANTOS > 0 Then '////////////////// TIENE ACTIVIDADES RELACIONADAS
                    MsgBox "TIENE REPUESTOS RELACIONADOS"
                    '//////////////////
                    strSQL = "DELETE FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
                    Conexion.SendHost strSQL, , , , gcTiempoEspera
                    '//////////////////
                    strSQL = "DELETE FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
                    Conexion.SendHost strSQL, , , , gcTiempoEspera
                    lvwActividades.ListItems.Remove lvwActividades.SelectedItem.Index
                    lvwRepuestos.ListItems.Clear
                Else '////////////////// NO TIENE ACTIVIDADES RELACIONADAS
                    strSQL = "DELETE FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' "
                    Conexion.SendHost strSQL, , , , gcTiempoEspera
                    lvwActividades.ListItems.Remove lvwActividades.SelectedItem.Index
                    
                End If
            End With
        End If
        
    End If
Case 2 '/////////elimina repuesto
    If MsgBox(mcstrMensaje & "Repuestos de la Actividad", 4 + 32) = vbYes Then
        strSQL = "DELETE FROM Tllr_Actividad_Repuesto WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' AND Id_Item = '" & strRepuesto & "'"
        Conexion.SendHost strSQL, , , , gcTiempoEspera
        lvwRepuestos.ListItems.Remove lvwRepuestos.SelectedItem.Index
'        If lvwActividades.ListItems.Count > 0 Then
'            Repuestos_de_la_Actividad dtcMarca.BoundText, dtcModelo.BoundText, lvwServicios.SelectedItem, lvwActividades.SelectedItem
'        End If
    End If
End Select

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



Sub Servicios_del_Modelo(strMarca As String, strModelo As String)
    
    lvwServicios.ListItems.Clear
    
    mstrSql = "SELECT Tllr_Servicio_Modelo.Id_Servicio AS CODIGO,"
    mstrSql = mstrSql & " Tllr_Servicio.Descripcion AS NOMBRE, "
    mstrSql = mstrSql & " Tllr_Servicio_Modelo.Horas AS HORAS,"
    mstrSql = mstrSql & " Tllr_Servicio.Seccion AS OBJETO,"
    mstrSql = mstrSql & " Tllr_Servicio_Modelo.Valor AS VALOR"
    mstrSql = mstrSql & " FROM Tllr_Servicio RIGHT OUTER JOIN"
    mstrSql = mstrSql & " Tllr_Servicio_Modelo ON"
    mstrSql = mstrSql & " Tllr_Servicio.Id_Servicio = Tllr_Servicio_Modelo.Id_Servicio"
    mstrSql = mstrSql & " WHERE Tllr_Servicio_Modelo.Id_Marca = '" & strMarca & "' AND"
    mstrSql = mstrSql & " Tllr_Servicio_Modelo.Id_Modelo = '" & strModelo & "' "
    
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set lsiItem = lvwServicios.ListItems.Add(, , !Codigo)
                    lsiItem.SubItems(1) = !Nombre
                    lsiItem.SubItems(2) = !Horas
                    lsiItem.SubItems(3) = Format(!Valor, "###,##0")
                    lsiItem.SubItems(4) = IIf(!OBJETO = "M", "MECANICA", "CARROCERIA")
                    .MoveNext
                Wend
            End If
        End With
    End If

End Sub

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
    mstrSql = "Select Id_modelo as CODIGO, Descripcion as Nombre from Glbl_Modelo where VIGENCIA = 'S' and Id_marca = '" & strMarca & "'  order by Descripcion"
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
If mblnSw Then
    FillMarcas
    '///////////////////////////
    mblnSw = False
End If

End Sub

Private Sub Form_Load()

mblnSw = True

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
        .txtMarca = frmTemparioServiciosMarMod.dtcMarca.Text
        .txtModelo = frmTemparioServiciosMarMod.dtcModelo.Text
        .txtServicio = frmTemparioServiciosMarMod.lvwServicios.SelectedItem.SubItems(1): .txtServicio.Enabled = False
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
    .txtMarca = frmTemparioServiciosMarMod.dtcMarca.Text
    .txtModelo = frmTemparioServiciosMarMod.dtcModelo.Text
    .txtServicio = frmTemparioServiciosMarMod.lvwServicios.SelectedItem.SubItems(1)
    .txtActividad = frmTemparioServiciosMarMod.lvwActividades.SelectedItem.SubItems(1)
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
        .txtMarca = frmTemparioServiciosMarMod.dtcMarca.Text
        .txtModelo = frmTemparioServiciosMarMod.dtcModelo.Text
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
'            frmBuscaServiciosPorModelo.Show 1
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
