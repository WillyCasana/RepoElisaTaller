VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddServiciosMarMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Servicios"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmAddServiciosMarMod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNroRecord 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "10"
      Top             =   1650
      Width           =   510
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   735
      Width           =   2235
   End
   Begin VB.TextBox txtDes 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   1185
      Width           =   5130
   End
   Begin VB.ComboBox cboCoincidir 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmAddServiciosMarMod.frx":038A
      Left            =   1440
      List            =   "frmAddServiciosMarMod.frx":039A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1635
      Width           =   2220
   End
   Begin VB.CheckBox optCriterios 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   765
      Width           =   990
   End
   Begin VB.Frame fmeServicios 
      Caption         =   "Servicios del Modelo"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   7455
      Begin MSComctlLib.ListView lvwServicios 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5953
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
         NumItems        =   4
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
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   450
         Top             =   -45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":03ED
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":04FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":0957
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":0DAF
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1207
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1319
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":142B
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":153D
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":164F
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1761
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1873
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1985
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1A97
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1BA9
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1CBB
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1DCD
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1EDF
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":1FF1
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":2103
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":2215
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":2667
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":2AB9
               Key             =   "Copiar"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServiciosMarMod.frx":2BCB
               Key             =   "Salir"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Todos"
            Key             =   "SelectAll"
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ninguno"
            Key             =   "UnSelectAll"
            Object.ToolTipText     =   "Quitar Servicio"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   5760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "Agregar"
            Object.ToolTipText     =   "Quitar Servicio"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox optCriterios 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   10
      Top             =   1230
      Width           =   1305
   End
   Begin MSComctlLib.Toolbar tlbModelo 
      Height          =   330
      Left            =   7320
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
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
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.UpDown updNroRecord 
      Height          =   315
      Left            =   6150
      TabIndex        =   16
      Top             =   1650
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   5
      BuddyControl    =   "txtNroRecord"
      BuddyDispid     =   196609
      OrigLeft        =   8445
      OrigTop         =   300
      OrigRight       =   8685
      OrigBottom      =   615
      Max             =   100
      Min             =   5
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.Toolbar tlbMarca 
      Height          =   330
      Left            =   2880
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
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
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMarca 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   810
      TabIndex        =   12
      Top             =   90
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro. de Registros :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3990
      TabIndex        =   18
      Top             =   1695
      Width           =   1860
   End
   Begin VB.Label lblModelo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4125
      TabIndex        =   14
      Top             =   75
      Width           =   3315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   90
      X2              =   7440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   90
      X2              =   7440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Coincidir en :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modelo :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3390
      TabIndex        =   2
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Marca :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmAddServiciosMarMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnSW As Boolean
Dim adoPrincipal As New ADODB.Recordset
Dim AdoModelo As New ADODB.Recordset
Dim mstrSql As String
Dim mstrWhere As String

Dim lsiItemSelected As Boolean
Dim lsiItem As ListItem, itmFound As ListItem
Dim intContador As Integer
Const mcintHeight As Integer = 7900
Const mcintWidth As Integer = 11900
Const mcstrMensaje As String = "Confirma Eliminar El Item Seleccionado desde "



Sub EliminarItem(intTipo As Integer, strMarca As String, strModelo As String, Optional strServicio As String, Optional strActividad As String, Optional strRepuesto As String)
Dim strSql As String

Select Case intTipo
    
    Case 0 '////////elimina servicio
        If MsgBox(mcstrMensaje & "Servicios por Modelos", 4 + 32) = vbYes Then
            strSql = "SELECT COUNT(*) AS CUANTOS FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
            If Conexion.SendHost(strSql, AdoModelo, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
                With AdoModelo
                    .MoveFirst
                    If !CUANTOS > 0 Then
                        '////////////////// TIENE ACTIVIDADES RELACIONADAS
                        MsgBox "TIENE ACTIVIDADES RELACIONADAS"
                    Else
                        '////////////////// NO TIENE ACTIVIDADES RELACIONADAS
                        MsgBox "NO TIENE ACTIVIDADES RELACIONADAS"
                        strSql = "DELETE FROM Tllr_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                        Conexion.SendHost strSql, , , , gcTiempoEspera
                        lvwServicios.ListItems.Remove lvwServicios.SelectedItem.Index
                    End If
                End With
            End If
            
        End If
        
    
End Select

End Sub




Sub ServiciosdelModelo(strCondicion As String, strOrden As String)
    
lvwServicios.ListItems.Clear
Dim Valor As Double

Valor = 10000

If gstrServiciosMarca = "S" Then
    mstrSql = "SELECT  TOP " & CStr(updNroRecord.Value) & " Tllr_Servicio_Modelo.Id_Servicio AS ID,"
    mstrSql = mstrSql & " Tllr_Servicio.Descripcion AS DES,"
    mstrSql = mstrSql & " Tllr_Servicio_Modelo.Horas AS TIEMPO,"
    mstrSql = mstrSql & " " & ValorHora(gstrIdEmpresa, gstrIdSucursal) & " AS VALOR"
    mstrSql = mstrSql & " FROM Tllr_Servicio_Modelo RIGHT OUTER JOIN Tllr_Servicio ON Tllr_Servicio_Modelo.Id_Servicio = Tllr_Servicio.Id_Servicio AND Tllr_Servicio_Modelo.Id_Marca = Tllr_Servicio.Id_Marca"
    mstrSql = mstrSql & strCondicion & " " & strOrden
Else
    mstrSql = "SELECT  TOP " & CStr(updNroRecord.Value) & " Tllr_Servicio_Modelo.Id_Servicio AS ID,"
    mstrSql = mstrSql & " Tllr_Servicio.Descripcion AS DES,"
    mstrSql = mstrSql & " Tllr_Servicio_Modelo.Horas AS TIEMPO,"
    mstrSql = mstrSql & " " & ValorHora(gstrIdEmpresa, gstrIdSucursal) & " AS VALOR"
    mstrSql = mstrSql & " FROM Tllr_Servicio_Modelo LEFT OUTER JOIN Tllr_Servicio ON Tllr_Servicio_Modelo.Id_Servicio = Tllr_Servicio.Id_Servicio"
    mstrSql = mstrSql & strCondicion & " " & strOrden
End If


If Conexion.SendHost(mstrSql, AdoModelo, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoModelo
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set lsiItem = lvwServicios.ListItems.Add(, , !ID)
                lsiItem.SubItems(1) = !Des
                lsiItem.SubItems(2) = !TIEMPO
                lsiItem.SubItems(3) = FormatoValor(!Valor, "", gintDecimalesMoneda)
                .MoveNext
            Wend
        End If
    End With
End If
Conexion.CloseHost AdoModelo

End Sub

Private Sub Form_Activate()
If mblnSW Then
    If gstrProcedencia = "Movimientos" Or gstrProcedencia = "Presupuestos" Then
        lblMarca.Caption = frmRecepcion.lblMarca.Caption
        lblMarca.Tag = frmRecepcion.lblIdMarca.Caption
        lblModelo.Caption = frmRecepcion.lblModelo.Caption
        lblModelo.Tag = frmRecepcion.lblIdModelo.Caption
    ElseIf gstrProcedencia = "Temparios" Then
        lblMarca.Caption = frmRecepcion.lblMarca.Caption
        lblMarca.Tag = frmRecepcion.lblIdMarca.Caption
        lblModelo.Caption = frmRecepcion.lblModelo.Caption
        lblModelo.Tag = frmRecepcion.lblIdModelo.Caption
    ElseIf gstrProcedencia = "Reserva_Horas" Then
        lblMarca.Caption = frmReservadeHoras.lblMarca.Caption
        lblMarca.Tag = frmReservadeHoras.lblIdMarca.Caption
        lblModelo.Caption = frmReservadeHoras.lblModelo.Caption
        lblModelo.Tag = frmReservadeHoras.lblIdModelo.Caption
    End If
    cboCoincidir.ListIndex = 0
    mblnSW = False
End If

End Sub

Private Sub Form_Load()

mblnSW = True
updNroRecord.Value = gintNroRecDefectoQry
End Sub

Private Sub optCriterios_Click(Index As Integer)
With Me
    Select Case Index
    Case 0
        If .optCriterios(0).Value = 1 Then ' codigo
            .optCriterios(1).Value = 0
            .txtDes.Enabled = False
            .txtDes.Text = ""
            .txtCodigo.Enabled = True
            .txtCodigo.SetFocus
        Else
            .txtCodigo.Enabled = False
            .txtCodigo.Text = ""
        End If
    Case 1 '////////////////---------------descripcion
        If .optCriterios(1).Value = 1 Then
            .optCriterios(0).Value = 0
            .txtDes.Enabled = True
            .txtCodigo.Enabled = False
            .txtCodigo.Text = ""
            .txtDes.SetFocus
        Else
            .txtDes.Enabled = False
            .txtDes.Text = ""
        
        End If
    End Select
End With
End Sub

Private Sub tlbMarca_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Glbl_Marca", "Id_Marca", "Descripcion", "Busca Marca")
    lblMarca.Tag = gstrBusca
    lblMarca.Caption = MarcaD(gstrBusca)
    lblModelo.Caption = ""
End If

End Sub

Private Sub tlbModelo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    If lblMarca.Tag <> "" Then
        gstrBusca = apfFormulario.BuscarRegistrosModelo(Conexion, "Glbl_Modelo", "Id_Modelo", "Id_Marca", "Descripcion", "Busca Modelo", lblMarca.Tag)
        lblModelo.Tag = gstrBusca
        lblModelo.Caption = ModeloD(lblMarca.Tag, gstrBusca)
    Else
        MsgBox "Seleccione la Marca"
    End If
End If
End Sub

Private Sub tlbOpciones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Select Case Index
Case 0 '//////////////////////////////// seleccionar todos o no
    Select Case Button.Key
    Case "SelectAll" '////Todos
        SelectingItem lvwServicios, gcSelectAll
    Case "UnSelectAll" '////Ninguno
        SelectingItem lvwServicios, gcUnSelectAll
    End Select
Case 1 '//////////////////////////////// buscar, Agregar y cerrar
    'gstrIdCargo = gstrIdCargoDefecto
    Select Case Button.Key
    Case "Agregar" '////Agregar
        If gstrProcedencia = "Movimientos" Or gstrProcedencia = "Presupuestos" Then
            For intContador = 1 To lvwServicios.ListItems.Count
                Set lvwServicios.SelectedItem = lvwServicios.ListItems(intContador)
                If lvwServicios.ListItems(intContador).Checked = True Then
                    Set itmFound = frmRecepcion.lvwServiciosMecanica.FindItem(lvwServicios.SelectedItem, lvwText, , 0)
                    If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                        Set itmFound = frmRecepcion.lvwServiciosMecanica.ListItems.Add(, , lvwServicios.ListItems(intContador))
                        Set frmRecepcion.lvwServiciosMecanica.SelectedItem = itmFound
                        itmFound.SubItems(1) = lvwServicios.ListItems(intContador).SubItems(1)
                        itmFound.SubItems(2) = lvwServicios.ListItems(intContador).SubItems(2)
                        itmFound.SubItems(3) = lvwServicios.ListItems(intContador).SubItems(3)
                        itmFound.SubItems(4) = FormatoValor(0, "", 1)
                        itmFound.SubItems(5) = 0
                        itmFound.SubItems(6) = gstrIdCargo   'gstrIdCargoDefecto
                        itmFound.SubItems(7) = TraeCargoDes(gstrIdCargo)   'Defecto)
                        itmFound.SubItems(8) = gstrMecanicoDefectoSecMec
                        itmFound.SubItems(9) = MecanicoD(gstrMecanicoDefectoSecMec)
                        itmFound.SubItems(10) = FormatoValor(frmRecepcion.CalculoSubTotal(mcFichaMecanica), "", gintDecimalesMoneda)
                        itmFound.SubItems(11) = "N"
                        itmFound.SubItems(12) = "S"
                    End If
                End If
            Next
            Unload Me
        ElseIf gstrProcedencia = "Presupuesto" Then
            For intContador = 1 To lvwServicios.ListItems.Count
                Set lvwServicios.SelectedItem = lvwServicios.ListItems(intContador)
                If lvwServicios.ListItems(intContador).Checked = True Then
                    Set itmFound = frmPresupuesto.lvwServiciosMecanica.FindItem(lvwServicios.SelectedItem, lvwText, , 0)
                    If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                        Set itmFound = frmPresupuesto.lvwServiciosMecanica.ListItems.Add(, , lvwServicios.ListItems(intContador))
                        Set frmPresupuesto.lvwServiciosMecanica.SelectedItem = itmFound
                        itmFound.SubItems(1) = lvwServicios.ListItems(intContador).SubItems(1)
                        itmFound.SubItems(2) = lvwServicios.ListItems(intContador).SubItems(2)
                        itmFound.SubItems(3) = lvwServicios.ListItems(intContador).SubItems(3)
                        itmFound.SubItems(4) = FormatoValor(0, "", 1)
                        itmFound.SubItems(5) = 0
                        itmFound.SubItems(6) = gstrIdCargoDefecto
                        itmFound.SubItems(7) = TraeCargoDes(gstrIdCargoDefecto)
                        itmFound.SubItems(8) = gstrMecanicoDefectoSecMec
                        itmFound.SubItems(9) = MecanicoD(gstrMecanicoDefectoSecMec)
                        itmFound.SubItems(10) = FormatoValor(frmPresupuesto.CalculoSubTotal(mcFichaMecanica), "", gintDecimalesMoneda)
                    End If
                End If
            Next
        ElseIf gstrProcedencia = "Reserva_Horas" Then
            For intContador = 1 To lvwServicios.ListItems.Count
                Set lvwServicios.SelectedItem = lvwServicios.ListItems(intContador)
                If lvwServicios.ListItems(intContador).Checked = True Then
                
                    'Set itmFound = frmReservadeHoras.lvwServiciosMecanica.FindItem(lvwServicios.SelectedItem, lvwText, , 0)
                    'If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                        Set itmFound = frmReservadeHoras.lvwServiciosMecanica.ListItems.Add(, , lvwServicios.ListItems(intContador))
                        Set frmReservadeHoras.lvwServiciosMecanica.SelectedItem = itmFound
                        itmFound.SubItems(1) = lvwServicios.ListItems(intContador).SubItems(1)
                        itmFound.SubItems(2) = lvwServicios.ListItems(intContador).SubItems(2)
                    'End If
                End If
            Next
        End If
        Unload Me
    Case "Cerrar" '////cerrar
        Unload Me
    Case "Buscar" '/////buscar
        If optCriterios(0).Value = 1 Then '/////////////// codigo
            mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And  Tllr_Servicio_Modelo.id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
            ServiciosdelModelo mstrWhere, "Order By Tllr_Servicio_Modelo.Id_Servicio"
        ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
            mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And Tllr_Servicio.Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
            ServiciosdelModelo mstrWhere, " Order by Descripcion"
        Else
            mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  "
            ServiciosdelModelo mstrWhere, ""
        End If
    End Select
End Select

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optCriterios(0).Value = 1 Then '/////////////// codigo
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And  Tllr_Servicio_Modelo.id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, "Order By Tllr_Servicio_Modelo.Id_Servicio"
     ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And Tllr_Servicio.Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, " Order by Descripcion"
     Else
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  "
         ServiciosdelModelo mstrWhere, ""
     End If
End If
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optCriterios(0).Value = 1 Then '/////////////// codigo
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And  Tllr_Servicio_Modelo.id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, "Order By Tllr_Servicio_Modelo.Id_Servicio"
     ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And Tllr_Servicio.Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, " Order by Descripcion"
     Else
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  "
         ServiciosdelModelo mstrWhere, ""
     End If
End If
End Sub
