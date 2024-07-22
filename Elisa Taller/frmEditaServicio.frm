VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEditaServicio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servicios Solicitados"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCerrarAlSeleccionar 
      Caption         =   "Cerrar al seleccionar"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   4575
   End
   Begin MSComctlLib.ListView lsvServicioEspecifico 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Linea"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Servicio Específico"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Id_Servicio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   6255
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   4920
         TabIndex        =   16
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   3600
         TabIndex        =   15
         Top             =   210
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPorcentDescuento 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtDescuento 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc datTipoServicio 
      Height          =   270
      Left            =   2280
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   476
      ConnectMode     =   0
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
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
      Caption         =   "datTipoServicio"
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
   Begin MSDataListLib.DataCombo dbcboTipoServicio 
      Bindings        =   "frmEditaServicio.frx":0000
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Descripcion"
      BoundColumn     =   "Id_Concepto_Servicio"
      Text            =   "dbcboTipoServicio"
   End
   Begin MSAdodcLib.Adodc datMecanico 
      Height          =   270
      Left            =   1920
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   476
      ConnectMode     =   0
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
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
      Caption         =   "datMecanico"
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
   Begin MSDataListLib.DataCombo dbcboMecanico 
      Bindings        =   "frmEditaServicio.frx":001E
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Id_Mecanico"
      Text            =   "dbcboMecanico"
   End
   Begin VB.Label Label8 
      Caption         =   "Desc. (%)"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Mecánico Asignado"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
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
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Descuento"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Valor"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Servicio"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmEditaServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Item As ListItem

Public Sub Calcule_Total()

If Trim$(Me.txtValor.Text) = "" Then
    Me.txtValor.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
End If
If Trim$(Me.txtCantidad.Text) = "" Then
    Me.txtCantidad.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
End If
If Trim$(Me.txtDescuento.Text) = "" Then
    Me.txtDescuento.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
End If
If Trim$(Me.txtPorcentDescuento.Text) = "" Then
    Me.txtPorcentDescuento.Text = "0"
End If

Me.txtTotal.Text = (CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal)) _
                   * CDbl(Me.txtCantidad.Text)) _
                   - CDbl(SacarFormatoValor(Me.txtDescuento.Text, gstrMonedaLocal))
                
Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)

End Sub

Public Sub LLena_TipoServicio()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Concepto_Servicio, Descripcion FROM Srvt_Concepto_Servicio WHERE Vigencia='S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datTipoServicio.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Mecanico()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Mecanico, Nombre FROM Tllr_Mecanicos WHERE Vigencia='S' ORDER BY Nombre"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datMecanico.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Servicio(IdTipoServicio As String)
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Servicio, Descripcion FROM Srvt_Servicios WHERE Id_Concepto_Servicio='" & IdTipoServicio & "' AND Vigencia='S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then
        Tabla.MoveFirst
        While Tabla.EOF = False
            Set Item = Me.lsvServicioEspecifico.ListItems.Add(, , Me.lsvServicioEspecifico.ListItems.Count + 1)
            Item.SubItems(1) = Tabla!Descripcion
            Item.SubItems(2) = Tabla!Id_Servicio
            Item.SubItems(3) = Tabla!Id_Servicio
            Tabla.MoveNext
        Wend
    End If
End If

End Sub

Private Sub cmdAceptar_Click()

If Me.dbcboTipoServicio.BoundText = "" Then
    MsgBox "Debe seleccionar un Tipo de Servicio.", vbExclamation, "ServiPro"
    Me.dbcboTipoServicio.SetFocus
    Exit Sub
End If

If Me.lsvServicioEspecifico.ListItems.Count = 0 Then
    MsgBox "Imposible continuar. No hay servicios registrados.", vbExclamation, "ServiPro"
    Exit Sub
End If

If Me.dbcboMecanico.BoundText = "" Then
    MsgBox "Debe asignar un Mecánico al Servicio.", vbExclamation, "ServiPro"
    Me.dbcboMecanico.SetFocus
    Exit Sub
End If

If gblnNuevo = True Then
    Set Item = frmOtServiteca.lsvServicios.ListItems.Add(, , frmOtServiteca.lsvServicios.ListItems.Count + 1)
    Item.SubItems(1) = Me.dbcboTipoServicio.Text
    Item.SubItems(2) = Me.lsvServicioEspecifico.SelectedItem.SubItems(1)
    Item.SubItems(3) = Me.txtCantidad.Text
    Item.SubItems(4) = Me.txtValor.Text
    Item.SubItems(5) = Me.txtDescuento.Text
    Item.SubItems(6) = Me.txtTotal.Text
    Item.SubItems(7) = Me.dbcboMecanico.Text
    Item.SubItems(8) = Me.dbcboTipoServicio.BoundText
    Item.SubItems(9) = Me.lsvServicioEspecifico.SelectedItem.SubItems(2)
    Item.SubItems(10) = Me.dbcboMecanico.BoundText
Else
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(1) = Me.dbcboTipoServicio.Text
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(2) = Me.lsvServicioEspecifico.SelectedItem.SubItems(1)
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(3) = Me.txtCantidad.Text
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(4) = Me.txtValor.Text
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(5) = Me.txtDescuento.Text
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(6) = Me.txtTotal.Text
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(7) = Me.dbcboMecanico.Text
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(8) = Me.dbcboTipoServicio.BoundText
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(9) = Me.lsvServicioEspecifico.SelectedItem.SubItems(2)
    frmOtServiteca.lsvServicios.SelectedItem.SubItems(10) = Me.dbcboMecanico.BoundText
End If
If Me.chkCerrarAlSeleccionar.Value = 1 Then
    Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub


Private Sub dbcboTipoServicio_Click(Area As Integer)
If Area = 2 Then
    Me.lsvServicioEspecifico.ListItems.Clear
    Me.txtValor.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    LLena_Servicio (Me.dbcboTipoServicio.BoundText)
End If
End Sub

Private Sub Form_Activate()
Me.lsvServicioEspecifico.SetFocus
End Sub

Private Sub Form_Load()
Dim ldblCont As Double
Dim ldblCont2 As Double

LLena_TipoServicio
LLena_Mecanico
Me.txtCantidad.Text = 0
Me.txtValor.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
Me.txtDescuento.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
Me.txtPorcentDescuento.Text = "0"
Me.txtTotal.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)

If gblnNuevo = True Then
    Me.Caption = Me.Caption & " - AGREGANDO REGISTRO"
    Me.chkCerrarAlSeleccionar.Value = 1
    Me.chkCerrarAlSeleccionar.Visible = True
Else
    Me.Caption = Me.Caption & " - EDITANDO REGISTRO"
    Me.chkCerrarAlSeleccionar.Value = 1
    Me.chkCerrarAlSeleccionar.Visible = False
    For ldblCont = 1 To frmOtServiteca.lsvServicios.ListItems.Count
        If ldblCont > frmOtServiteca.lsvServicios.ListItems.Count Then
            Exit For
        End If
        If frmOtServiteca.lsvServicios.ListItems(ldblCont).Selected = True Then
            Me.dbcboTipoServicio.BoundText = frmOtServiteca.lsvServicios.SelectedItem.SubItems(8)
            LLena_Servicio (Me.dbcboTipoServicio.BoundText)
            If Me.lsvServicioEspecifico.ListItems.Count > 0 Then
                For ldblCont2 = 1 To Me.lsvServicioEspecifico.ListItems.Count
                    If Me.lsvServicioEspecifico.ListItems(ldblCont2).SubItems(2) = frmOtServiteca.lsvServicios.SelectedItem.SubItems(9) Then
                        Me.lsvServicioEspecifico.ListItems(ldblCont2).Selected = True
                        Exit For
                    End If
                Next ldblCont2
            End If
            Me.txtCantidad.Text = frmOtServiteca.lsvServicios.SelectedItem.SubItems(3)
            Me.txtValor.Text = frmOtServiteca.lsvServicios.SelectedItem.SubItems(4)
            Me.txtDescuento.Text = frmOtServiteca.lsvServicios.SelectedItem.SubItems(5)
            Me.txtTotal.Text = frmOtServiteca.lsvServicios.SelectedItem.SubItems(6)
            Me.dbcboMecanico.BoundText = frmOtServiteca.lsvServicios.SelectedItem.SubItems(10)
        End If
    Next ldblCont
End If

End Sub

Private Sub lsvServicioEspecifico_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Valor FROM Srvt_Servicios WHERE Id_Concepto_Servicio='" & Me.dbcboTipoServicio.BoundText & "' AND Id_Servicio='" & Me.lsvServicioEspecifico.SelectedItem.SubItems(2) & "' AND Vigencia='S'"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Me.txtValor.Text = FormatoValor(Tabla!Valor, gstrMonedaLocal, gintDecimalesMoneda)
    If CDbl(Me.txtCantidad.Text) = 0 Then
        Me.txtCantidad.Text = "1"
    End If
Else
    MsgBox "Falla al intentar conectar con el servidor." & Chr(13) & "Intente nuevamente."
    Me.txtValor.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    Me.txtCantidad.Text = "0"
End If
Conexion.CloseHost Tabla
Me.txtValor.Tag = Me.txtValor.Text
End Sub

Private Sub txtCantidad_Change()

If Trim$(Me.txtCantidad.Text) = "" Then
    Me.txtCantidad.Text = "0"
End If
If Not IsNumeric(Me.txtCantidad.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtCantidad.Text = "0"
End If
If Me.txtPorcentDescuento.Text = "" Then Me.txtPorcentDescuento.Text = "0"
If CDbl(Me.txtPorcentDescuento.Text) <> 0 Then
    Me.txtDescuento.Text = (CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal)) * CDbl(Me.txtCantidad.Text)) * (Me.txtPorcentDescuento.Text / 100)
    Me.txtDescuento.Text = FormatoValor(Me.txtDescuento.Text, gstrMonedaLocal, gintDecimalesMoneda)
Else
    Me.txtDescuento.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
End If
Calcule_Total
End Sub

Private Sub txtCantidad_GotFocus()
Me.txtCantidad.SelStart = 0
Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
End Sub

Private Sub txtCantidad_LostFocus()
If Not IsNumeric(Me.txtCantidad.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtCantidad.Text = "0"
End If
End Sub

Private Sub txtDescuento_Change()
If Trim$(Me.txtDescuento.Text) = "" Then
    Me.txtDescuento.Text = "0"
End If
If Not IsNumeric(Me.txtDescuento.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtDescuento.Text = "0"
End If
Calcule_Total
End Sub

Private Sub txtDescuento_GotFocus()
Me.txtDescuento.Text = SacarFormatoValor(Me.txtDescuento.Text, gstrMonedaLocal)
Me.txtDescuento.SelStart = 0
Me.txtDescuento.SelLength = Len(Me.txtDescuento.Text)
End Sub

Private Sub txtDescuento_LostFocus()
If Not IsNumeric(Me.txtDescuento.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtDescuento.Text = "0"
End If
Me.txtPorcentDescuento.Text = Round((CDbl(Me.txtDescuento.Text) / (CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal)) * CDbl(Me.txtCantidad.Text))) * 100, 4)
If Mid$(Me.txtDescuento.Text, 1, 1) <> gstrMonedaLocal Then Me.txtDescuento.Text = FormatoValor(Me.txtDescuento.Text, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtPorcentDescuento_Change()

    If Trim$(Me.txtValor.Text) = "" Then
        Me.txtValor.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    End If
    If Not IsNumeric(Me.txtPorcentDescuento.Text) Then
        MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
        Me.txtPorcentDescuento.Text = "0"
    End If
    
    If Trim$(Me.txtCantidad.Text) = "" Then
        Me.txtCantidad.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    End If
    If Trim$(Me.txtDescuento.Text) = "" Then
        Me.txtDescuento.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    End If
    If Trim$(Me.txtPorcentDescuento.Text) = "" Then
        Me.txtPorcentDescuento.Text = "0"
    End If
    
    If CDbl(Me.txtPorcentDescuento.Text) < 0 Or CDbl(Me.txtPorcentDescuento.Text) > 100 Then
        MsgBox "El Rango de descuento debe comprender entre 0 (cero )y 100 (cien)", vbExclamation, "ServiPro"
        Me.txtPorcentDescuento.Text = "0"
        Exit Sub
    End If
    
    If CDbl(Me.txtPorcentDescuento.Text) <> 0 Then
        Me.txtDescuento.Text = (CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal)) * CDbl(Me.txtCantidad.Text)) * (Me.txtPorcentDescuento.Text / 100)
        Me.txtDescuento.Text = FormatoValor(Me.txtDescuento.Text, gstrMonedaLocal, gintDecimalesMoneda)
    Else
        Me.txtDescuento.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    End If
End Sub

Private Sub txtPorcentDescuento_GotFocus()
Me.txtPorcentDescuento.SelStart = 0
Me.txtPorcentDescuento.SelLength = Len(Me.txtPorcentDescuento.Text)
End Sub

Private Sub txtPorcentDescuento_LostFocus()
If Not IsNumeric(Me.txtPorcentDescuento.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtPorcentDescuento.Text = "0"
End If
End Sub

Private Sub txtValor_Change()
If Not IsNumeric(Me.txtValor.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtValor.Text = "0"
End If
If Not IsNumeric(Me.txtValor.Text) Then
    MsgBox "El valor ingresado debe ser del tipo numérico."
    Me.txtValor.Text = Me.txtValor.Tag
End If
If CDbl(Me.txtPorcentDescuento.Text) <> 0 Then
    Me.txtDescuento.Text = (CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal)) * CDbl(Me.txtCantidad.Text)) * (Me.txtPorcentDescuento.Text / 100)
    Me.txtDescuento.Text = FormatoValor(Me.txtDescuento.Text, gstrMonedaLocal, gintDecimalesMoneda)
Else
    Me.txtDescuento.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
End If
Calcule_Total
End Sub

Private Sub txtValor_GotFocus()
Me.txtValor.Text = SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal)
Me.txtValor.SelStart = 0
Me.txtValor.SelLength = Len(Me.txtValor.Text)
End Sub

Private Sub txtValor_LostFocus()
If Not IsNumeric(Me.txtValor.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtValor.Text = "0"
End If
Me.txtValor.Text = FormatoValor(Me.txtValor.Text, gstrMonedaLocal, gintDecimalesMoneda)
End Sub
