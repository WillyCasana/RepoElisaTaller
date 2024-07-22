VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMaestroFactores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro Factores de Comisión Serviteca"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmMaestroFactores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   6960
      TabIndex        =   19
      Top             =   6480
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Factor Mecánicos"
      TabPicture(0)   =   "frmMaestroFactores.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lsvMecanicos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Factor Servicios"
      TabPicture(1)   =   "frmMaestroFactores.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "lsvServicios"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   10
         Top             =   4560
         Width           =   8415
         Begin VB.TextBox txtFactorPorcentajeS 
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtFactorMontoS 
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Top             =   1200
            Width           =   2895
         End
         Begin VB.CommandButton cmdAplicarS 
            Caption         =   "&Aplicar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6480
            TabIndex        =   11
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Factor Porcentaje"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Factor Monto"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lbl 
            Caption         =   "Servicio"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblServicio 
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
            Height          =   315
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   7215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   4560
         Width           =   8415
         Begin VB.CommandButton cmdAplicar 
            Caption         =   "&Aplicar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6480
            TabIndex        =   7
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtFactorMonto 
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox txtFactorPorcentaje 
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblMecanico 
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
            Height          =   315
            Left            =   1080
            TabIndex        =   9
            Top             =   240
            Width           =   7215
         End
         Begin VB.Label Label3 
            Caption         =   "Mecánico"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Factor Monto"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Factor Porcentaje"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lsvMecanicos 
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7011
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
            Text            =   "Linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código Mecánico"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre Mecánico"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Factor Porcentaje"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Factor Monto"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lsvServicios 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7011
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
            Text            =   "Linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Concepto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Código Servicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Servicio"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Factor Porcentaje"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Factor Monto"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMaestroFactores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Item As ListItem

Private Sub LLenaListaMecanicos()
Dim Tabla As New ADODB.Recordset
Dim sql As String

Me.lsvMecanicos.ListItems.Clear

sql = ""
sql = "SELECT Tllr_Mecanicos.Id_Mecanico, Tllr_Mecanicos.Nombre, Srvt_Mecanico_Factor.Factor_Porcentaje, Srvt_Mecanico_Factor.Factor_Monto "
sql = sql & "FROM Tllr_Mecanicos LEFT JOIN Srvt_Mecanico_Factor ON Tllr_Mecanicos.Id_Mecanico = Srvt_Mecanico_Factor.Id_Mecanico ORDER BY Tllr_Mecanicos.Nombre"

If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then
        Tabla.MoveFirst
        While Tabla.EOF = False
            Set Item = Me.lsvMecanicos.ListItems.Add(, , Me.lsvMecanicos.ListItems.Count + 1)
            Item.SubItems(1) = Tabla!Id_Mecanico
            Item.SubItems(2) = Tabla!nombre
            If Not IsNull(Tabla!Factor_Porcentaje) Then
                Item.SubItems(3) = FormatoValor(Tabla!Factor_Porcentaje, "%", 2)
            Else
                Item.SubItems(3) = FormatoValor(0, "%", 2)
            End If
            If Not IsNull(Tabla!Factor_Monto) Then
                Item.SubItems(4) = FormatoValor(Tabla!Factor_Monto, gstrMonedaLocal, gintDecimalesMoneda)
            Else
                Item.SubItems(4) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
            End If
            Tabla.MoveNext
        Wend
    End If
End If

End Sub

Private Sub LLenaListaServicios()
Dim Tabla As New ADODB.Recordset
Dim sql As String

Me.lsvServicios.ListItems.Clear

sql = ""
sql = "SELECT Srvt_Servicios.Id_Servicio, Srvt_Concepto_servicio.Descripcion AS Concepto, Srvt_Servicios.Descripcion AS Servicio, Srvt_Servicios.Factor_Porcentaje, Srvt_Servicios.Factor_Monto "
sql = sql & "FROM Srvt_Concepto_servicio RIGHT JOIN Srvt_Servicios ON Srvt_Concepto_servicio.Id_Concepto_Servicio = Srvt_Servicios.Id_Concepto_Servicio ORDER BY Concepto, Servicio"

If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then
        Tabla.MoveFirst
        While Tabla.EOF = False
            Set Item = Me.lsvServicios.ListItems.Add(, , Me.lsvServicios.ListItems.Count + 1)
            Item.SubItems(1) = Tabla!Concepto
            Item.SubItems(2) = Tabla!Id_Servicio
            Item.SubItems(3) = Tabla!Servicio
            If Not IsNull(Tabla!Factor_Porcentaje) Then
                Item.SubItems(4) = FormatoValor(Tabla!Factor_Porcentaje, "%", 2)
            Else
                Item.SubItems(4) = FormatoValor(0, "%", 2)
            End If
            If Not IsNull(Tabla!Factor_Monto) Then
                Item.SubItems(5) = FormatoValor(Tabla!Factor_Monto, gstrMonedaLocal, gintDecimalesMoneda)
            Else
                Item.SubItems(5) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
            End If
            Tabla.MoveNext
        Wend
    End If
End If

End Sub

Private Sub cmdAplicar_Click()
Guardar
End Sub

Private Sub cmdAplicarS_Click()
GuardarServicio
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
LLenaListaMecanicos
LLenaListaServicios
Me.SSTab1.tab = 0
End Sub

Private Sub lblMecanico_Change()
If Trim$(Me.lblMecanico.Tag) = "" Then
    Me.cmdAplicar.Enabled = False
Else
    Me.cmdAplicar.Enabled = True
End If
End Sub

Private Sub lblServicio_Change()
If Trim$(Me.lblServicio.Tag) = "" Then
    Me.cmdAplicarS.Enabled = False
Else
    Me.cmdAplicarS.Enabled = True
End If
End Sub

Private Sub lsvMecanicos_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.txtFactorPorcentaje.Text = Me.lsvMecanicos.SelectedItem.SubItems(3)
Me.txtFactorMonto.Text = Me.lsvMecanicos.SelectedItem.SubItems(4)
Me.lblMecanico.Tag = Me.lsvMecanicos.SelectedItem.SubItems(1)
Me.lblMecanico.Caption = Me.lsvMecanicos.SelectedItem.SubItems(2)
End Sub

Private Sub lsvServicios_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.txtFactorPorcentajeS.Text = Me.lsvServicios.SelectedItem.SubItems(4)
Me.txtFactorMontoS.Text = Me.lsvServicios.SelectedItem.SubItems(5)
Me.lblServicio.Tag = Me.lsvServicios.SelectedItem.SubItems(2)
Me.lblServicio.Caption = Me.lsvServicios.SelectedItem.SubItems(3)
End Sub

Private Sub txtFactorMonto_GotFocus()
Me.txtFactorMonto.Text = SacarFormatoValor(Me.txtFactorMonto.Text, gstrMonedaLocal)
Me.txtFactorMonto.SelStart = 0
Me.txtFactorMonto.SelLength = Len(Me.txtFactorMonto.Text)
End Sub

Private Sub txtFactorMonto_LostFocus()
Me.txtFactorMonto.Text = FormatoValor(Me.txtFactorMonto.Text, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtFactorMontoS_GotFocus()
Me.txtFactorMontoS.Text = SacarFormatoValor(Me.txtFactorMontoS.Text, gstrMonedaLocal)
Me.txtFactorMontoS.SelStart = 0
Me.txtFactorMontoS.SelLength = Len(Me.txtFactorMontoS.Text)
End Sub

Private Sub txtFactorMontoS_LostFocus()
Me.txtFactorMontoS.Text = FormatoValor(Me.txtFactorMontoS.Text, gstrMonedaLocal, 2)
End Sub

Private Sub txtFactorPorcentaje_GotFocus()
Me.txtFactorPorcentaje.Text = SacarFormatoValor(Me.txtFactorPorcentaje.Text, "%")
Me.txtFactorPorcentaje.SelStart = 0
Me.txtFactorPorcentaje.SelLength = Len(Me.txtFactorPorcentaje.Text)
End Sub

Private Sub txtFactorPorcentaje_LostFocus()
Me.txtFactorPorcentaje.Text = FormatoValor(Me.txtFactorPorcentaje.Text, "%", 2)
End Sub

Private Sub Guardar()
Dim Tabla As New ADODB.Recordset
Dim sql As String
Dim mstrSql As String
Dim ldblCont As Double

If Trim$(Me.txtFactorMonto.Text) = "" Then
    Me.txtFactorMonto.Text = FormatoValor(Me.txtFactorMonto.Text, gstrMonedaLocal, gintDecimalesMoneda)
End If
If Trim$(Me.txtFactorPorcentaje.Text) = "" Then
    Me.txtFactorPorcentaje.Text = FormatoValor(Me.txtFactorPorcentaje.Text, "%", 2)
    Me.lsvMecanicos.SetFocus
End If

If Not IsNumeric(SacarFormatoValor(Me.txtFactorMonto.Text, gstrMonedaLocal)) Then
    MsgBox "Debe ingresar un Valor del tipo numérico.", vbExclamation, "Maestro de Factores"""
    Me.txtFactorMonto.SetFocus
    Exit Sub
End If
If Not IsNumeric(SacarFormatoValor(Me.txtFactorPorcentaje.Text, "%")) Then
    MsgBox "Debe ingresar un Valor del tipo numérico.", vbExclamation, "Maestro de Factores"""
    Me.txtFactorPorcentaje.SetFocus
    Exit Sub
End If

If SacarFormatoValor(Me.txtFactorMonto.Text, gstrMonedaLocal) > 0 And SacarFormatoValor(Me.txtFactorPorcentaje.Text, "%") > 0 Then
    MsgBox "Sólo puede ingresar uno de los dos tipos de valores (Porcentaje o Monto)." & Chr(13) & "Intente nuevamente.", vbExclamation, "Maestro de Factores"
    Me.txtFactorMonto.SetFocus
    Exit Sub
End If

sql = ""
mstrSql = ""
sql = "SELECT * FROM Srvt_Mecanico_Factor WHERE Id_Mecanico='" & Me.lsvMecanicos.SelectedItem.SubItems(1) & "'"
If Conexion.SendHost(sql, Tabla, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then 'SI EXISTE MECANICO EN LA TABLA DE FACTORES
        mstrSql = ""
        mstrSql = "UPDATE Srvt_Mecanico_Factor SET "
        mstrSql = mstrSql & "Factor_Porcentaje=" & SacarFormatoValor(Me.txtFactorPorcentaje.Text, "%") & ", "
        mstrSql = mstrSql & "Factor_Monto=" & SacarFormatoValor(Me.txtFactorMonto.Text, gstrMonedaLocal) & " "
        mstrSql = mstrSql & "WHERE Id_Mecanico='" & Me.lsvMecanicos.SelectedItem.SubItems(1) & "'"
    Else
        mstrSql = ""
        mstrSql = mstrSql & "INSERT INTO Srvt_Mecanico_Factor (Id_Mecanico, Factor_Monto, Factor_Porcentaje, Vigencia, Usr_Id, Usr_Fecha) "
        mstrSql = mstrSql & "VALUES ('" & Me.lsvMecanicos.SelectedItem.SubItems(1) & "', "
        mstrSql = mstrSql & SacarFormatoValor(Me.txtFactorMonto.Text, gstrMonedaLocal) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.txtFactorPorcentaje.Text, "%") & ", "
        mstrSql = mstrSql & "'S', "
        mstrSql = mstrSql & "'" & gstrIdUsuario & "', "
        mstrSql = mstrSql & "'" & Date & "')"
    End If
End If
Conexion.CloseHost Tabla


If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
    MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
End If

LLenaListaMecanicos

For ldblCont = 1 To Me.lsvMecanicos.ListItems.Count
    If Me.lsvMecanicos.ListItems(ldblCont).SubItems(1) = Me.lblMecanico.Tag Then
        Me.lsvMecanicos.ListItems(ldblCont).Selected = True
        Me.lsvMecanicos.SetFocus
        Exit For
    End If
Next ldblCont

End Sub

Private Sub GuardarServicio()
Dim Tabla As New ADODB.Recordset
Dim sql As String
Dim mstrSql As String
Dim ldblCont As Double

If Trim$(Me.txtFactorMontoS.Text) = "" Then
    Me.txtFactorMontoS.Text = FormatoValor(Me.txtFactorMontoS.Text, gstrMonedaLocal, gintDecimalesMoneda)
End If
If Trim$(Me.txtFactorPorcentajeS.Text) = "" Then
    Me.txtFactorPorcentajeS.Text = FormatoValor(Me.txtFactorPorcentajeS.Text, "%", 2)
End If

If Not IsNumeric(SacarFormatoValor(Me.txtFactorMontoS.Text, gstrMonedaLocal)) Then
    MsgBox "Debe ingresar un Valor del tipo numérico.", vbExclamation, "Maestro de Factores"""
    Me.txtFactorMontoS.SetFocus
    Exit Sub
End If
If Not IsNumeric(SacarFormatoValor(Me.txtFactorPorcentajeS.Text, "%")) Then
    MsgBox "Debe ingresar un Valor del tipo numérico.", vbExclamation, "Maestro de Factores"""
    Me.txtFactorPorcentajeS.SetFocus
    Exit Sub
End If

If SacarFormatoValor(Me.txtFactorMontoS.Text, gstrMonedaLocal) > 0 And SacarFormatoValor(Me.txtFactorPorcentajeS.Text, "%") > 0 Then
    MsgBox "Sólo puede ingresar uno de los dos tipos de valores (Porcentaje o Monto)." & Chr(13) & "Intente nuevamente.", vbExclamation, "Maestro de Factores"
    Me.txtFactorMontoS.SetFocus
    Exit Sub
End If

mstrSql = ""
mstrSql = "UPDATE Srvt_Servicios SET "
mstrSql = mstrSql & "Factor_Porcentaje=" & SacarFormatoValor(Me.txtFactorPorcentajeS.Text, "%") & ", "
mstrSql = mstrSql & "Factor_Monto=" & SacarFormatoValor(Me.txtFactorMontoS.Text, gstrMonedaLocal) & " "
mstrSql = mstrSql & "WHERE Id_Servicio='" & Me.lsvServicios.SelectedItem.SubItems(2) & "'"

If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
    MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
End If

LLenaListaServicios

For ldblCont = 1 To Me.lsvServicios.ListItems.Count
    If Me.lsvServicios.ListItems(ldblCont).SubItems(1) = Me.lblServicio.Tag Then
        Me.lsvServicios.ListItems(ldblCont).Selected = True
        Me.lsvServicios.SetFocus
        Exit For
    End If
Next ldblCont

End Sub

Private Sub txtFactorPorcentajeS_GotFocus()
Me.txtFactorPorcentajeS.Text = SacarFormatoValor(Me.txtFactorPorcentajeS.Text, "%")
Me.txtFactorPorcentajeS.SelStart = 0
Me.txtFactorPorcentajeS.SelLength = Len(Me.txtFactorPorcentajeS.Text)
End Sub

Private Sub txtFactorPorcentajeS_LostFocus()
Me.txtFactorPorcentajeS.Text = FormatoValor(Me.txtFactorPorcentajeS.Text, "%", 2)
End Sub
