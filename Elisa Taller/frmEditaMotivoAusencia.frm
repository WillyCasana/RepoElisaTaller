VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditaMotivoAusencia 
   Caption         =   "Edición Motivo de Ausencia"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   Icon            =   "frmEditaMotivoAusencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   5895
      Begin VB.ComboBox cboHoraHasta 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cboHoraDesde 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker pckFecha 
         Height          =   320
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83427329
         CurrentDate     =   37384
      End
      Begin VB.TextBox txtTotalHoras 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label7 
         Caption         =   "Horas"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin MSDataListLib.DataCombo dtcSucursal 
         Bindings        =   "frmEditaMotivoAusencia.frx":179A
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datSucursal 
         Height          =   330
         Left            =   2160
         Top             =   240
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
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
      Begin MSDataListLib.DataCombo dtcMotivo 
         Bindings        =   "frmEditaMotivoAusencia.frx":17B4
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datMotivos 
         Height          =   330
         Left            =   2280
         Top             =   720
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
      Begin MSDataListLib.DataCombo dtcSupervisor 
         Bindings        =   "frmEditaMotivoAusencia.frx":17CD
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc datSupervisor 
         Height          =   330
         Left            =   3360
         Top             =   1200
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
      Begin VB.Label Label3 
         Caption         =   "Mecanico   :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Motivo        :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal      :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   4200
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "frmEditaMotivoAusencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset
Dim Item As ListItem

Private Sub cboHoraDesde_LostFocus()
If Len(cboHoraDesde.Text) > 5 Or Len(cboHoraDesde) < 4 Then
    MsgBox "Formato de hora erroneo", vbExclamation, "Formato de Hora"
    Me.cboHoraDesde.SetFocus
    Exit Sub
End If
If Len(Me.cboHoraDesde.Text) = 4 Then
    Me.cboHoraDesde.Text = Mid(Me.cboHoraDesde.Text, 1, 2) & ":" & Mid(Me.cboHoraDesde.Text, 3, 2)
End If
If Not IsNumeric(Mid(Me.cboHoraDesde.Text, 1, 2)) Or Not IsNumeric(Mid(Me.cboHoraDesde.Text, 4, 2)) Then
    MsgBox "Formato de hora erroneo, ingrese solo numeros", vbExclamation, "Formato de Hora"
    Me.cboHoraDesde.SetFocus
    Exit Sub
End If
End Sub

Private Sub cboHoraHasta_LostFocus()
Dim ldblhoras As Double
Dim ldblminutos As Double

'valida
If Len(cboHoraHasta.Text) > 5 Or Len(cboHoraHasta) < 4 Then
    MsgBox "Formato de hora erroneo", vbExclamation, "Formato de Hora"
    Me.cboHoraHasta.SetFocus
    Exit Sub
End If
If Len(Me.cboHoraHasta.Text) = 4 Then
    Me.cboHoraHasta.Text = Mid(Me.cboHoraHasta.Text, 1, 2) & ":" & Mid(Me.cboHoraHasta.Text, 3, 2)
End If
If Not IsNumeric(Mid(Me.cboHoraHasta.Text, 1, 2)) Or Not IsNumeric(Mid(Me.cboHoraHasta.Text, 4, 2)) Then
    MsgBox "Formato de hora erroneo, ingrese solo numeros", vbExclamation, "Formato de Hora"
    Me.cboHoraHasta.SetFocus
    Exit Sub
End If

'calcula horas
If Val(Mid(Me.cboHoraHasta, 1, 2) & Mid(Me.cboHoraHasta, 4, 2)) > Val(Mid(Me.cboHoraDesde, 1, 2) & Mid(Me.cboHoraDesde, 4, 2)) Then '////horas
    ldblhoras = Val(Mid(Me.cboHoraHasta, 1, 2)) - Val(Mid(Me.cboHoraDesde, 1, 2))
    ldblminutos = Val(Mid(Me.cboHoraHasta, 4, 2)) - Val(Mid(Me.cboHoraDesde, 4, 2))
    Me.txtTotalHoras = FormatoValor(ldblhoras + (Abs(ldblminutos) / 60), "", 2)
Else
    MsgBox "La hora Hasta debe ser Mayor a la hora Desde", vbExclamation, "Formato Hora"
    Me.txtTotalHoras = ""
    Me.cboHoraHasta.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()

If Me.dtcSucursal.Text = "" Then
    MsgBox "La sucursal debe contener un valor...", vbInformation, "Advertencia"
    Me.dtcSucursal.SetFocus
    Exit Sub
End If
If Me.dtcMotivo.Text = "" Then
    MsgBox "El Motivo de Ausencia debe contener un valor...", vbInformation, "Advertencia"
    Me.dtcMotivo.SetFocus
    Exit Sub
End If
If Me.dtcSupervisor.Text = "" Then
    MsgBox "El Mecanico debe contener un valor...", vbInformation, "Advertencia"
    Me.dtcSupervisor.SetFocus
    Exit Sub
End If

If Me.cboHoraDesde = "" Then
    MsgBox "La Hora Desde debe contener un valor...", vbInformation, "Advertencia"
    Me.cboHoraDesde.SetFocus
    Exit Sub
End If
If Me.cboHoraHasta = "" Then
    MsgBox "La Hora Hasta debe contener un valor...", vbInformation, "Advertencia"
    Me.cboHoraHasta.SetFocus
    Exit Sub
End If
If Me.txtTotalHoras = "" Then
    MsgBox "El Total de Horas debe contener un valor...", vbInformation, "Advertencia"
    Me.txtTotalHoras.SetFocus
    Exit Sub
End If

'//Verifica si existe un registro...
If Me.Tag = "Crear" Then
    Dim AdoTemp As New ADODB.Recordset
    mstrSql = "select id_mecanico from Tllr_Mecanicos_Ausencias where Id_empresa='" & gstrIdEmpresa & "'"
    mstrSql = mstrSql & " And id_sucursal='" & Me.dtcSucursal.BoundText & "'"
    mstrSql = mstrSql & " And id_mecanico='" & Me.dtcSupervisor.BoundText & "'"
    mstrSql = mstrSql & " And Id_Ausencia='" & Me.dtcMotivo.BoundText & "'"
    mstrSql = mstrSql & " And Id_Fecha='" & Me.pckFecha & "'"
    mstrSql = mstrSql & " And Hora_Desde='" & Me.cboHoraDesde.Text & "'"
    mstrSql = mstrSql & " And Hora_Hasta='" & Me.cboHoraHasta.Text & "'"
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            MsgBox "Ya se encuentra asignada esta Ausencia ", vbInformation, "Advertencia"
            Me.dtcSupervisor.SetFocus
            Exit Sub
        End If
    End If
    Conexion.CloseHost AdoTemp
End If

DescargaDatos
GrabarRegistro
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If mblnSW Then
    CargaSucursal
    CargaMotivo
    FillMecanicos
    FillTime gintHoraInicio, gintHoratermino, cboHoraDesde
    FillTime gintHoraInicio, gintHoratermino, cboHoraHasta
    Me.pckFecha.Value = Date
    CargaDatos
    mblnSW = False
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
    End Select
End Sub

Private Sub Form_Load()
mblnSW = True
End Sub
Sub CargaDatos()
    If Me.Tag = "Crear" Then
        Me.dtcSucursal.BoundText = gstrIdSucursal
    Else
        Me.dtcSucursal.BoundText = frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(9)
        Me.dtcSucursal.Locked = True
        Me.dtcMotivo.BoundText = frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(10)
        Me.dtcMotivo.Locked = True
        Me.dtcSupervisor.BoundText = frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(2)
        Me.dtcSupervisor.Locked = True
        Me.pckFecha.Value = frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(4)
        Me.pckFecha.Enabled = False
        Me.cboHoraDesde = frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(5)
        Me.cboHoraHasta = frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(6)
        Me.txtTotalHoras = frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(7)
    End If
End Sub
Sub DescargaDatos()

With frmAusenciaMecanicos.lvwConceptos
    If Me.Tag = "Crear" Then
        Set Item = .ListItems.Add(, , Me.dtcSucursal.Text)
        Item.SubItems(1) = Me.dtcMotivo.Text
        Item.SubItems(2) = Me.dtcSupervisor.BoundText
        Item.SubItems(3) = Me.dtcSupervisor.Text
        Item.SubItems(4) = Me.pckFecha
        Item.SubItems(5) = Me.cboHoraDesde
        Item.SubItems(6) = Me.cboHoraHasta
        Item.SubItems(7) = FormatoValor(Me.txtTotalHoras, "", 2)
        
        Item.SubItems(8) = frmAusenciaMecanicos.lvwConceptos.ListItems.Count
        Item.SubItems(9) = Me.dtcSucursal.BoundText
        Item.SubItems(10) = Me.dtcMotivo.BoundText
    Else
        .SelectedItem.SubItems(4) = Me.pckFecha
        .SelectedItem.SubItems(5) = Me.cboHoraDesde
        .SelectedItem.SubItems(6) = Me.cboHoraHasta
        .SelectedItem.SubItems(7) = FormatoValor(Me.txtTotalHoras, "", 2)
    End If
End With
End Sub
Sub CargaSucursal()
mstrSql = "SELECT Id_Sucursal AS CODIGO, Descripcion AS NOMBRE FROM Glbl_Sucursal where VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datSucursal
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcSucursal.ListField = "Nombre"
            dtcSucursal.BoundColumn = "Codigo"
        End If
    End With
End If
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal
End Sub

Sub CargaMotivo()
mstrSql = "SELECT Id_Ausencia AS CODIGO, Descripcion AS NOMBRE FROM Tllr_Motivo_Ausencia where VIGENCIA = 'S'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datMotivos
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcMotivo.ListField = "Nombre"
            dtcMotivo.BoundColumn = "Codigo"
        End If
    End With
End If
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal
End Sub

Sub FillMecanicos()
gstrSql = "SELECT Id_Mecanico AS Codigo, Nombre FROM Tllr_Mecanicos where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and vigencia='S'"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
With datSupervisor
    Set .Recordset = gadoPrincipal
    If Not .Recordset.BOF And Not .Recordset.EOF Then
        .Recordset.MoveFirst
        dtcSupervisor.ListField = "Nombre"
        dtcSupervisor.BoundColumn = "Codigo"
    End If
End With
End If
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub

Private Sub GrabarRegistro()
    
    If Me.Tag = "Crear" Then
        mstrSql = "INSERT INTO TLLR_MECANICOS_AUSENCIAS (Id_Empresa,Id_Sucursal,Id_Mecanico,Id_Ausencia,Id_Item,Id_Fecha,Hora_Desde,Hora_Hasta,Total_Horas)"
        mstrSql = mstrSql & "values ('" & gstrIdEmpresa & "','" & Me.dtcSucursal.BoundText & "',"
        mstrSql = mstrSql & "'" & Me.dtcSupervisor.BoundText & "','" & Me.dtcMotivo.BoundText & "',"
        mstrSql = mstrSql & CorrelativoItem() & ","
        mstrSql = mstrSql & "'" & Me.pckFecha & "','" & Me.cboHoraDesde & "','" & Me.cboHoraHasta & "'," & Me.txtTotalHoras & ")"
    Else
        mstrSql = "UPDATE Tllr_Mecanicos_Ausencias SET "
        mstrSql = mstrSql & "Hora_Desde='" & Me.cboHoraDesde & "',"
        mstrSql = mstrSql & "Hora_Hasta='" & Me.cboHoraHasta & "',"
        mstrSql = mstrSql & "Total_Horas=" & CDbl(Me.txtTotalHoras)
        mstrSql = mstrSql & " where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & Me.dtcSucursal.BoundText & "'"
        mstrSql = mstrSql & " And Id_Mecanico='" & Me.dtcSupervisor.BoundText & "' And Id_Ausencia='" & Me.dtcMotivo.BoundText & "' And Id_Item=" & frmAusenciaMecanicos.lvwConceptos.SelectedItem.SubItems(8)
        mstrSql = mstrSql & " And Id_Fecha='" & Me.pckFecha & "'"
    End If
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        Me.Tag = ""
    End If
End Sub

Private Sub txtTotalHoras_GotFocus()
MarcaTexto txtTotalHoras
End Sub
Function CorrelativoItem() As Integer
Dim strSql As String
Dim AdoTemp As New ADODB.Recordset

    strSql = "Select max(Id_Item) as item From Tllr_Mecanicos_Ausencias where Id_Empresa='" & gstrIdEmpresa & "'"
    strSql = strSql & " And Id_Sucursal='" & Me.dtcSucursal.BoundText & "'"
    strSql = strSql & " And Id_Mecanico='" & Me.dtcSupervisor.BoundText & "'"
    strSql = strSql & " And Id_Ausencia='" & Me.dtcMotivo.BoundText & "'"
    strSql = strSql & " And Id_Fecha='" & Me.pckFecha & "'"
    If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            CorrelativoItem = IIf(IsNull(AdoTemp!Item), 1, AdoTemp!Item + 1)
        End If
    End If

End Function

