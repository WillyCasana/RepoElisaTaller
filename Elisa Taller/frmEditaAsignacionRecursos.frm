VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditaAsignacionRecursos 
   Caption         =   "Edición Asignación de Turnos"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   Icon            =   "frmEditaAsignacionRecursos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5895
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103219201
         CurrentDate     =   37382
      End
      Begin MSComCtl2.DTPicker pckFechadesde 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   46596097
         CurrentDate     =   37382
      End
      Begin MSDataListLib.DataCombo dtcSupervisor 
         Bindings        =   "frmEditaAsignacionRecursos.frx":179A
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   3330
         _ExtentX        =   5874
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
         Left            =   3480
         Top             =   240
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
      Begin VB.Label Label5 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Mecanico"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin MSDataListLib.DataCombo dtcSucursal 
         Bindings        =   "frmEditaAsignacionRecursos.frx":17B6
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
      Begin MSDataListLib.DataCombo dtcTurnos 
         Bindings        =   "frmEditaAsignacionRecursos.frx":17D0
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
      Begin MSAdodcLib.Adodc datTurnos 
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
      Begin VB.Label Label2 
         Caption         =   "Turnos     :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal   :"
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
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   4200
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
End
Attribute VB_Name = "frmEditaAsignacionRecursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim mstrSQL As String
Dim AdoPrincipal As New ADODB.Recordset
Dim Item As ListItem

Private Sub cmdAceptar_Click()

If Me.dtcSucursal.Text = "" Then
    MsgBox "La sucursal debe contener un valor...", vbInformation, "Advertencia"
    Me.dtcSucursal.SetFocus
    Exit Sub
End If
If Me.dtcTurnos.Text = "" Then
    MsgBox "El Turnos debe contener un valor...", vbInformation, "Advertencia"
    Me.dtcTurnos.SetFocus
    Exit Sub
End If
If Me.dtcSupervisor.Text = "" Then
    MsgBox "El Mecanico debe contener un valor...", vbInformation, "Advertencia"
    Me.dtcSupervisor.SetFocus
    Exit Sub
End If

'//Verifica si existe un registro...
If Me.Tag = "Crear" Then
    Dim adoTemp As New ADODB.Recordset
    mstrSQL = "select id_mecanico from Tllr_Mecanicos_Turnos where Id_empresa='" & gstrIdEmpresa & "'"
    mstrSQL = mstrSQL & " And id_sucursal='" & Me.dtcSucursal.BoundText & "'"
    mstrSQL = mstrSQL & " And id_mecanico='" & Me.dtcSupervisor.BoundText & "'"
    mstrSQL = mstrSQL & " And Id_turno='" & Me.dtcTurnos.BoundText & "'"
    mstrSQL = mstrSQL & " And Fecha_Desde='" & Me.pckFechaDesde & "'"
    mstrSQL = mstrSQL & " And Fecha_Hasta='" & Me.pckFechaHasta & "'"
    If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            MsgBox "Ya se encuentra asignado este Turno en estas fechas ", vbInformation, "Advertencia"
            Me.dtcSupervisor.SetFocus
            Exit Sub
        End If
    End If
    Conexion.CloseHost adoTemp
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
    CargaTurnos
    FillMecanicos
    Me.pckFechaDesde.Value = BOM(Date)
    Me.pckFechaHasta.Value = EOM(Date)
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
        Me.dtcSucursal.BoundText = frmAsignacionTurnos.lvwConceptos.SelectedItem.SubItems(7)
        Me.dtcSucursal.Locked = True
        Me.dtcTurnos.BoundText = frmAsignacionTurnos.lvwConceptos.SelectedItem.SubItems(8)
        Me.dtcTurnos.Locked = True
        Me.dtcSupervisor.Text = frmAsignacionTurnos.lvwConceptos.SelectedItem.SubItems(3)
        Me.dtcSupervisor.Locked = True
        Me.pckFechaDesde.Value = frmAsignacionTurnos.lvwConceptos.SelectedItem.SubItems(4)
        Me.pckFechaHasta.Value = frmAsignacionTurnos.lvwConceptos.SelectedItem.SubItems(5)

    End If
End Sub
Sub DescargaDatos()

With frmAsignacionTurnos.lvwConceptos
    If Me.Tag = "Crear" Then
        Set Item = .ListItems.Add(, , Me.dtcSucursal.Text)
        Item.SubItems(1) = Me.dtcTurnos.Text
        Item.SubItems(2) = Me.dtcSupervisor.BoundText
        Item.SubItems(3) = Me.dtcSupervisor.Text
        Item.SubItems(4) = Me.pckFechaDesde
        Item.SubItems(5) = Me.pckFechaHasta
        Item.SubItems(6) = frmAsignacionTurnos.lvwConceptos.ListItems.Count
        Item.SubItems(7) = Me.dtcSucursal.BoundText
        Item.SubItems(8) = Me.dtcTurnos.BoundText
    Else
        .SelectedItem.SubItems(4) = Me.pckFechaDesde
        .SelectedItem.SubItems(5) = Me.pckFechaHasta
    End If
End With
End Sub
Sub CargaSucursal()
mstrSQL = "SELECT Id_Sucursal AS CODIGO, Descripcion AS NOMBRE FROM Glbl_Sucursal where TieneTaller='S' and VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datSucursal
        Set .Recordset = AdoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcSucursal.ListField = "Nombre"
            dtcSucursal.BoundColumn = "Codigo"
        End If
    End With
End If
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal
End Sub

Sub CargaTurnos()
mstrSQL = "SELECT Id_Turno AS CODIGO, Descripcion AS NOMBRE FROM Tllr_Turnos where VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datTurnos
        Set .Recordset = AdoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcTurnos.ListField = "Nombre"
            dtcTurnos.BoundColumn = "Codigo"
        End If
    End With
End If
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal
End Sub

Sub FillMecanicos()
gstrSql = "SELECT Id_Mecanico AS Codigo, Nombre FROM Tllr_Mecanicos where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and vigencia='S'  AND (Es_Recepcionista = 'N') AND (Es_Supervisor = 'N') AND (Es_Liquidador = 'N') AND Nombre not like '%definir%' order by Nombre "

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
        mstrSQL = "INSERT INTO TLLR_MECANICOS_TURNOS (Id_Empresa,Id_Sucursal,Id_Mecanico,Id_Turno,Id_Item,Fecha_Desde,Fecha_Hasta,"
        mstrSQL = mstrSQL & "usr_id, usr_fecha) "
        mstrSQL = mstrSQL & "values ('" & gstrIdEmpresa & "','" & Me.dtcSucursal.BoundText & "',"
        mstrSQL = mstrSQL & "'" & Me.dtcSupervisor.BoundText & "','" & Me.dtcTurnos.BoundText & "',"
        mstrSQL = mstrSQL & CorrelativoItem() & ","
        mstrSQL = mstrSQL & "'" & Me.pckFechaDesde & "','" & Me.pckFechaHasta & "',"
        mstrSQL = mstrSQL & "'" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "')"
    Else
        mstrSQL = "UPDATE Tllr_Mecanicos_Turnos SET "
        mstrSQL = mstrSQL & "Fecha_Desde='" & Me.pckFechaDesde & "',"
        mstrSQL = mstrSQL & "Fecha_Hasta='" & Me.pckFechaHasta & "',"
        mstrSQL = mstrSQL & "usr_id='" & gstrUsuario & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "'"
        mstrSQL = mstrSQL & " where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & Me.dtcSucursal.BoundText & "'"
        mstrSQL = mstrSQL & " And Id_Mecanico='" & Me.dtcSupervisor.BoundText & "' And Id_Turno='" & Me.dtcTurnos.BoundText & "' And Id_Item=" & frmAsignacionTurnos.lvwConceptos.SelectedItem.SubItems(6)
    End If
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        Me.Tag = ""
    End If
End Sub

Function CorrelativoItem() As Integer
Dim strSql As String
Dim adoTemp As New ADODB.Recordset

    strSql = "Select max(Id_Item) as item From Tllr_Mecanicos_Turnos where Id_Empresa='" & gstrIdEmpresa & "'"
    strSql = strSql & " And Id_Sucursal='" & Me.dtcSucursal.BoundText & "'"
    strSql = strSql & " And Id_Mecanico='" & Me.dtcSupervisor.BoundText & "'"
    strSql = strSql & " And Id_Turno='" & Me.dtcTurnos.BoundText & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            CorrelativoItem = IIf(IsNull(adoTemp!Item), 1, adoTemp!Item + 1)
        End If
    End If

End Function
