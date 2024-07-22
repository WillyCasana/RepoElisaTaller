VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmCajaCajero 
   Caption         =   "Seleccione Caja - Cajero"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaCajero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin MSDataListLib.DataCombo dtcCajero 
         Bindings        =   "frmCajaCajero.frx":038A
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo dtcCaja 
         Bindings        =   "frmCajaCajero.frx":03A2
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc datCaja 
         Height          =   330
         Left            =   1920
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSAdodcLib.Adodc datCajero 
         Height          =   375
         Left            =   2160
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Caption         =   "Cajero    :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Caja       :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCajaCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset

Private Sub cmdAceptar_Click()
If MsgBox("Está a punto de emitir facturación interna para las OT seleccionadas" & Chr(13) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, "Facturación Interna") = vbYes Then
    frmFacturarCargosInternos.Text1 = Me.dtcCaja.BoundText
    frmFacturarCargosInternos.Text2 = Me.dtcCajero.BoundText
    blnContinuar = True
Else
    frmFacturarCargosInternos.Text1 = ""
    frmFacturarCargosInternos.Text2 = ""
    blnContinuar = False
End If
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    frmFacturarCargosInternos.Text1 = ""
    frmFacturarCargosInternos.Text2 = ""
    Unload Me
End Sub

Private Sub dtcCaja_Change()
CargaCajeros Me.dtcCaja.BoundText
End Sub

Private Sub Form_Load()
    CargaCajas
End Sub
Sub CargaCajas()
    mstrSql = "Select Id_Caja as CODIGO, Descripcion as Nombre from Vpro_Caja where VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Descripcion"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With Me.datCaja
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                Me.dtcCaja.ListField = "Nombre"
                Me.dtcCaja.BoundColumn = "Codigo"
                Me.dtcCaja.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Sub CargaCajeros(strCaja As String)
    mstrSql = "Select Vpro_caja_cajero.id_cajero as codigo, vpro_Cajero.Descripcion as nombre"
    mstrSql = mstrSql & " from vpro_caja_cajero inner join vpro_cajero on vpro_caja_cajero.Id_Cajero=Vpro_cajero.Id_Cajero"
    mstrSql = mstrSql & " Where Id_Caja='" & strCaja & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With Me.datCajero
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                Me.dtcCajero.ListField = "Nombre"
                Me.dtcCajero.BoundColumn = "Codigo"
                Me.dtcCajero.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

