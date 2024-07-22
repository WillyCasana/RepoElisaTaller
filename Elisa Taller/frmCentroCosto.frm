VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCentroCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centro de Costo"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCentroCosto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4140
      Begin MSAdodcLib.Adodc datCentroCosto 
         Height          =   270
         Left            =   1815
         Top             =   135
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "datCentroCosto"
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
      Begin MSDataListLib.DataCombo dbcboCentroCosto 
         Bindings        =   "frmCentroCosto.frx":038A
         Height          =   315
         Left            =   285
         TabIndex        =   2
         Top             =   420
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "CODIGO"
         Text            =   "dbcboCentroCosto"
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3285
      TabIndex        =   0
      Top             =   1215
      Width           =   1110
   End
End
Attribute VB_Name = "frmCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()

If Me.dbcboCentroCosto.BoundText <> "" Then
    gCentroCosto = Me.dbcboCentroCosto.BoundText
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String
Dim Tabla As New ADODB.Recordset

Sql = "SELECT ID_GRUPO_CENTRO_COSTO as CODIGO,NOMBRE AS DESCRIPCION FROM Cont_Grupo_Centro_Costo WHERE Id_empresa='" & gstrIdEmpresa & "' and  VIGENCIA='S'"
If Conexion.SendHost(Sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datCentroCosto.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
 gCentroCosto = Me.dbcboCentroCosto.BoundText
End Sub
