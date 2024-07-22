VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmRptReservaAten 
   Caption         =   "Reporte Reserva de Atención"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExportar 
      Caption         =   "&Exportar"
      Height          =   495
      Left            =   10560
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   720
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport crRptReservaHoras 
      Left            =   240
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataListLib.DataCombo dcEstados 
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton btnBuscar 
      Caption         =   "&Buscar"
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgReservas 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dpFecIniReserva 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   104660993
      CurrentDate     =   44691
   End
   Begin MSComCtl2.DTPicker dpFecFinReserva 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   104660993
      CurrentDate     =   44691
   End
   Begin VB.CommandButton btnImprimir 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   9480
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin MSDataListLib.DataCombo dcSucursal 
      Bindings        =   "frmRptReservaAten.frx":0000
      Height          =   315
      Left            =   5880
      TabIndex        =   9
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Descripcion"
      BoundColumn     =   "id_sucursal"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adosucursal 
      Height          =   330
      Left            =   6720
      Top             =   480
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      Caption         =   "Estado:"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Fin:"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicio:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Sucursal:"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmRptReservaAten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AdoRecordSucursal As New ADODB.Recordset
Dim mstrSQL As String

Private Sub btnBuscar_Click()

BuscarReservas

End Sub

Private Sub BuscarReservas()



Dim fecini As Date
Dim fecfin As Date
Dim estado As String
Dim Sucursal As String


fecini = dpFecIniReserva.Value
fecfin = dpFecFinReserva.Value
estado = dcEstados.BoundText
Sucursal = dcSucursal.BoundText


Dim cadfecini As String

Dim conex As New ADODB.Connection
conex.ConnectionString = "Provider='sqloledb'; Data Source='WIRACOCHA';Initial Catalog=Prueba;User Id=sa;Password=Llosa1936"
'conex.ConnectionString = "Provider='sqloledb'; Data Source='WIRACOCHA';Initial Catalog=Prueba;User Id=sa;Password=Llosa1936"
'Cambiar al pasar a PRO
conex.CursorLocation = adUseClient 'necesario para asignar recordset al datasource

If conex.State = 1 Then conex.Close

conex.Open


Dim cmd As New ADODB.Command

Dim rsReservas As New ADODB.Recordset

cmd.CommandType = adCmdStoredProc
cmd.CommandText = "usp_ListarReservaHora"
cmd.ActiveConnection = conex

cmd.Parameters.Append cmd.CreateParameter("@fecini", adDate, adParamInput, 4, fecini)
cmd.Parameters.Append cmd.CreateParameter("@fecfin", adDate, adParamInput, 4, fecfin)
cmd.Parameters.Append cmd.CreateParameter("@idempresa", adVarChar, adParamInput, 50, gstrIdEmpresa)
cmd.Parameters.Append cmd.CreateParameter("@idsucursal", adVarChar, adParamInput, 50, Sucursal)
cmd.Parameters.Append cmd.CreateParameter("@estado", adVarChar, adParamInput, 2, estado)


Set rsReservas = cmd.Execute

'al datagrid no hay que hacerle modificacion alguna al agregar al formulario
Set dgReservas.DataSource = rsReservas



End Sub

Private Sub btnExportar_Click()

 'Evalua si un datagrid esta vacio(empty)
 Dim rsReservaAux As New ADODB.Recordset
 Set rsReservaAux = dgReservas.DataSource
 
    If rsReservaAux.State = 0 Then
      MsgBox "No existen elementos en la lista, presione Buscar", vbExclamation, "Imprimir"
      Exit Sub
    End If

'Para el caso de que se haya posicionado el cursor al final
   If rsReservaAux.State <> 0 Then
        If rsReservaAux.EOF Then
            rsReservaAux.MoveFirst
        End If
    End If
    

Screen.MousePointer = vbHourglass


Dim cab As String
cab = ""

For i = 0 To dgReservas.Columns.Count - 1
    If dgReservas.Columns(i).Caption <> "" Then
        cab = cab & dgReservas.Columns(i).Caption & Chr(9)
    End If

Next i

Dim rs As Recordset
Set rs = dgReservas.DataSource


 ExportarDatosGrid cab, rs, Me.cdExportar, Me.hwnd

Screen.MousePointer = vbDefault

End Sub

Private Sub btnImprimir_Click()
ImprimirReservaHoras
End Sub

Private Sub ImprimirReservaHoras()

Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String

    'Devuelve la ruta del directorio Windows
    Dim rc As Long
    Dim WinPath As String
    WinPath = Space$(300)
    rc = GetWindowsDirectory(WinPath, 300)
    GcamBaseTem = Trim$(WinPath)
    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
    '---------------------------------------
    
     Dim rsReservaAux As New ADODB.Recordset
    Set rsReservaAux = dgReservas.DataSource
 
    If rsReservaAux.State = 0 Then
      MsgBox "No existen elementos en la lista, presione Buscar", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    
    If Dir(gstrPathReporte & "\BDNuevaReservaH.mdb") <> "" Then Kill gstrPathReporte & "\BDNuevaReservaH.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.

    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNuevaReservaH.mdb", dbLangGeneral) ' Crea a una base de datos nueva
'
'    gstrSql = "CREATE TABLE T_REPORTERESERVAH (Patente text,"
'    gstrSql = gstrSql & " Fecha_Emision date,"
'     gstrSql = gstrSql & " Fecha_Reserva date,"
'    gstrSql = gstrSql & " Hora_Reserva text,"
'    gstrSql = gstrSql & " Reparacion memo,"
'    gstrSql = gstrSql & " Fecha_Confirmacion date,"
'    gstrSql = gstrSql & " Quien_Confirma text,"
'    gstrSql = gstrSql & " Id_OT text,"
'    gstrSql = gstrSql & " Nombre text,"
'    gstrSql = gstrSql & " Vehiculo text,"
'    gstrSql = gstrSql & " Telefono text)"
'    Dbsnueva.Execute gstrSql



    gstrSql = "CREATE TABLE T_REPORTERESERVAH (Id_Reserva text,"
    gstrSql = gstrSql & " Sucursal text,"
    gstrSql = gstrSql & " Placa text,"
    gstrSql = gstrSql & " RealizadoPor text,"
    gstrSql = gstrSql & " Estado text,"
    gstrSql = gstrSql & " FechaInicio date,"
    gstrSql = gstrSql & " Fecha_Reserva date,"
    gstrSql = gstrSql & " Hora_Reserva text,"
    gstrSql = gstrSql & " Reparacion memo,"
    gstrSql = gstrSql & " Quien_Anula text,"
    gstrSql = gstrSql & " MotivoAnula text,"
    gstrSql = gstrSql & " Fecha_Activacion date,"
    gstrSql = gstrSql & " Quien_Activa text,"
    gstrSql = gstrSql & " Fecha_Confirmacion date,"
    gstrSql = gstrSql & " Quien_Confirma text,"
    gstrSql = gstrSql & " Fecha_Cancelacion date,"
    gstrSql = gstrSql & " MotivoCancela text,"
    gstrSql = gstrSql & " Quien_Cancela text,"
    gstrSql = gstrSql & " Id_OT text,"
    gstrSql = gstrSql & " Recepcionista text,"
    gstrSql = gstrSql & " Nombre text,"
    gstrSql = gstrSql & " Vehiculo text,"
    gstrSql = gstrSql & " Telefono text,"
    gstrSql = gstrSql & " Fecha_Ingreso date,"
    gstrSql = gstrSql & " Hora_Ingreso text)"
    Dbsnueva.Execute gstrSql
    
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_REPORTERESERVAH")
   
   
    
    If rsReservaAux.State <> 0 Then
        If rsReservaAux.EOF Then
            rsReservaAux.MoveFirst
        End If
       
    End If
    
    While Not rsReservaAux.EOF
        
        Tabla.AddNew
        Tabla!Id_Reserva = rsReservaAux!Id_Reserva
        Tabla!Sucursal = rsReservaAux!Sucursal
        Tabla!Placa = rsReservaAux!Placa
        Tabla!RealizadoPor = rsReservaAux!RealizadoPor
        Tabla!estado = rsReservaAux!estado
        Tabla!FechaInicio = rsReservaAux!FechaInicio
        Tabla!Fecha_Reserva = rsReservaAux!Fecha_Reserva
        Tabla!Hora_Reserva = rsReservaAux!Hora_Reserva
        Tabla!Reparacion = rsReservaAux!Reparacion
        Tabla!Quien_Anula = rsReservaAux!Quien_Anula
        Tabla!MotivoAnula = rsReservaAux!MotivoAnula
        Tabla!Fecha_Activacion = rsReservaAux!Fecha_Activacion
        Tabla!Quien_Activa = rsReservaAux!Quien_Activa
        
          If Not IsNull(rsReservaAux!Fecha_Confirmacion) Then
            Tabla!Fecha_Confirmacion = rsReservaAux!Fecha_Confirmacion
        End If
        
        Tabla!Fecha_Cancelacion = rsReservaAux!Fecha_Cancelacion
        Tabla!MotivoCancela = rsReservaAux!MotivoCancela
        Tabla!QUIEN_CAnCELA = rsReservaAux!QUIEN_CAnCELA
        Tabla!Id_OT = rsReservaAux!Id_OT
        Tabla!Recepcionista = rsReservaAux!Recepcionista
        Tabla!Nombre = rsReservaAux!Nombre
        Tabla!Vehiculo = rsReservaAux!Vehiculo
        Tabla!Telefono = rsReservaAux!Telefono
        Tabla!Fecha_Ingreso = rsReservaAux!Fecha_Ingreso
        Tabla!Hora_Ingreso = rsReservaAux!Hora_Ingreso
        
        
        Tabla.Update
        
        
        rsReservaAux.MoveNext
    Wend
    
   Tabla.Close
   Dbsnueva.Close
   With crRptReservaHoras
        .ReportFileName = gstrPathReporte & "\rptReservaHoras.rpt"
        .WindowTitle = "Reporte de Reserva de horas"
        
        
        .DataFiles(0) = gstrPathReporte & "\BDNuevaReservaH.mdb"
'        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
'        .Formulas(1) = "TITULO='LISTADO DE ORDENES DE COMPRA'"
'        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
'        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
'        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
'        .Formulas(5) = "Tdecimales=" & gintDecimalesMoneda
'        .Formulas(6) = "NombreIva='" & gstrNombreIva & "'"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = True
   End With
   
   
   Screen.MousePointer = 1



End Sub


Private Sub Form_Load()

CargarEstados
CargarSucursal
Inicializar


End Sub

Private Sub Inicializar()
    
    dpFecFinReserva.Value = DateAdd("d", 7, Now)
    dpFecIniReserva.Value = Now
    
'    Dim Mes As Integer
'    Dim mesCad As String
'    Dim anio As Integer
'
'    Mes = Month(Now)
'    anio = Year(Now)
'
'    mesCad = IIf(Mes < 10, "0" & Trim(Str(Mes)), Str(Mes))
'    dpFecIniReserva.Value = "01" & "/" & mesCad & "/" & Str(anio)
    

End Sub

Private Sub CargarSucursal()

'Llena sucursal
mstrSQL = "Select Id_Sucursal, Descripcion From Glbl_Sucursal Where Id_Empresa ='" + gstrIdEmpresa + "' Order by Descripcion"
 If Conexion.SendHost(mstrSQL, AdoRecordSucursal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    Set Me.adosucursal.Recordset = AdoRecordSucursal
 End If
 Me.dcSucursal.BoundText = gstrIdSucursal
 'Me.cmbsucursal.BoundText = gstrIdSucursal

End Sub

Private Sub CargarEstados()

'conex a bd para obtener listado de Estados
Dim conex As New ADODB.Connection
'conex.ConnectionString = "Provider='sqloledb'; Data Source='WIRACOCHA';Initial Catalog=Prueba;User Id=sa;Password=Llosa1936"
conex.ConnectionString = "Provider='sqloledb'; Data Source='WIRACOCHA';Initial Catalog=Prueba;User Id=sa;Password=Llosa1936"
conex.CursorLocation = adUseClient 'necesario para asignar recordset al datasource
If conex.State = 1 Then conex.Close
conex.Open

Dim cmd As New ADODB.Command
Dim rsEstadoListado As New ADODB.Recordset
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "usp_ListarEstados"
cmd.ActiveConnection = conex

Set rEstadoListado = cmd.Execute
dcEstados.Text = ""
dcEstados.BoundColumn = "Id_Estado"
dcEstados.ListField = "Descripcion"
Set dcEstados.RowSource = rEstadoListado

dcEstados.BoundText = "C"


End Sub
