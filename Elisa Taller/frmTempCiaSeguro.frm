VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTempCiaSeguro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tempario de Concepto/Parte-Pieza -----> Compañia de Seguro"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "frmTempCiaSeguro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11370
   Begin VB.OptionButton optMuestraHoras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Horas"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   10305
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   210
      Width           =   975
   End
   Begin VB.OptionButton optMuestraValor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Valor"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   9330
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   210
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFlexGrid 
      Height          =   4860
      Left            =   90
      TabIndex        =   2
      Top             =   690
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   8573
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   2
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   2
      _Band(0).TextStyleHeader=   1
   End
   Begin MSDataListLib.DataCombo dtcCiaSeg 
      Bindings        =   "frmTempCiaSeguro.frx":179A
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc datCiaSeg 
      Height          =   330
      Left            =   2385
      Top             =   210
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Compañia de Seguro :"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   15
      Width           =   1575
   End
End
Attribute VB_Name = "frmTempCiaSeguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset
Dim mblnSW As Boolean

'/////////////////
Dim mstrFilObjeto As String
Dim mstrColObjeto As String

Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Dim mstrTipoObjeto As String


Function CodigoPieza(lngFil As Long, tpoField As tpoFields) As String
Dim lngAuxC As Long, lngAuxF As Long

With HFlexGrid
    If tpoField = tpoCodigo Then
        lngAuxF = .Row: lngAuxC = .Col
        .Col = 0: .Row = lngFil
        CodigoPieza = Trim(.Text)
        .Row = lngAuxF: .Col = lngAuxC
    End If
    If tpoField = tpoNombre Then
        lngAuxF = .Row: lngAuxC = .Col
        .Col = 1: .Row = lngFil
        CodigoPieza = Trim(.Text)
        .Row = lngAuxF: .Col = lngAuxC
    End If
End With
End Function


Function CodigoConcepto(lngCol As Long, tpoField As tpoFields) As String
Dim lngAuxC As Long, lngAuxF As Long


With HFlexGrid
    If tpoField = tpoCodigo Then
        lngAuxC = .Col: lngAuxF = .Row '////////////////GUARDO LAS COORDENADAS ORIGINALES
        .Row = 0: .Col = lngCol '////////////////ASIGNO LAS COORDENADAS PARCIALES
        CodigoConcepto = Trim(.Text)
        .Col = lngAuxC: .Row = lngAuxF '////////////////REESTAURO LAS COORDENADAS ORIGINALES
    End If
    If tpoField = tpoNombre Then
        lngAuxC = .Col: lngAuxF = .Row '////////////////GUARDO LAS COORDENADAS ORIGINALES
        .Row = 1: .Col = lngCol '////////////////ASIGNO LAS COORDENADAS PARCIALES
        CodigoConcepto = Trim(.Text)
        .Col = lngAuxC: .Row = lngAuxF '////////////////REESTAURO LAS COORDENADAS ORIGINALES
    End If
End With
End Function
Sub Genera(strCompañia As String)
Dim intRows As Integer, intCols As Integer
Dim intQ As Integer, intW As Integer
Dim strWhere As String

intRows = 0
intCols = 0

HFlexGrid.Rows = 3: HFlexGrid.FixedRows = 2: HFlexGrid.Cols = 3: HFlexGrid.FixedCols = 2
mstrSql = "SELECT Descripcion, Id_Parte_Pieza FROM Tllr_Parte_Pieza"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst: intRows = .RecordCount
            HFlexGrid.Rows = intRows + 2: HFlexGrid.FixedRows = 2
            intQ = 2
            While Not .EOF
                With HFlexGrid
                    .Row = intQ
                    .Col = 0: .Text = adoPrincipal!Id_Parte_Pieza: .ColWidth(0) = 5
                    .Col = 1: .Text = adoPrincipal!Descripcion: .ColWidth(1) = 3000
                End With
                intQ = intQ + 1
                .MoveNext
            Wend
        End If
    End With
End If
'/////////////////////AQUI SE LLENA LAS PARTES Y PIEZAS////////////ARRIBA
strWhere = IIf(strCompañia <> "", " WHERE Tllr_CiaSeguro_Concepto.Id_Compañia_Seguro = '" & strCompañia & "' ", "")
mstrSql = "SELECT Tllr_CiaSeguro_Concepto.Id_Concepto, Tllr_Concepto.Descripcion, Tllr_Concepto.D_P, Tllr_Concepto.Orden FROM Tllr_CiaSeguro_Concepto LEFT OUTER JOIN Tllr_Concepto ON Tllr_CiaSeguro_Concepto.Id_Concepto = Tllr_Concepto.Id_Concepto " & strWhere & " ORDER BY Tllr_Concepto.Orden"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst: intCols = .RecordCount
            HFlexGrid.Cols = intCols + 2: HFlexGrid.FixedCols = 2
            intQ = 2
            While Not .EOF
                With HFlexGrid
                    .Col = intQ
                    .Row = 0: .Text = adoPrincipal!Id_Concepto: .RowHeight(0) = 0
                    .Row = 1: .Text = adoPrincipal!Descripcion: .ColWidth(intQ) = Len(adoPrincipal!Descripcion) * 130
                End With
                intQ = intQ + 1
                .MoveNext
            Wend
        End If
    End With
End If

With HFlexGrid
    For intRows = 2 To .Rows - 1
        .Row = intRows
        For intCols = 2 To .Cols - 1
            .Col = intCols
            .Text = Valor_Hora(dtcCiaSeg.BoundText, CodigoConcepto(.Col, tpoCodigo), CodigoPieza(.Row, tpoCodigo), IIf(optMuestraValor.Value = True, "Valor", "Horas"))
        Next
    Next
End With

End Sub

Sub PartesYPiezas()
Dim intRows As Integer, intCols As Integer
Dim intQ As Integer, intW As Integer
Dim strWhere As String

intRows = 0
intCols = 0

HFlexGrid.Rows = 3: HFlexGrid.FixedRows = 2: HFlexGrid.Cols = 3: HFlexGrid.FixedCols = 2
mstrSql = "SELECT Descripcion, Id_Parte_Pieza FROM Tllr_Parte_Pieza"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst: intRows = .RecordCount
            HFlexGrid.Rows = intRows + 2: HFlexGrid.FixedRows = 2
            intQ = 2
            While Not .EOF
                With HFlexGrid
                    .Row = intQ
                    .Col = 0: .Text = adoPrincipal!Id_Parte_Pieza: .ColWidth(0) = 5
                    .Col = 1: .Text = adoPrincipal!Descripcion: .ColWidth(1) = 3000
                End With
                intQ = intQ + 1
                .MoveNext
            Wend
        End If
    End With
End If
End Sub

Function Valor_Hora(strCiaSeg As String, strConcepto As String, strPartePieza As String, strValorHora As String) As String

If strValorHora <> "" Then
    mstrSql = "SELECT " & strValorHora & " AS OBJETO FROM Tllr_CiaSeguro_Concepto_Parte_Pieza WHERE Id_Compañia_Seguro = '" & strCiaSeg & "' AND Id_Concepto = '" & strConcepto & "' AND Id_Parte_Pieza = '" & strPartePieza & "' ORDER BY Id_Compañia_Seguro, Id_Concepto, Id_Parte_Pieza"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                Valor_Hora = IIf(Not IsNull(!Objeto), CStr(!Objeto), "0")
            Else
                Valor_Hora = "0"
            End If
        End With
    End If
End If

End Function


Private Sub dtcCiaSeg_Change()
    If dtcCiaSeg.BoundText <> "" Then
        DoEvents
        HFlexGrid.Clear
        Genera dtcCiaSeg.BoundText
    End If
End Sub


Private Sub Form_Activate()
If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0080", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
    CompañiasDeSeguro
    mblnSW = False
End If
End Sub

Private Sub Form_Load()
mblnSW = True
'Genera ""
End Sub

Private Sub CompañiasDeSeguro()
    
    dtcCiaSeg.Enabled = True
    Set adoPrincipal = New ADODB.Recordset
    mstrSql = "SELECT Id_Compañia_Seguro as codigo, Nombre FROM Tllr_Compañia_Seguro where VIGENCIA = 'S' order by Nombre"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datCiaSeg
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcCiaSeg.ListField = "Nombre"
            dtcCiaSeg.BoundColumn = "Codigo"
'            dtcCiaSeg.BoundText = .Recordset!Codigo
'            If .Recordset.RecordCount < 2 Then dtcCiaSeg.Enabled = False
        End If
    End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
    
End Sub

Private Sub HFlexGrid_DblClick()
With frmEditTempCptoVsCiaSeg
    .lblCompañia.Tag = dtcCiaSeg.BoundText: .lblCompañia.Caption = dtcCiaSeg.Text
    .lblConcepto.Tag = CodigoConcepto(HFlexGrid.Col, tpoCodigo): .lblConcepto.Caption = CodigoConcepto(HFlexGrid.Col, tpoNombre)
    .lblPartePieza.Tag = CodigoPieza(HFlexGrid.Row, tpoCodigo): .lblPartePieza.Caption = CodigoPieza(HFlexGrid.Row, tpoNombre)
    .Label1(4).Caption = IIf(optMuestraValor.Value = True, "Valor", "Horas"): .txtValor = HFlexGrid.Text
    .Show 1
End With
End Sub

Private Sub optMuestraHoras_Click()
If optMuestraHoras.Value = True Then
    Genera dtcCiaSeg.BoundText
End If
End Sub

Private Sub optMuestraValor_Click()
If optMuestraValor.Value = True Then
    Genera dtcCiaSeg.BoundText
End If
End Sub


