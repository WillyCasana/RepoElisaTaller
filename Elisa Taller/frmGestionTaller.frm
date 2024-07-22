VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGestionTaller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Taller"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmGestionTaller.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5640
      TabIndex        =   32
      Top             =   4920
      Width           =   975
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frmGestionTaller.frx":179A
      TabIndex        =   31
      Top             =   5400
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdCalcular 
      Appearance      =   0  'Flat
      Caption         =   "C&alcular"
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Utilización Mano de Obra"
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   7455
      Begin VB.TextBox txtHorasCompradasMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2760
         TabIndex        =   28
         Text            =   "0"
         Top             =   680
         Width           =   1215
      End
      Begin VB.TextBox txtHorasRealesMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2760
         TabIndex        =   25
         Text            =   "0"
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   35
         Top             =   550
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "Resultado"
         Height          =   255
         Left            =   5160
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblResultadoMO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5160
         TabIndex        =   26
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line3 
         X1              =   2520
         X2              =   4200
         Y1              =   620
         Y2              =   620
      End
      Begin VB.Label Label13 
         Caption         =   "Horas Compradas     :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Horas Reales            :"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Eficiencia"
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   7455
      Begin VB.TextBox txtHorasRealesE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2760
         TabIndex        =   36
         Text            =   "0"
         Top             =   680
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   34
         Top             =   550
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Resultado"
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblResultadoEfi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5160
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblHorasRealesE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2760
         TabIndex        =   20
         Top             =   200
         Width           =   1200
      End
      Begin VB.Line Line2 
         X1              =   2520
         X2              =   4200
         Y1              =   620
         Y2              =   620
      End
      Begin VB.Label Label9 
         Caption         =   "Horas Reales Facturadas   :"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Horas Reales                     :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Productividad"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   7455
      Begin VB.TextBox txtHorasCompradasP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2760
         TabIndex        =   15
         Text            =   "0"
         Top             =   680
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   33
         Top             =   550
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Resultado"
         Height          =   255
         Left            =   5160
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblResultadoProd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5160
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   4200
         Y1              =   620
         Y2              =   620
      End
      Begin VB.Label lblHorasFacturadas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2760
         TabIndex        =   14
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Horas Compradas               :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Horas Reales Facturadas   :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Label lblFechaHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   6120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblFechaDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   6120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblSucursal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblMecanico 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Hasta      :"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Desde     :"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal      :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Mecánico    :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmGestionTaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW  As Boolean
Dim mstrSQL As String

Private Sub cmdCalcular_Click()

If CDbl(txtHorasCompradasP) = 0 Or CDbl(Me.txtHorasRealesE) = 0 Or CDbl(txtHorasCompradasMO) = 0 Then
    MsgBox "No se puede dividir por 0..., verifique los valores", vbExclamation, "División por 0"
    Exit Sub
End If

lblResultadoProd = FormatoValor((CDbl(lblHorasFacturadas) / CDbl(txtHorasCompradasP)) * 100, "", 2) & " "
lblResultadoEfi = FormatoValor((CDbl(lblHorasRealesE) / CDbl(txtHorasRealesE)) * 100, "", 2) & " "
lblResultadoMO = FormatoValor((CDbl(txtHorasRealesMO) / CDbl(txtHorasCompradasMO)) * 100, "", 2) & " "

MSChart1.Column = 1
MSChart1.Row = 1
MSChart1.Data = Me.lblResultadoProd

MSChart1.Column = 1
MSChart1.Row = 2
MSChart1.Data = Me.lblResultadoEfi

MSChart1.Column = 1
MSChart1.Row = 3
MSChart1.Data = Me.lblResultadoMO


End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdImprimir_Click()
cmdCalcular.Visible = False
cmdCancelar.Visible = False
cmdImprimir.Visible = False
frmGestionTaller.PrintForm
cmdCalcular.Visible = True
cmdCancelar.Visible = True
cmdImprimir.Visible = True

End Sub

Private Sub Form_Activate()
If mblnSW = True Then
    lblFechaDesde = frmInfProdMec.pckFechaDesde
    lblFechaHasta = frmInfProdMec.pckFechaHasta
    lblSucursal = gstrSucursal
    lblMecanico = IIf(frmInfProdMec.dtcSupervisor.Text = "", "TALLER", frmInfProdMec.dtcSupervisor.Text)
    CalculaHorasTaller frmInfProdMec.dtcSupervisor.BoundText
End If
End Sub

Private Sub Form_Load()
mblnSW = True

End Sub
Sub CalculaHorasTaller(lstrCodigoMecanico As String)
Dim TotalHoras As Double
Dim TotalAsignadas As Double
Dim TotalRealFacturado As Double

If lstrCodigoMecanico = "" Then     'taller completo
    'Horas otro y mecanica
    
'    mstrSql = "SELECT SUM(Tllr_Otro_OT.Horas) AS TOTALHORASOTRO, SUM(Tllr_Otro_Ot.Subtotal) as TotalRealFacturadoO "
'    mstrSql = mstrSql & "FROM Tllr_Otro_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_Facturacion ON Tllr_Otro_OT.Id_Empresa = Tllr_Facturacion.Id_Empresa AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Sucursal = Tllr_Facturacion.Id_Sucursal AND Tllr_Otro_OT.Id_OT = Tllr_Facturacion.Id_OT AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Tipo_Cargo = Tllr_Facturacion.Id_Cargo AND Tllr_Otro_OT.Seccion_OT = Tllr_Facturacion.Seccion_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_OT ON Tllr_Otro_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Otro_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_OT = Tllr_OT.Id_OT And Tllr_Otro_OT.Seccion_OT = Tllr_OT.Seccion_OT "
'    mstrSql = mstrSql & "WHERE Tllr_Otro_OT.Id_Empresa = '" & gstrIdEmpresa & "' AND Tllr_Otro_OT.Id_Sucursal = '" & gstrIdSucursal & "' And ((Tllr_OT.Fecha_Liquidacion Between '" & Me.lblFechaDesde & "'  And  '" & Me.lblFechaHasta & "') or ( Tllr_Facturacion.Fecha_Facturacion Between '" & Me.lblFechaDesde & "' And '" & Me.lblFechaHasta & "')) AND Tllr_OT.Estado " & gstrEstadoProdMecanico 'IN('L','B','F','C')"
'    If Conexion.SendHost(mstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
'        With gadoPrincipal
'            If Not .BOF And Not .EOF Then
'                TotalHoras = IIf(IsNull(!TOTALHORASOTRO), 0, !TOTALHORASOTRO)
'                TotalRealFacturado = IIf(IsNull(!totalrealfacturadoO), 0, !totalrealfacturadoO)
'            End If
'            .Close
'        End With
'    End If
'
'    'horas mecanica
'    mstrSql = "SELECT SUM(Tllr_Mecanica_OT.Horas) AS TOTALHORASMECANICA, SUM(Tllr_Mecanica_Ot.Subtotal) as TotalRealFacturadoM "
'    mstrSql = mstrSql & "FROM Tllr_Mecanica_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_Facturacion ON Tllr_Mecanica_OT.Id_Empresa = Tllr_Facturacion.Id_Empresa AND "
'    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_Sucursal = Tllr_Facturacion.Id_Sucursal AND Tllr_Mecanica_OT.Id_OT = Tllr_Facturacion.Id_OT AND "
'    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_Tipo_Cargo = Tllr_Facturacion.Id_Cargo AND Tllr_Mecanica_OT.Seccion_OT = Tllr_Facturacion.Seccion_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_OT ON Tllr_Mecanica_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Mecanica_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
'    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_OT = Tllr_OT.Id_OT And Tllr_Mecanica_OT.Seccion_OT = Tllr_OT.Seccion_OT "
'    mstrSql = mstrSql & "WHERE Tllr_Mecanica_OT.Id_Empresa = '" & gstrIdEmpresa & "' AND Tllr_Mecanica_OT.Id_Sucursal = '" & gstrIdSucursal & "' And ((Tllr_OT.Fecha_Liquidacion Between '" & Me.lblFechaDesde & "'  And  '" & Me.lblFechaHasta & "') or ( Tllr_Facturacion.Fecha_Facturacion Between '" & Me.lblFechaDesde & "' And '" & Me.lblFechaHasta & "')) AND Tllr_OT.Estado " & gstrEstadoProdMecanico 'IN('L','B','F','C')"
'    If Conexion.SendHost(mstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
'        With gadoPrincipal
'            If Not .BOF And Not .EOF Then
'                TotalHoras = TotalHoras + IIf(IsNull(!TotalHorasMecanica), 0, !TotalHorasMecanica)
'                TotalRealFacturado = TotalRealFacturado + IIf(IsNull(!TotalRealFacturadoM), 0, !TotalRealFacturadoM)
'            End If
'            .Close
'        End With
'    End If
    
    mstrSQL = "Exec Tllr_HorasFacturadas '" & gstrIdEmpresa & "','" & gstrIdSucursal & "','','" & Me.lblFechaDesde & "','" & Me.lblFechaHasta & "','" & gstrEstadoProdMecanico & "'"
    If Conexion.SendHost(mstrSQL, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
        With gadoPrincipal
            If Not .BOF And Not .EOF Then
                TotalHoras = IIf(IsNull(!TotalHorasOtro), 0, !TotalHorasOtro) + IIf(IsNull(!TotalHorasMecanica), 0, !TotalHorasMecanica)
                TotalAsignadas = IIf(IsNull(!TotalAsignadasOtro), 0, !TotalAsignadasOtro) + IIf(IsNull(!TotalHorasMecanica), 0, !TotalHorasMecanica)
                TotalRealFacturado = IIf(IsNull(!totalrealfacturadoO), 0, !totalrealfacturadoO) + IIf(IsNull(!TotalRealFacturadoM), 0, !TotalRealFacturadoM)
            End If
            .Close
        End With
    End If
    
'    mstrSql = "Select SUM(horascompradas) as TotalHrCompradas, SUM(horasreales) as TotalHrReales from tllr_mes_año_mecanico "
'    mstrSql = mstrSql & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Mes='" & Format(CStr(Month(Me.lblFechaHasta)), "00") & "' And Año=" & Year(Me.lblFechaHasta)
    
    mstrSQL = "Select SUM(isnull(Horas_Compradas,0)) as TotalHrCompradas from tllr_Hoja_Recursos "
    mstrSQL = mstrSQL & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Fecha Between '" & Me.lblFechaDesde & "' And '" & Me.lblFechaHasta & "'"
    If Conexion.SendHost(mstrSQL, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
        With gadoPrincipal
            If Not .BOF And Not .EOF Then
                txtHorasCompradasP = FormatoValor(IIf(IsNull(!TotalHrCompradas), 0, !TotalHrCompradas), "", 2)
                txtHorasCompradasMO = FormatoValor(IIf(IsNull(!TotalHrCompradas), 0, !TotalHrCompradas), "", 2)
                'txtHorasRealesMO = FormatoValor(IIf(IsNull(!TotalHrReales), 0, !TotalHrReales), "", 2)
            End If
            .Close
        End With
    End If
    
    'busca horas reales del taller
    txtHorasCompradasP = frmInfProdMec.lblHorasAsignadas
    txtHorasCompradasMO = frmInfProdMec.lblHorasAsignadas
    Me.txtHorasRealesMO = FormatoValor(TraeHorasRealesTaller, "", 2)
    lblHorasFacturadas = FormatoValor(TotalRealFacturado / gcurPrecioManoObra, "", 2) & "  "   'FormatoValor(TotalHoras, "", 2) & "  "
    txtHorasRealesE = txtHorasRealesMO   'FormatoValor(TotalHoras, "", 2) & "  "
    lblHorasRealesE = FormatoValor(TotalRealFacturado / gcurPrecioManoObra, "", 2) & "  "
    
Else        'por mecanico

'    mstrSql = "SELECT SUM(Tllr_Otro_OT.Horas) AS TOTALHORASOTRO, SUM(Tllr_Otro_Ot.Subtotal) as TotalRealFacturadoO "
'    mstrSql = mstrSql & "FROM Tllr_Otro_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_Facturacion ON Tllr_Otro_OT.Id_Empresa = Tllr_Facturacion.Id_Empresa AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Sucursal = Tllr_Facturacion.Id_Sucursal AND Tllr_Otro_OT.Id_OT = Tllr_Facturacion.Id_OT AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Tipo_Cargo = Tllr_Facturacion.Id_Cargo AND Tllr_Otro_OT.Seccion_OT = Tllr_Facturacion.Seccion_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_OT ON Tllr_Otro_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Otro_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_OT = Tllr_OT.Id_OT And Tllr_Otro_OT.Seccion_OT = Tllr_OT.Seccion_OT "
'    mstrSql = mstrSql & "WHERE Tllr_Otro_OT.Id_Empresa = '" & gstrIdEmpresa & "' AND Tllr_Otro_OT.Id_Sucursal = '" & gstrIdSucursal & "' And Tllr_Otro_Ot.Mecanico_Asignado='" & lstrCodigoMecanico & "' And ((Tllr_OT.Fecha_Liquidacion Between '" & Me.lblFechaDesde & "'  And  '" & Me.lblFechaHasta & "') or ( Tllr_Facturacion.Fecha_Facturacion Between '" & Me.lblFechaDesde & "' And '" & Me.lblFechaHasta & "')) AND Tllr_OT.Estado " & gstrEstadoProdMecanico 'IN('L','B','F','C')"
'    If Conexion.SendHost(mstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
'        With gadoPrincipal
'            If Not .BOF And Not .EOF Then
'                TotalHoras = IIf(IsNull(!TOTALHORASOTRO), 0, !TOTALHORASOTRO)
'                TotalRealFacturado = IIf(IsNull(!totalrealfacturadoO), 0, !totalrealfacturadoO)
'            End If
'            .Close
'        End With
'    End If
'
'    'horas mecanica
'    mstrSql = "SELECT SUM(Tllr_Mecanica_OT.Horas) AS TOTALHORASMECANICA, SUM(Tllr_Mecanica_Ot.Subtotal) as TotalRealFacturadoM "
'    mstrSql = mstrSql & "FROM Tllr_Mecanica_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_Facturacion ON Tllr_Mecanica_OT.Id_Empresa = Tllr_Facturacion.Id_Empresa AND "
'    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_Sucursal = Tllr_Facturacion.Id_Sucursal AND Tllr_Mecanica_OT.Id_OT = Tllr_Facturacion.Id_OT AND "
'    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_Tipo_Cargo = Tllr_Facturacion.Id_Cargo AND Tllr_Mecanica_OT.Seccion_OT = Tllr_Facturacion.Seccion_OT INNER JOIN "
'    mstrSql = mstrSql & "Tllr_OT ON Tllr_Mecanica_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Mecanica_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
'    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_OT = Tllr_OT.Id_OT And Tllr_Mecanica_OT.Seccion_OT = Tllr_OT.Seccion_OT "
'    mstrSql = mstrSql & "WHERE Tllr_Mecanica_OT.Id_Empresa = '" & gstrIdEmpresa & "' AND Tllr_Mecanica_OT.Id_Sucursal = '" & gstrIdSucursal & "' And Tllr_Mecanica_Ot.Mecanico_Designado = '" & lstrCodigoMecanico & "' And ((Tllr_OT.Fecha_Liquidacion Between '" & Me.lblFechaDesde & "'  And  '" & Me.lblFechaHasta & "') or ( Tllr_Facturacion.Fecha_Facturacion Between '" & Me.lblFechaDesde & "' And '" & Me.lblFechaHasta & "')) AND Tllr_OT.Estado " & gstrEstadoProdMecanico 'IN('L','B','F','C')"
'
'    If Conexion.SendHost(mstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
'        With gadoPrincipal
'            If Not .BOF And Not .EOF Then
'                TotalHoras = TotalHoras + IIf(IsNull(!TotalHorasMecanica), 0, !TotalHorasMecanica)
'                TotalRealFacturado = TotalRealFacturado + IIf(IsNull(!TotalRealFacturadoM), 0, !TotalRealFacturadoM)
'            End If
'            .Close
'        End With
'    End If
    
    mstrSQL = "Exec Tllr_HorasFacturadas_Mecanico '" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & lstrCodigoMecanico & "','" & Me.lblFechaDesde & "','" & Me.lblFechaHasta & "','" & gstrEstadoProdMecanico & "'"
    If Conexion.SendHost(mstrSQL, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
        With gadoPrincipal
            If Not .BOF And Not .EOF Then
                TotalHoras = IIf(IsNull(!TotalHorasOtro), 0, !TotalHorasOtro) + IIf(IsNull(!TotalHorasMecanica), 0, !TotalHorasMecanica)
                TotalAsignadas = IIf(IsNull(!TotalAsignadasOtro), 0, !TotalAsignadasOtro) + IIf(IsNull(!TotalHorasMecanica), 0, !TotalHorasMecanica)
                TotalRealFacturado = IIf(IsNull(!totalrealfacturadoO), 0, !totalrealfacturadoO) + IIf(IsNull(!TotalRealFacturadoM), 0, !TotalRealFacturadoM)
            End If
            .Close
        End With
    End If
    
'    mstrSql = "Select SUM(horascompradas) as TotalHrCompradas, SUM(horasreales) as TotalHrReales from tllr_mes_año_mecanico "
'    mstrSql = mstrSql & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Mes='" & Format(CStr(Month(Me.lblFechaHasta)), "00") & "' And Año=" & Year(Me.lblFechaHasta) & " And Id_Mecanico='" & lstrCodigoMecanico & "'"
    
    mstrSQL = "Select SUM(isnull(Horas_Compradas,0)) as TotalHrCompradas from tllr_Hoja_Recursos "
    mstrSQL = mstrSQL & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Fecha Between '" & Me.lblFechaDesde & "' And '" & Me.lblFechaHasta & "' And Id_Mecanico='" & lstrCodigoMecanico & "'"
    If Conexion.SendHost(mstrSQL, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
        With gadoPrincipal
            If Not .BOF And Not .EOF Then
                txtHorasCompradasP = FormatoValor(IIf(IsNull(!TotalHrCompradas), 0, !TotalHrCompradas), "", 2)
                txtHorasCompradasMO = FormatoValor(IIf(IsNull(!TotalHrCompradas), 0, !TotalHrCompradas), "", 2)
                'txtHorasRealesMO = FormatoValor(IIf(IsNull(!TotalHrReales), 0, !TotalHrReales), "", 2)
            End If
            .Close
        End With
    End If
    txtHorasCompradasP = frmInfProdMec.lblHorasAsignadas
    txtHorasCompradasMO = frmInfProdMec.lblHorasAsignadas
    txtHorasRealesMO = FormatoValor(CDbl(frmInfProdMec.lblTotalMec2) + CDbl(frmInfProdMec.lblTotalOtro2), "", 2)
    lblHorasFacturadas = FormatoValor(TotalRealFacturado / gcurPrecioManoObra, "", 2) & "  "   'FormatoValor(TotalHoras, "", 2) & "  "
    txtHorasRealesE = txtHorasRealesMO   'FormatoValor(TotalHoras, "", 2) & "  "
    lblHorasRealesE = FormatoValor(TotalRealFacturado / gcurPrecioManoObra, "", 2) & "  "
'    If frmInfProdMec.lblTotHorEst = "0" Then
'        txtHorasCompradasP = FormatoValor(NroDiasHabiles(CDate(Me.lblFechaDesde), CDate(Me.lblFechaHasta) & " 23:59:59") * gdblNroHorOblg, "", 0)
'    Else
'        txtHorasCompradasP = frmInfProdMec.lblTotHorEst
'    End If
End If
End Sub

Private Sub txtHorasCompradasP_Change()
'txtHorasCompradasMO = txtHorasCompradasP
End Sub

Private Sub txtHorasCompradasP_GotFocus()
txtHorasCompradasP = SacarFormatoValor(txtHorasCompradasP, "")
MarcaTexto txtHorasCompradasP
End Sub

Private Sub txtHorasCompradasP_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasCompradasP, strDot)
End Sub

Private Sub txtHorasCompradasP_LostFocus()
txtHorasCompradasP = FormatoValor(txtHorasCompradasP, "", 2)
txtHorasCompradasMO = txtHorasCompradasP
End Sub

Private Sub txtHorasRealesE_GotFocus()
txtHorasRealesE = SacarFormatoValor(txtHorasRealesE, "")
MarcaTexto txtHorasRealesE
End Sub

Private Sub txtHorasRealesE_LostFocus()
txtHorasRealesE = FormatoValor(txtHorasRealesE, "", 2)
End Sub

Private Sub txtHorasRealesMO_GotFocus()
txtHorasRealesMO = SacarFormatoValor(txtHorasRealesMO, "")
MarcaTexto txtHorasRealesMO
End Sub

Private Sub txtHorasRealesMO_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasRealesMO, strDot)
End Sub

Private Sub txtHorasRealesMO_LostFocus()
txtHorasRealesMO = FormatoValor(txtHorasRealesMO, "", 2)
End Sub

Function TraeHorasRealesTaller() As Double
Dim dblHorasRealesMecanica As Double
Dim dblHorasRealesOtro As Double

gstrSql = "SELECT SUM(Tllr_Mecanica_OT.HorasReales) AS TOTALHORASMECANICA FROM Tllr_Facturacion"
gstrSql = gstrSql & " LEFT OUTER JOIN Tllr_Mecanica_OT ON Tllr_Facturacion.Id_Empresa = Tllr_Mecanica_OT.Id_Empresa"
gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Sucursal = Tllr_Mecanica_OT.Id_Sucursal AND Tllr_Facturacion.Id_OT = Tllr_Mecanica_OT.Id_OT"
gstrSql = gstrSql & " AND Tllr_Facturacion.Seccion_OT = Tllr_Mecanica_OT.Seccion_OT AND Tllr_Facturacion.Id_Cargo = Tllr_Mecanica_OT.Id_Tipo_Cargo"
gstrSql = gstrSql & " WHERE Tllr_Mecanica_OT.Id_OT NOT IN (SELECT Id_Ot FROM Tllr_Actividades_Mecanico)"
gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Empresa = '" & gstrIdEmpresa & "'"
gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Sucursal = '" & gstrIdSucursal & "'"
gstrSql = gstrSql & " And Tllr_Facturacion.Id_Cargo = Tllr_Mecanica_Ot.Id_Tipo_Cargo"
gstrSql = gstrSql & " AND Tllr_Facturacion.Fecha_Facturacion BETWEEN '" & Me.lblFechaDesde & "' AND '" & Me.lblFechaHasta & "'"
If gstrEstadoProdMecanico = "F" Then
    gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B')"
ElseIf gstrEstadoProdMecanico = "A" Then
    gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B','V')"
Else
    gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('V')"
End If

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
With gadoPrincipal
    If Not .BOF And Not .EOF Then
        dblHorasRealesMecanica = IIf(IsNull(!TotalHorasMecanica), 0, !TotalHorasMecanica)
    End If
    .Close
End With
End If

'otros servicios
gstrSql = "SELECT SUM(Tllr_Otro_Ot.HorasReales) AS TOTALHORASOTRO FROM Tllr_Facturacion"
gstrSql = gstrSql & " LEFT OUTER JOIN Tllr_Otro_Ot ON Tllr_Facturacion.Id_Empresa = Tllr_Otro_Ot.Id_Empresa"
gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Sucursal = Tllr_Otro_Ot.Id_Sucursal AND Tllr_Facturacion.Id_OT = Tllr_Otro_Ot.Id_OT"
gstrSql = gstrSql & " AND Tllr_Facturacion.Seccion_OT = Tllr_Otro_Ot.Seccion_OT AND Tllr_Facturacion.Id_Cargo = Tllr_Otro_Ot.Id_Tipo_Cargo"
gstrSql = gstrSql & " WHERE Tllr_Facturacion.Id_Empresa = '" & gstrIdEmpresa & "'"
gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Sucursal = '" & gstrIdSucursal & "'"
gstrSql = gstrSql & " And Tllr_Facturacion.Id_Cargo = Tllr_Otro_Ot.Id_Tipo_Cargo"
gstrSql = gstrSql & " AND Tllr_Facturacion.Fecha_Facturacion BETWEEN '" & Me.lblFechaDesde & "' AND '" & Me.lblFechaHasta & "'"
If gstrEstadoProdMecanico = "F" Then
    gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B')"
ElseIf gstrEstadoProdMecanico = "A" Then
    gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B','V')"
Else
    gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('V')"
End If

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
With gadoPrincipal
    If Not .BOF And Not .EOF Then
        dblHorasRealesOtro = IIf(IsNull(!TotalHorasOtro), 0, !TotalHorasOtro)
    End If
    .Close
End With
End If

TraeHorasRealesTaller = dblHorasRealesMecanica + dblHorasRealesOtro

End Function
