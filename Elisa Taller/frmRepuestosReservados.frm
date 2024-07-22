VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmRepuestosReservados 
   Caption         =   "Repuestos a Reservar"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   Icon            =   "frmRepuestosReservados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlpicking 
      Left            =   2160
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Elija hacia Donde quiere Imprimir"
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   7920
      TabIndex        =   3
      Top             =   5820
      Width           =   1140
   End
   Begin VB.CommandButton cmdEnviar 
      Appearance      =   0  'Flat
      Caption         =   "&Enviar"
      Height          =   360
      Left            =   6555
      TabIndex        =   2
      Top             =   5820
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Repuestos Confirmados"
      ForeColor       =   &H00C00000&
      Height          =   5565
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9150
      Begin MSMAPI.MAPISession MAPISession1 
         Left            =   3960
         Top             =   2640
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin MSMAPI.MAPIMessages MAPIMessages1 
         Left            =   3240
         Top             =   2640
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSComctlLib.ListView lvwRepuestosReservados 
         Height          =   2265
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción Repuesto"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Solicitado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Reservado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio Unitario"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Bodega"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Ubicación"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cargo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "IDBODEGA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "IDUBICACION"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "NroReserva"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "E_Mail"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "Saldo"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView lvwRepuestosFaltantes 
         Height          =   2265
         Left            =   75
         TabIndex        =   4
         Top             =   3240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483647
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción Repuesto"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Solicitado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Reservado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Faltante"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Precio Unitario"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Bodega"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Ubicación"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Cargo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "IDBODEGA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "IDUBICACION"
            Object.Width           =   0
         EndProperty
      End
      Begin Crystal.CrystalReport rptPatente 
         Left            =   5160
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Repuestos Faltantes"
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   165
         TabIndex        =   5
         Top             =   2955
         Width           =   2025
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Reserva OK"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo... NO Reservado"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Saldo... Reservado"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   135
   End
End
Attribute VB_Name = "frmRepuestosReservados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset
Dim NroRegularizacion As String
Dim NroReserva As String
Dim nroReservaAux As String
Dim lsiItem As ListItem
Dim lsiItem2 As ListItem
Dim intIndice As Integer
Dim CodigoRepuesto As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEnviar_Click()

If cmdEnviar.Caption = "Imprimir" Then
    ImprimirReporte
Else

    Dim lstrBodegaAux As String
    Dim lstrSwBodega As String
    Dim swDesborde As Boolean
    Dim i As Integer
    Dim x As Integer
    
    If Me.lvwRepuestosReservados.ListItems.Count > 0 Then
    
        ReOrdenaLista Me.lvwRepuestosReservados, Me.lvwRepuestosReservados.ColumnHeaders(9)
        
        lstrBodegaAux = ""
        For i = 1 To Me.lvwRepuestosReservados.ListItems.Count
            lstrBodegaAux = Me.lvwRepuestosReservados.ListItems(i).SubItems(8)
            lstrSwBodega = "0"
            x = i
            swDesborde = False
            Do While lstrBodegaAux = Me.lvwRepuestosReservados.ListItems(x).SubItems(8)
                If lstrSwBodega = "0" Then
                    GrabarRegularizacion
                    GrabarRegularizacionDetalle NroRegularizacion, lstrBodegaAux
                    lstrSwBodega = "1"
                End If
                x = x + 1
                If x > Me.lvwRepuestosReservados.ListItems.Count Then
                    Exit For
                End If
                i = i + 1
            Loop
            i = i - 1
        Next
        
        'Actualizar tllr_ot con estado de reserva
        mstrSql = "UPDATE TLLR_OT SET Estado_Reserva='R' "
        mstrSql = mstrSql & "Where Id_OT='" & frmRecepcion.lblNroRecepcion & "' "
        mstrSql = mstrSql & "And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Seccion_OT='" & gstrSeccion & "'"
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        
        
        'grabar repuestos Reservados (Tllr_Repuestos_reservados)
        GrabarRepuestosReservados
        
        'grabar repuestos faltantes (Tllr_Repuestos_Faltantes)
        GrabarRepuestosFaltantes
        
        'emitir prepicking
        ImprimePrePicking
        
        'emitir mensaje a repuestos
        EnviarMailaBodega
        
        'desactiva boton de reserva
        frmRecepcion.cmdReserva.Enabled = False
        frmRecepcion.cmdAnularReserva.Enabled = True
    Else
        GrabarRepuestosFaltantes
    End If
    Unload Me
End If

End Sub

Private Sub Form_Activate()
    If mblnSW Then
        LlenaRepuestosReservados
        mblnSW = False
    End If
End Sub

Private Sub Form_Load()
mblnSW = True
If InStr(gstrProcedencia, "Presupuesto") > 0 Or gstrProcedencia = "Consulta" Then
    cmdEnviar.Caption = "Imprimir"
    'cmdEnviar.Visible = False
Else
    cmdEnviar.Caption = "Enviar"
    'cmdEnviar.Visible = True
End If
End Sub
Sub LlenaRepuestosReservados()
Dim i As Integer
Dim lvwListaRepuestos As ListView
Dim lintColumnaSaldo As Integer

If gstrProcedencia = "Presupuestos" Then
    Set lvwListaRepuestos = frmRecepcion.lvwRepuestos
    lintColumnaSaldo = 12
ElseIf gstrProcedencia = "Presupuesto Mantencion" Then
    Set lvwListaRepuestos = frmPresupuestoMantenciones.lvwRepuestos
    lintColumnaSaldo = 7
Else
    Set lvwListaRepuestos = frmRecepcion.lvwRepuestosMantencion
    lintColumnaSaldo = 7
End If

    Me.lvwRepuestosReservados.ListItems.Clear
    For i = 1 To lvwListaRepuestos.ListItems.Count
        '///// valida si el saldo es mayor igual al solicitado
        mstrSql = "Select top 1 stck_saldos.*,Glbl_Bodega.Id_Bodega ,Glbl_Bodega.Descripcion as Bodega, Glbl_Bodega.E_Mail, Stck_Ubicacion.Descripcion as Ubicacion, Stck_Ubicacion.id_ubicacion "
        mstrSql = mstrSql & "From stck_saldos "
        mstrSql = mstrSql & "inner join glbl_bodega on glbl_bodega.id_bodega = stck_saldos.id_bodega and glbl_bodega.Id_sucursal=stck_saldos.id_sucursal "
        mstrSql = mstrSql & "inner join stck_ubicacion on stck_ubicacion.id_ubicacion = stck_saldos.id_ubicacion "
        mstrSql = mstrSql & "where substring(id_item," & InStr(lvwListaRepuestos.ListItems(i), "°") + 1 & "," & (Len(lvwListaRepuestos.ListItems(i)) - InStr(lvwListaRepuestos.ListItems(i), "°")) & ")='"
        mstrSql = mstrSql & Mid(lvwListaRepuestos.ListItems(i), InStr(lvwListaRepuestos.ListItems(i), "°") + 1, (Len(lvwListaRepuestos.ListItems(i)) - InStr(lvwListaRepuestos.ListItems(i), "°")))
        mstrSql = mstrSql & "' And Stck_saldos.Id_empresa='" & gstrIdEmpresa & "' And stck_saldos.Id_Sucursal='" & gstrIdSucursal & "' "
        mstrSql = mstrSql & "And stck_saldos.saldo >=" & CDbl(lvwListaRepuestos.ListItems(i).SubItems(2))
        
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            With adoPrincipal
                If Not .BOF And Not .EOF Then
                    Set lsiItem = lvwRepuestosReservados.ListItems.Add(, , !Id_Item)
                    lsiItem.SubItems(1) = lvwListaRepuestos.ListItems(i).SubItems(1)
                    lsiItem.SubItems(2) = lvwListaRepuestos.ListItems(i).SubItems(2)
                    lsiItem.SubItems(3) = lvwListaRepuestos.ListItems(i).SubItems(2)
                    lsiItem.SubItems(4) = lvwListaRepuestos.ListItems(i).SubItems(3)
                    lsiItem.SubItems(5) = ValorNulo(!bodega)
                    lsiItem.SubItems(6) = ValorNulo(!ubicacion)
                    lsiItem.SubItems(7) = lvwListaRepuestos.ListItems(i).SubItems(5)
                    lsiItem.SubItems(8) = ValorNulo(!id_bodega)
                    lsiItem.SubItems(9) = ValorNulo(!id_Ubicacion)
                    lsiItem.SubItems(11) = "R"
                    lsiItem.SubItems(12) = IIf(lvwListaRepuestos.ListItems(i).SubItems(6) = "", "T", lvwListaRepuestos.ListItems(i).SubItems(6))
                    lsiItem.SubItems(13) = ValorNulo(!E_Mail)
                    lvwListaRepuestos.ListItems(i).SubItems(lintColumnaSaldo) = "c/s"   'existe saldo
                    lsiItem.SubItems(14) = ValorNulo(!Saldo)
                Else
                    mstrSql = "Select top 1 stck_saldos.*,Glbl_Bodega.Id_Bodega ,Glbl_Bodega.Descripcion as Bodega, Glbl_Bodega.E_Mail, Stck_Ubicacion.Descripcion as Ubicacion, Stck_Ubicacion.id_ubicacion "
                    mstrSql = mstrSql & "From stck_saldos "
                    mstrSql = mstrSql & "inner join glbl_bodega on glbl_bodega.id_bodega = stck_saldos.id_bodega and glbl_bodega.Id_sucursal=stck_saldos.id_sucursal "
                    mstrSql = mstrSql & "inner join stck_ubicacion on stck_ubicacion.id_ubicacion = stck_saldos.id_ubicacion "
                    mstrSql = mstrSql & "where substring(id_item," & InStr(lvwListaRepuestos.ListItems(i), "°") + 1 & "," & (Len(lvwListaRepuestos.ListItems(i)) - InStr(lvwListaRepuestos.ListItems(i), "°")) & ")='"
                    mstrSql = mstrSql & Mid(lvwListaRepuestos.ListItems(i), InStr(lvwListaRepuestos.ListItems(i), "°") + 1, (Len(lvwListaRepuestos.ListItems(i)) - InStr(lvwListaRepuestos.ListItems(i), "°")))
                    mstrSql = mstrSql & "' And Stck_saldos.Id_empresa='" & gstrIdEmpresa & "' And stck_saldos.Id_Sucursal='" & gstrIdSucursal & "' "
                    mstrSql = mstrSql & "And stck_saldos.saldo > 0 "
                    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                        With adoPrincipal
                            If Not .BOF And Not .EOF Then
                                '///// Si encuentra saldo menor al solicitado
                                If MsgBox("No se Encontro saldo suficiente para el repuesto " & lvwListaRepuestos.ListItems(i).SubItems(1) & " " & Chr(13) & "Saldo = " & !Saldo & Chr(13) & "Desea Reservar el Saldo", vbQuestion + vbYesNo, "Confirma Reserva") = vbYes Then
                                    Set lsiItem = lvwRepuestosReservados.ListItems.Add(, , !Id_Item)
                                    lsiItem.SubItems(1) = lvwListaRepuestos.ListItems(i).SubItems(1)
                                    lsiItem.SubItems(2) = lvwListaRepuestos.ListItems(i).SubItems(2)
                                    lsiItem.SubItems(3) = FormatoValor(!Saldo, "", 1)
                                    lsiItem.SubItems(4) = lvwListaRepuestos.ListItems(i).SubItems(3)
                                    lsiItem.SubItems(5) = ValorNulo(!bodega)
                                    lsiItem.SubItems(6) = ValorNulo(!ubicacion)
                                    lsiItem.SubItems(7) = lvwListaRepuestos.ListItems(i).SubItems(5)
                                    lsiItem.SubItems(8) = ValorNulo(!id_bodega)
                                    lsiItem.SubItems(9) = ValorNulo(!id_Ubicacion)
                                    lsiItem.SubItems(11) = "S"
                                    lsiItem.SubItems(12) = lvwListaRepuestos.ListItems(i).SubItems(6)
                                    lsiItem.SubItems(13) = ValorNulo(!E_Mail)
                                    lvwListaRepuestos.ListItems(i).SubItems(lintColumnaSaldo) = "s/p"  'existe saldo
                                    lsiItem.SubItems(14) = ValorNulo(!Saldo)
                                    
                                    'cambia color a saldo reservado
                                    lvwRepuestosReservados.ListItems(lvwRepuestosReservados.ListItems.Count).ForeColor = &HFF0000
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(1).ForeColor = &HFF0000
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(2).ForeColor = &HFF0000
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(3).ForeColor = &HFF0000
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(4).ForeColor = &HFF0000
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(5).ForeColor = &HFF0000
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(6).ForeColor = &HFF0000
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(7).ForeColor = &HFF0000
                                    
                                    
                                    Set lsiItem = lvwRepuestosFaltantes.ListItems.Add(, , !Id_Item)
                                    lsiItem.SubItems(1) = lvwListaRepuestos.ListItems(i).SubItems(1)
                                    lsiItem.SubItems(2) = lvwListaRepuestos.ListItems(i).SubItems(2)
                                    lsiItem.SubItems(3) = ValorNulo(!Saldo)
                                    lsiItem.SubItems(4) = CDbl(lsiItem.SubItems(2)) - CDbl(lsiItem.SubItems(3))
                                    lsiItem.SubItems(5) = lvwListaRepuestos.ListItems(i).SubItems(3)
                                    lsiItem.SubItems(6) = ValorNulo(!bodega)
                                    lsiItem.SubItems(7) = ValorNulo(!ubicacion)
                                    lsiItem.SubItems(8) = lvwListaRepuestos.ListItems(i).SubItems(5)
                                    lsiItem.SubItems(9) = ValorNulo(!id_bodega)
                                    lsiItem.SubItems(10) = ValorNulo(!id_Ubicacion)
                                Else
                                    Set lsiItem = lvwRepuestosReservados.ListItems.Add(, , !Id_Item)
                                    lsiItem.SubItems(1) = lvwListaRepuestos.ListItems(i).SubItems(1)
                                    lsiItem.SubItems(2) = lvwListaRepuestos.ListItems(i).SubItems(2)
                                    lsiItem.SubItems(3) = FormatoValor(ValorNulo(!Saldo), "", 1)
                                    lsiItem.SubItems(4) = lvwListaRepuestos.ListItems(i).SubItems(3)
                                    lsiItem.SubItems(5) = ValorNulo(!bodega)
                                    lsiItem.SubItems(6) = ValorNulo(!ubicacion)
                                    lsiItem.SubItems(7) = lvwListaRepuestos.ListItems(i).SubItems(5)
                                    lsiItem.SubItems(8) = ValorNulo(!id_bodega)
                                    lsiItem.SubItems(9) = ValorNulo(!id_Ubicacion)
                                    lsiItem.SubItems(11) = "P"
                                    lsiItem.SubItems(12) = lvwListaRepuestos.ListItems(i).SubItems(6)
                                    lsiItem.SubItems(13) = ValorNulo(!E_Mail)
                                    lsiItem.SubItems(14) = ValorNulo(!Saldo)
                                    
                                    'cambia color a saldo no reservado
                                    lvwRepuestosReservados.ListItems(lvwRepuestosReservados.ListItems.Count).ForeColor = &HC0&
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(6).ForeColor = &HC0&
                                    lvwRepuestosReservados.ListItems(Me.lvwRepuestosReservados.ListItems.Count).ListSubItems(7).ForeColor = &HC0&
                                    
                                    
                                End If
                            Else  '//// Simplemente no encuentra stock del repuesto
                                Set lsiItem = lvwRepuestosFaltantes.ListItems.Add(, , lvwListaRepuestos.ListItems(i))
                                lsiItem.SubItems(1) = lvwListaRepuestos.ListItems(i).SubItems(1)
                                lsiItem.SubItems(2) = lvwListaRepuestos.ListItems(i).SubItems(2)
                                lsiItem.SubItems(3) = "0"
                                lsiItem.SubItems(4) = CDbl(lsiItem.SubItems(2)) - CDbl(lsiItem.SubItems(3))
                                lsiItem.SubItems(5) = lvwListaRepuestos.ListItems(i).SubItems(3)
                                lsiItem.SubItems(6) = ""
                                lsiItem.SubItems(7) = ""
                                lsiItem.SubItems(8) = ""
                                lvwListaRepuestos.ListItems(i).SubItems(lintColumnaSaldo) = "n/s"  'existe saldo
                                
                            End If
                        End With
                    End If
                End If
            End With
        End If
    Next
End Sub

Sub GrabarRegularizacion()
        
        mstrSql = "SELECT MAX(Cast(id_regularizacion as float)) as Resultado FROM Stck_Regularizacion where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                NroRegularizacion = IIf(IsNull(adoPrincipal!resultado), 1, adoPrincipal!resultado)
            End If
            
        End If '//////////////
        
        NroReserva = Retorna_Valor_General("Select MAX(cast(isnull(id_reserva,0) as float)) as CorrelativoReserva from Stck_Regularizacion where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
        
        mstrSql = "INSERT INTO Stck_Regularizacion "
        mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal, "
        mstrSql = mstrSql & "Id_Regularizacion , Id_Concepto, "
        mstrSql = mstrSql & "Fecha, Comentario, "
        mstrSql = mstrSql & "Subtotal, Usr_Id, "
        mstrSql = mstrSql & "Usr_Fecha, Id_Reserva, Estado_Reserva, Id_OT) "
        mstrSql = mstrSql & "VALUES ("
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "',"
        mstrSql = mstrSql & "'" & CStr(Val(NroRegularizacion) + 1) & "','" & 21 & "',"
        mstrSql = mstrSql & "'" & Date & "','Reservado Por " & gstrIdUsuario & "',"
        mstrSql = mstrSql & "0," & "'" & gstrIdUsuario & "','" & Date & "',"
        mstrSql = mstrSql & "'" & CStr(Val(NroReserva) + 1) & "','V','" & gstrSeccion & frmRecepcion.lblNroRecepcion & "')"
        
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
            'mblnTablaVacia = False
            'ActivaBotones
            'Me.Tag = ""
        End If '//////////////

End Sub
Sub GrabarRegularizacionDetalle(pstrRegularizacion As String, pstrBodega As String)
Dim i As Integer
    
    For i = 1 To Me.lvwRepuestosReservados.ListItems.Count
        If pstrBodega = Me.lvwRepuestosReservados.ListItems(i).SubItems(8) Then
            mstrSql = "INSERT INTO Stck_Regularizacion_Detalle "
            mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal, "
            mstrSql = mstrSql & "Id_Regularizacion , Id_Linea, "
            mstrSql = mstrSql & "Id_bodega, id_Ubicacion, id_item, "
            mstrSql = mstrSql & "Canrtidad, Precio_Unitario, Subtotal, "
            mstrSql = mstrSql & "Usr_id, Usr_Fecha, Tipo_Cargo) "
            mstrSql = mstrSql & "VALUES ("
            mstrSql = mstrSql & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "',"
            mstrSql = mstrSql & "'" & CStr(Val(NroRegularizacion) + 1) & "','" & i & "',"
            mstrSql = mstrSql & "'" & Me.lvwRepuestosReservados.ListItems(i).SubItems(8) & "',"
            mstrSql = mstrSql & "'" & Me.lvwRepuestosReservados.ListItems(i).SubItems(9) & "',"
            mstrSql = mstrSql & "'" & Me.lvwRepuestosReservados.ListItems(i) & "',"
            mstrSql = mstrSql & "'" & Me.lvwRepuestosReservados.ListItems(i).SubItems(3) & "',"
            mstrSql = mstrSql & "'" & CDbl(Me.lvwRepuestosReservados.ListItems(i).SubItems(4)) & "',"
            mstrSql = mstrSql & "0,"
            mstrSql = mstrSql & "'" & gstrIdUsuario & "',"
            mstrSql = mstrSql & "'" & Date & "',"
            mstrSql = mstrSql & "'" & Me.lvwRepuestosReservados.ListItems(i).SubItems(7) & "')"
            
            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                'MsgBox "si"
            End If '//////////////
            
            '/// Actualiza saldos en linea
            Actualiza_Saldos lvwRepuestosReservados.ListItems(i).SubItems(2), "S", gstrIdEmpresa, gstrIdSucursal, lvwRepuestosReservados.ListItems(i).SubItems(8), lvwRepuestosReservados.ListItems(i).SubItems(9), lvwRepuestosReservados.ListItems(i)
            
            '/// Guarda el numero de reserva para imprimir después
            lvwRepuestosReservados.ListItems(i).SubItems(10) = CStr(Val(NroReserva) + 1)
        End If
        
    Next
End Sub
Sub GrabarRepuestosFaltantes()
Dim i As Integer
    
    For i = 1 To Me.lvwRepuestosFaltantes.ListItems.Count
        mstrSql = "INSERT INTO Tllr_Repuestos_Faltantes "
        mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal, "
        mstrSql = mstrSql & "Id_OT, Id_item, Fecha, Solicitado, Despachado, "
        mstrSql = mstrSql & "Precio_Unitario, Patente, Seccion_OT) "
        mstrSql = mstrSql & "VALUES ("
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "',"
        mstrSql = mstrSql & "'" & frmRecepcion.lblNroRecepcion & "',"
        mstrSql = mstrSql & "'" & Me.lvwRepuestosFaltantes.ListItems(i) & "',"
        mstrSql = mstrSql & "'" & frmRecepcion.pckFechaAtencion & "',"
        mstrSql = mstrSql & CDbl(Me.lvwRepuestosFaltantes.ListItems(i).SubItems(2)) & ","
        mstrSql = mstrSql & CDbl(Me.lvwRepuestosFaltantes.ListItems(i).SubItems(3)) & ","
        mstrSql = mstrSql & CDbl(Me.lvwRepuestosFaltantes.ListItems(i).SubItems(5)) & ","
        mstrSql = mstrSql & "'" & frmRecepcion.txtPatente & "','" & gstrSeccion & "')"
        
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
            MsgBox "Problemas para guardar Repuestos Faltantes", vbExclamation, "Guardar Repuestos Faltantes"
        End If '//////////////
    Next
    
End Sub

Sub GrabarRepuestosReservados()
    With lvwRepuestosReservados
        If .ListItems.Count > 0 Then
            'elimina los repuestos que fueron cargados desde un presupuesto,para que no se repitan
            mstrSql = "Delete from Tllr_Repuestos_Reservados where Id_Empresa='" & gstrIdEmpresa & "'"
            mstrSql = mstrSql & " And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrSql = mstrSql & " And Seccion_Ot='" & gstrSeccion & "'"
            mstrSql = mstrSql & " And Id_Ot='" & frmRecepcion.lblNroRecepcion & "'"
            Conexion.SendHost mstrSql, , , , gcTiempoEspera
            
            For intIndice = 1 To .ListItems.Count
                Set .SelectedItem = .ListItems(intIndice)
                mstrSql = "Insert Into Tllr_Repuestos_Reservados"
                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,"
                mstrSql = mstrSql & " Id_OT , Seccion_OT, "
                mstrSql = mstrSql & " Id_Item,Precio_Unitario,"
                mstrSql = mstrSql & " Solicitado,Reservado, "
                mstrSql = mstrSql & " Estado,Tipo, Nro_Reserva)"
                mstrSql = mstrSql & " Values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
                mstrSql = mstrSql & " '" & frmRecepcion.lblNroRecepcion & "', '" & gstrSeccion & "',"
                mstrSql = mstrSql & " '" & Trim(.SelectedItem) & "',"
                mstrSql = mstrSql & CDbl(.SelectedItem.SubItems(4)) & ", "
                mstrSql = mstrSql & CDbl(.SelectedItem.SubItems(2)) & ", "
                mstrSql = mstrSql & CDbl(.SelectedItem.SubItems(3)) & ", "
                mstrSql = mstrSql & "'" & .SelectedItem.SubItems(11) & "','" & .SelectedItem.SubItems(12) & "','" & .SelectedItem.SubItems(10) & "')"
                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
                    MsgBox "Problemas Para Guardar Repuestos Reservados", vbExclamation, "Guarda Repuestos Reservados"
                    Exit Sub
                End If
            Next
        End If
    End With
End Sub

Sub ImprimePrePicking()
Dim lintCuentaItems As Integer
Dim lstrCodigoItem As String

Dim i As Integer
Dim x As Integer
Dim fila As Integer

On Error GoTo Cancela_impresion

    For i = 1 To Me.lvwRepuestosReservados.ListItems.Count
        MsgBox "Ud. Imprimira la reserva Nro." & Me.lvwRepuestosReservados.ListItems(i).SubItems(10) & " en la BODEGA " & Me.lvwRepuestosReservados.ListItems(i).SubItems(5), vbInformation, "Imprimir Reserva"
        cdlpicking.CancelError = True
        cdlpicking.DialogTitle = "Elija Donde quiere Imprimir la Reserva N° " & Me.lvwRepuestosReservados.ListItems(i).SubItems(10)
        cdlpicking.ShowPrinter
        
        Encabezado i
        fila = 4700
        x = i
        nroReservaAux = Me.lvwRepuestosReservados.ListItems(i).SubItems(10)
        While nroReservaAux = Me.lvwRepuestosReservados.ListItems(x).SubItems(10)
            CodigoRepuesto = Retorna_Valor_General("Select Prefijo + '-' + Basico + '-' + Sufijo as Codigo from Stck_Item where id_item='" & Me.lvwRepuestosReservados.ListItems(i) & "'", gcdynamic)
            ImprimeObjeto fila, 500, CodigoRepuesto, 8
            ImprimeObjeto fila, 2000, Me.lvwRepuestosReservados.ListItems(i).SubItems(1), 8
            ImprimeObjeto fila, 5500, Me.lvwRepuestosReservados.ListItems(i).SubItems(3), 8
            ImprimeObjeto fila, 6500, Me.lvwRepuestosReservados.ListItems(i).SubItems(5), 8
            ImprimeObjeto fila, 10000, Me.lvwRepuestosReservados.ListItems(i).SubItems(6), 8
            x = x + 1
            If x > Me.lvwRepuestosReservados.ListItems.Count Then
                ImprimeObjeto fila + 300, 100, "____________________________________________________________________", 16    '//////////  Encabezado
                Printer.EndDoc
                Exit For
            End If
            i = i + 1
            fila = fila + 200
        Wend
        i = i - 1
        ImprimeObjeto fila + 100, 100, "____________________________________________________________________", 16    '//////////  Encabezado
        Printer.EndDoc
    Next
    
Cancela_impresion:

If Err.Number = 32755 Then
    MsgBox "Impresión Cancelada", vbInformation, "Imprimiendo"
End If

End Sub
Sub Encabezado(x As Integer)
    ImprimeObjeto 100, 100, "ElisaTaller", 8, True, True           '//////////  Encabezado
    ImprimeObjeto 300, 9000, "Lima, " & Date, 8                  '//////////  Encabezado
    ImprimeObjeto 500, 9000, Time, 8                                 '//////////  Encabezado
    ImprimeObjeto 500, 100, gstrEmpresa, 8                           '//////////  Encabezado
    ImprimeObjeto 700, 100, gstrSucursal, 8                          '//////////  Encabezado
    ImprimeObjeto 900, 100, gstrUsuario, 8                           '//////////  Encabezado
    ImprimeObjeto 1400, 3000, "SOLICITUD DE REPUESTO A ALMACEN - N° DE RESERVA " & Me.lvwRepuestosReservados.ListItems(x).SubItems(10), 10, True    '//////////  Encabezado
    ImprimeObjeto 2000, 100, "SOLICITA  : " & gstrUsuario, 12       '//////////  Encabezado
    ImprimeObjeto 2500, 100, "ORIGEN    : TALLER ", 12              '//////////  Encabezado
    ImprimeObjeto 3000, 100, "N° OT       : " & frmRecepcion.lblNroRecepcion, 12              '//////////  Encabezado
    ImprimeObjeto 3500, 100, "Solicitud a Bodega de los siguientes repuestos:", 9, True, True '//////////  Encabezado"
    ImprimeObjeto 3800, 100, "____________________________________________________________________", 16    '//////////  Encabezado
    ImprimeObjeto 4200, 500, "Item", 8
    ImprimeObjeto 4200, 2000, "Descripción", 8
    ImprimeObjeto 4200, 5500, "Cantidad", 8
    ImprimeObjeto 4200, 6500, "Almacén", 8
    ImprimeObjeto 4200, 10000, "Ubicación", 8
End Sub
Sub EnviarMailaBodega()
Dim mstrMensaje As String
Dim i As Integer
Dim x As Integer
Dim j As Integer
        
If gblnEnviaMailBodega = True Then  '/// parametro si envia mail
        
    MAPISession1.SignOn
    
    For i = 1 To Me.lvwRepuestosReservados.ListItems.Count 'repuestos reservados
        x = i
        nroReservaAux = Me.lvwRepuestosReservados.ListItems(i).SubItems(10)
        mstrMensaje = "Sirvase Preparar los siguiente Repuestos :" & Chr(13) & Chr(13)
        mstrMensaje = mstrMensaje & "N° de Reserva: " & nroReservaAux & Chr(13)
        mstrMensaje = mstrMensaje & "N° de OT     : " & frmRecepcion.lblNroRecepcion & Chr(13) & Chr(13)
        
        While nroReservaAux = Me.lvwRepuestosReservados.ListItems(x).SubItems(10)
            'CodigoRepuesto = Retorna_Valor_General("Select Prefijo + '-' + Basico + '-' + Sufijo as Codigo from Stck_Item where id_item='" & Me.lvwRepuestosReservados.ListItems(i) & "'", gcdynamic)
            mstrMensaje = mstrMensaje & lvwRepuestosReservados.ListItems(i) & "   " & CStr(CDbl(lvwRepuestosReservados.ListItems(i).SubItems(2))) & " "
            mstrMensaje = mstrMensaje & lvwRepuestosReservados.ListItems(i).SubItems(1) & Chr(13)
            x = x + 1
            If x > Me.lvwRepuestosReservados.ListItems.Count Then
                With MAPIMessages1
                '/// Id de conexion
                .SessionID = MAPISession1.SessionID
                '/// crea un nuevo mensaje
                .Compose
                .MsgReceiptRequested = True
                .RecipAddress = lvwRepuestosReservados.ListItems(i).SubItems(13)
                .AddressResolveUI = True
                .ResolveName
                '/// Busca remitente
                .MsgSubject = "Reserva de Repuestos N° " & nroReservaAux
                .MsgNoteText = mstrMensaje
                .Send False
                End With
                Exit For
            End If
            i = i + 1
            
        Wend
        i = i - 1
        With MAPIMessages1
        '/// Id de conexion
        .SessionID = MAPISession1.SessionID
        '/// crea un nuevo mensaje
        .Compose
        .MsgReceiptRequested = True
        .RecipAddress = lvwRepuestosReservados.ListItems(i).SubItems(13)
        .AddressResolveUI = True
        .ResolveName
        '/// Busca remitente
        .MsgSubject = "Reserva de Repuestos N° " & nroReservaAux
        .MsgNoteText = mstrMensaje
        .Send False
        End With
    Next i
    
    If gstrMailRepuestosFallidos <> "" Then
        If Me.lvwRepuestosFaltantes.ListItems.Count > 0 Then
            mstrMensaje = "Los Siguientes repuestos no se Encuentran en Stock (Venta Fallida) :" & Chr(13) & Chr(13)
            mstrMensaje = mstrMensaje & "N° de OT     : " & frmRecepcion.lblNroRecepcion & Chr(13) & Chr(13)
            For i = 1 To Me.lvwRepuestosFaltantes.ListItems.Count  'repuestos faltantes
                CodigoRepuesto = Retorna_Valor_General("Select Prefijo + '-' + Basico + '-' + Sufijo as Codigo from Stck_Item where id_item='" & Me.lvwRepuestosReservados.ListItems(i) & "'", gcdynamic)
                mstrMensaje = mstrMensaje & CodigoRepuesto & "  " & CStr(CDbl(lvwRepuestosFaltantes.ListItems(i).SubItems(2))) & " "
                mstrMensaje = mstrMensaje & lvwRepuestosFaltantes.ListItems(i).SubItems(1) & Chr(13)
            Next
            With MAPIMessages1
            '/// Id de conexion
            .SessionID = MAPISession1.SessionID
            '/// crea un nuevo mensaje
            .Compose
            .MsgReceiptRequested = True
            .RecipAddress = gstrMailRepuestosFallidos
            .AddressResolveUI = True
            .ResolveName
            '/// Busca remitente
            .MsgSubject = "Venta Fallida "
            .MsgNoteText = mstrMensaje
            .Send False
            End With
        End If
    End If
    
    
    
    MAPISession1.SignOff
End If
End Sub

Sub ImprimirReporte()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer

    
    If Me.lvwRepuestosReservados.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Item text,Descripcion text,Solicitado Text,Saldo text,Precio Double,Ubicacion text,Bodega text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To Me.lvwRepuestosReservados.ListItems.Count
        Set lvwRepuestosReservados.SelectedItem = lvwRepuestosReservados.ListItems(i)
        Tabla.AddNew
        Tabla!item = IIf(lvwRepuestosReservados.SelectedItem = "", " ", lvwRepuestosReservados.SelectedItem)
        Tabla!Descripcion = IIf(lvwRepuestosReservados.SelectedItem.SubItems(1) = "", " ", lvwRepuestosReservados.SelectedItem.SubItems(1))
        Tabla!Solicitado = IIf(lvwRepuestosReservados.SelectedItem.SubItems(3) = "", "", lvwRepuestosReservados.SelectedItem.SubItems(3))
        Tabla!Saldo = IIf(lvwRepuestosReservados.SelectedItem.SubItems(14) = "", " ", lvwRepuestosReservados.SelectedItem.SubItems(14))
        Tabla!Precio = IIf(lvwRepuestosReservados.SelectedItem.SubItems(4) = "", " ", lvwRepuestosReservados.SelectedItem.SubItems(4))
        Tabla!ubicacion = IIf(lvwRepuestosReservados.SelectedItem.SubItems(6) = "", " ", lvwRepuestosReservados.SelectedItem.SubItems(6))
        Tabla!bodega = IIf(lvwRepuestosReservados.SelectedItem.SubItems(5) = "", " ", lvwRepuestosReservados.SelectedItem.SubItems(5))
        Tabla.Update
    Next i
    For i = 1 To Me.lvwRepuestosFaltantes.ListItems.Count
        Set lvwRepuestosFaltantes.SelectedItem = lvwRepuestosFaltantes.ListItems(i)
        Tabla.AddNew
        Tabla!item = IIf(lvwRepuestosFaltantes.SelectedItem = "", " ", lvwRepuestosFaltantes.SelectedItem)
        Tabla!Descripcion = IIf(lvwRepuestosFaltantes.SelectedItem.SubItems(1) = "", " ", lvwRepuestosFaltantes.SelectedItem.SubItems(1))
        Tabla!Solicitado = IIf(lvwRepuestosFaltantes.SelectedItem.SubItems(2) = "", "", lvwRepuestosFaltantes.SelectedItem.SubItems(2))
        Tabla!Saldo = IIf(lvwRepuestosFaltantes.SelectedItem.SubItems(3) = "", " ", lvwRepuestosFaltantes.SelectedItem.SubItems(3))
        Tabla!Precio = IIf(lvwRepuestosFaltantes.SelectedItem.SubItems(5) = "", " ", lvwRepuestosFaltantes.SelectedItem.SubItems(5))
        Tabla!ubicacion = IIf(lvwRepuestosFaltantes.SelectedItem.SubItems(6) = "", " ", lvwRepuestosFaltantes.SelectedItem.SubItems(6))
        Tabla!bodega = IIf(lvwRepuestosFaltantes.SelectedItem.SubItems(7) = "", " ", lvwRepuestosFaltantes.SelectedItem.SubItems(7))
        Tabla.Update
    Next i
   Tabla.Close
   Dbsnueva.Close
   
   With rptPatente
        .ReportFileName = gstrPathReporte & "\ConsultaSaldos.Rpt"
        .WindowTitle = "Consulta de Saldos"
        .WindowState = crptMaximized
        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Consulta de Saldos'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "TDecimal=" & gintDecimalesMoneda
        
        .Destination = crptToWindow
        .Action = True
   End With
   
   
   Screen.MousePointer = 1

End Sub
