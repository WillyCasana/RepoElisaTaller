VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistorialPlaca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Placa"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   Icon            =   "frmHistorialPlaca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6932
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha Emisión"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Seccion"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tipo"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Kilometros"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Trabajo Efectuado"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Patente"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   2
      Top             =   4320
      Width           =   480
   End
End
Attribute VB_Name = "frmHistorialPlaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Placa As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset

Sub BuscarHistoricoPlaca(strPlaca As String)
Dim mstrSql As String
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim ContLinea As Integer

FechaDesde = "01/01/2014"
FechaHasta = Format(Now(), "DD/MM/YYYY")

lvDetalle.ListItems.Clear
mstrWhere = "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & strPlaca & "%','','','','','" & FechaDesde & "','" & FechaHasta & " 23:59:00','T'"


mstrSql = "Exec Tllr_HistoricoPatente " & mstrWhere
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        With adoTemp
            If Not .BOF And Not .EOF Then
                ContLinea = 0
                While Not .EOF
                    Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
                    If !est = "F" Or !est = "B" Then
                       mstrNumeroDocumento = ValorNulo(!Nro_Factura_Emitida)
                    Else
                      mstrNumeroDocumento = "S/N"
                    End If
                    itmItem.SubItems(1) = ValorNulo(IIf(!est = "L", "LIQUIDADA", IIf(!est = "V", "VIGENTE", IIf(!est = "N", "NULA", IIf(!est = "C", "CERRADA", IIf(!est = "F", "FACTURADA", IIf(!est = "B", "BOLETEADA", "OTRO"))))))) & "(" & mstrNumeroDocumento & ")"
                    itmItem.SubItems(2) = Format(ValorNulo(!FEC), "dd/mm/yyyy")
                    itmItem.SubItems(3) = ValorNulo(!RECEP)
                    itmItem.SubItems(4) = ValorNulo(IIf(!Sec = "M", "MECANICA", "CARROCERIA"))
                    itmItem.SubItems(5) = ValorNulo(!GAR)
                    itmItem.SubItems(6) = ValorNulo(FormatoValor(!KMS, "", 0))
                    itmItem.SubItems(9) = ValorNulo(!Pat)
                    
                    adoTemp.MoveNext
                 Wend
             End If
                  
        End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

If frmRecepcion.txtPatente <> "" Then
    BuscarHistoricoPlaca (frmRecepcion.txtPatente)
End If
End Sub


Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
End Sub
