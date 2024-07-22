VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmValoresPorCargo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valores por Cargo"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmValoresPorCargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6960
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvwValores 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo Cargo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Valor (neto)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor (bruto)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Estado Cargo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Fecha Liquidación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fecha Facturación"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmValoresPorCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Item As ListItem

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = Default
End Sub

Private Sub Form_Load()
Dim lstrSQL As String
Dim adoTemp As New ADODB.Recordset
Dim ldblNeto As Double
Dim ldblTotal As Double
Dim lstrEstado As String

lstrSQL = ""
lstrSQL = lstrSQL & "select tllr_facturacion.Id_Cargo, "
lstrSQL = lstrSQL & "Tllr_Tipo_Cargo.Descripcion, "
lstrSQL = lstrSQL & "tllr_facturacion.Total_Neto, "
lstrSQL = lstrSQL & "tllr_facturacion.Total, "
lstrSQL = lstrSQL & "tllr_facturacion.Estado, "
lstrSQL = lstrSQL & "tllr_facturacion.Nro_Factura_Emitida, "
lstrSQL = lstrSQL & "tllr_facturacion.Fecha_Liquidacion, "
lstrSQL = lstrSQL & "tllr_facturacion.Fecha_Facturacion "
lstrSQL = lstrSQL & "from tllr_facturacion "
lstrSQL = lstrSQL & "INNER JOIN Tllr_Tipo_Cargo "
lstrSQL = lstrSQL & "ON Tllr_Facturacion.Id_Cargo = Tllr_Tipo_Cargo.Id_Tipo_Cargo and Tllr_Facturacion.Id_Empresa = Tllr_Tipo_Cargo.Id_Empresa "
lstrSQL = lstrSQL & "where tllr_facturacion.id_ot='" & frmRecepcion.lblNroRecepcion.Text & "'"
lstrSQL = lstrSQL & " and Tllr_Facturacion.Id_Sucursal='" & gstrIdSucursal & "'"
If Conexion.SendHost(lstrSQL, adoTemp, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoTemp
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set Item = Me.lvwValores.ListItems.Add(, , ValorNulo(!Id_Cargo))
                Item.SubItems(1) = ValorNulo(!Descripcion)
                If IsNumeric(ValorNulo(!Total_Neto)) Then ldblNeto = !Total_Neto Else ldblNeto = 0
                Item.SubItems(2) = FormatoValor(ldblNeto, "", gintDecimalesMoneda)
                Item.SubItems(3) = FormatoValor(ldblNeto * 1.18, "", gintDecimalesMoneda)
                Select Case ValorNulo(!estado)
                    Case "B"
                        lstrEstado = "Boleteado" & " (" & ValorNulo(!Nro_Factura_Emitida) & ")"
                    Case "F"
                        lstrEstado = "Facturado" & " (" & ValorNulo(!Nro_Factura_Emitida) & ")"
                    Case "V"
                        lstrEstado = "Vigente"
                    Case Else
                        lstrEstado = "?"
                End Select
                Item.SubItems(4) = lstrEstado
                Item.SubItems(5) = ValorNulo(!Fecha_Liquidacion)
                Item.SubItems(6) = ValorNulo(!Fecha_Facturacion)
                .MoveNext
            Wend
        End If
    End With
End If
Conexion.CloseHost adoTemp

End Sub
