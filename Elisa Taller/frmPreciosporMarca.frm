VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreciosporMarca 
   Caption         =   "Ingrese Precios por Marca"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   Icon            =   "frmPreciosporMarca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "idmarca"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Marca"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Costo Mano Obra"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Venta Mano Obra"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Costo M. Obra Gtia."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Venta M. Obra Gtia."
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmPreciosporMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
GrabarPreciosMarcas
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
CargaMarcas
End Sub
Private Sub CargaMarcas()
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset

mstrSql = "SELECT Glbl_Marca.Id_Marca, Glbl_Marca.Descripcion AS Marca, Tllr_Marca_Precios_MO.CostoManoObra, "
mstrSql = mstrSql & "Tllr_Marca_Precios_MO.VentaManoObra , Tllr_Marca_Precios_MO.VentaMOGarantia, Tllr_Marca_Precios_MO.CostoMOGarantia "
mstrSql = mstrSql & "FROM Glbl_Marca LEFT OUTER JOIN "
mstrSql = mstrSql & "Tllr_Marca_Precios_MO ON Glbl_Marca.Id_Marca = Tllr_Marca_Precios_MO.Id_Marca "

Me.lvDetalle.ListItems.Clear

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst
            While Not .EOF
                Set Item = lvDetalle.ListItems.Add(, , !Id_Marca)
                Item.SubItems(1) = !Marca
                Item.SubItems(2) = IIf(IsNull(!CostoManoObra), FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!CostoManoObra), gstrMonedaLocal, gintDecimalesMoneda))
                Item.SubItems(3) = IIf(IsNull(!VentaManoObra), FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!VentaManoObra), gstrMonedaLocal, gintDecimalesMoneda))
                Item.SubItems(4) = IIf(IsNull(!CostoMOGarantia), FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!CostoMOGarantia), gstrMonedaLocal, gintDecimalesMoneda))
                Item.SubItems(5) = IIf(IsNull(!VentaMOGarantia), FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!VentaMOGarantia), gstrMonedaLocal, gintDecimalesMoneda))
                .MoveNext
            Wend
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal

End Sub
Private Sub GrabarPreciosMarcas()
Dim mstrSql As String

    Screen.MousePointer = vbHourglass
        
    'elimina todos los precios
    mstrSql = "delete from Tllr_Marca_Precios_MO"
    Conexion.SendHost mstrSql, , , , gcTiempoEspera
    
    'inserta todos los nuevos precios
    For i = 1 To Me.lvDetalle.ListItems.Count
        mstrSql = "Insert Into Tllr_Marca_Precios_MO (Id_Marca,CostoManoObra,VentaManoObra,CostoMOGarantia,VentaMOGarantia,Usr_Id,Usr_Fecha) "
        mstrSql = mstrSql & "Values ('" & Me.lvDetalle.ListItems(i) & "'," & SacarFormatoValor(Me.lvDetalle.ListItems(i).SubItems(2), gstrMonedaLocal) & ","
        mstrSql = mstrSql & SacarFormatoValor(Me.lvDetalle.ListItems(i).SubItems(3), gstrMonedaLocal) & "," & SacarFormatoValor(Me.lvDetalle.ListItems(i).SubItems(4), gstrMonedaLocal) & ","
        mstrSql = mstrSql & SacarFormatoValor(Me.lvDetalle.ListItems(i).SubItems(5), gstrMonedaLocal) & ",'" & gstrIdUsuario & "','" & Date & "')"
        
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        
    Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub lvDetalle_DblClick()
frmEditaPreciosporMarca.Show vbModal
End Sub

Private Sub lvDetalle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmEditaPreciosporMarca.Show vbModal
    End If
End Sub
