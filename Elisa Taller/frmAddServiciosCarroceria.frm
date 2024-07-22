VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAddServiciosCarroceria 
   Caption         =   "Seleccione Servicios de Carrocería"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frmAddServiciosCarroceria.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4590
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   6225
      Begin MSComctlLib.ListView lvwResultado 
         Height          =   3420
         Left            =   60
         TabIndex        =   6
         Top             =   945
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   6033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "IDPARTEPIEZAS"
            Text            =   "Código Parte - Pieza"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "PARTEPIEZA"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.TextBox txtD_P 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   540
         Width           =   570
      End
      Begin MSDataListLib.DataCombo dtcConceptoDyP 
         Bindings        =   "frmAddServiciosCarroceria.frx":0442
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   180
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datConceptoDyP 
         Height          =   330
         Left            =   4245
         Top             =   180
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
         Caption         =   "Conceptos D y P :"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sección :"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   555
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   6375
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":045F
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":0571
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":09C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":0E21
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":1279
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":138B
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":149D
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":15AF
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":16C1
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":17D3
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":18E5
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":19F7
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":1B09
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":1C1B
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":1D2D
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":1E3F
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":1F51
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":2063
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":2175
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":2287
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":26D9
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddServiciosCarroceria.frx":2B2B
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Left            =   4290
      TabIndex        =   0
      Top             =   4605
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "Agregar"
            Object.ToolTipText     =   "Quitar Servicio"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddServiciosCarroceria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSw As Boolean
Dim curValor As Currency

Sub TraspasarLinea()
Dim intX As Integer
With frmRecepcion
For intX = 1 To lvwResultado.ListItems.Count
    Set lvwResultado.SelectedItem = lvwResultado.ListItems(intX)
    If lvwResultado.SelectedItem.Checked = True Then
        Set glsiItem = .lvwServiciosCarroceria.ListItems.Add(, , dtcConceptoDyP.Text)
        glsiItem.SubItems(1) = dtcConceptoDyP.BoundText
        glsiItem.SubItems(2) = txtD_P
        glsiItem.SubItems(3) = lvwResultado.SelectedItem.SubItems(1)
        glsiItem.SubItems(4) = lvwResultado.SelectedItem.SubItems(2)
        glsiItem.SubItems(5) = .TRAEHORASDEFINIDAS("", dtcConceptoDyP.BoundText, lvwResultado.SelectedItem.SubItems(2))
        glsiItem.SubItems(6) = X
        glsiItem.SubItems(7) = X
        glsiItem.SubItems(8) = X
        glsiItem.SubItems(9) = X
        glsiItem.SubItems(10) = X
        glsiItem.SubItems(11) = X
        glsiItem.SubItems(12) = X
        glsiItem.SubItems(13) = X
        glsiItem.SubItems(14) = X
    End If
Next

End With
End Sub


Function TipoConcepto(strIdConcepto As String) As String
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset

mstrSql = "SELECT TOP 1 D_P AS TIPO FROM Tllr_Concepto WHERE ID_CONCEPTO='" & strIdConcepto & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            TipoConcepto = !tipo
        Else
            TipoConcepto = "N"
        End If
    End With
End If
End Function
Sub FillConceptosDyP(strIdCompañia As String)
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset

mstrSql = "SELECT Tllr_CiaSeguro_Concepto.Id_Concepto AS CODIGO,"
mstrSql = mstrSql & " Tllr_Concepto.Descripcion AS NOMBRE "
mstrSql = mstrSql & " FROM Tllr_CiaSeguro_Concepto LEFT OUTER JOIN Tllr_Concepto ON Tllr_CiaSeguro_Concepto.Id_Concepto = Tllr_Concepto.Id_Concepto "
mstrSql = mstrSql & " WHERE Tllr_CiaSeguro_Concepto.Id_Compañia_Seguro = '" & strIdCompañia & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datConceptoDyP
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcConceptoDyP.ListField = "NOMBRE"
            dtcConceptoDyP.BoundColumn = "CODIGO"
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal

End Sub

Sub FillPartePieza()
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset

mstrSql = "SELECT Id_Parte_Pieza AS CODIGO, Descripcion As NOMBRE"
mstrSql = mstrSql & " From Tllr_Parte_Pieza order by Descripcion"

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
             Set glsiItem = lvwResultado.ListItems.Add(, , !Codigo)
             glsiItem.SubItems(1) = !Nombre
             .MoveNext
        Wend
    End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal

End Sub

Sub ResultadoBusqueda(strIdCompañia As String, strIdConcepto As String)
Dim mstrSql As String, itmAux As ListItem
Dim adoPrincipal As New ADODB.Recordset

mstrSql = "SELECT Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Parte_Pieza AS IDPIEZA,"
mstrSql = mstrSql & " Tllr_Parte_Pieza.Descripcion AS PIEZA,"
mstrSql = mstrSql & " Tllr_CiaSeguro_Concepto_Parte_Pieza.Valor AS VALOR,"
mstrSql = mstrSql & " Tllr_CiaSeguro_Concepto_Parte_Pieza.Horas AS HORAS"
mstrSql = mstrSql & " FROM Tllr_CiaSeguro_Concepto RIGHT OUTER JOIN Tllr_CiaSeguro_Concepto_Parte_Pieza ON  Tllr_CiaSeguro_Concepto.Id_Compañia_Seguro = Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Compañia_Seguro AND Tllr_CiaSeguro_Concepto.Id_Concepto = Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Concepto LEFT OUTER JOIN"
mstrSql = mstrSql & " Tllr_Parte_Pieza ON Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Parte_Pieza = Tllr_Parte_Pieza.Id_Parte_Pieza"
mstrSql = mstrSql & " WHERE (Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Concepto = '" & strIdCompañia & "') AND  (Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Compañia_Seguro =  '" & strIdConcepto & "')"

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst
            While Not .EOF
                Set itmAux = lvwResultado.ListItems.Add(, , !IDPIEZA)
                itmAux.SubItems(1) = !PIEZA
                itmAux.SubItems(2) = Format(!Valor, "###,##0.0")
                itmAux.SubItems(3) = Format(!HORAS, "#0.0")
                .MoveNext
            Wend
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal

End Sub

Private Sub dtcConceptoDyP_Change()
    txtD_P = TipoConcepto(dtcConceptoDyP.BoundText)
End Sub

Private Sub Form_Activate()
If mblnSw Then
    FillConceptosDyP frmRecepcion.lblIdCompañia
    FillPartePieza
    mblnSw = False
End If
End Sub

Private Sub Form_Load()
mblnSw = True
End Sub

Private Sub tlbOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Agregar"
'    TraspasarLinea
Case "Cerrar"
    
End Select
End Sub
