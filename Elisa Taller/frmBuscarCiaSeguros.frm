VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBuscarCiaSeguros 
   Caption         =   "Buscar Compañia de Seguros"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   Icon            =   "frmBuscarCiaSeguros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton optdescripcion 
         Appearance      =   0  'Flat
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optcodigo 
         Appearance      =   0  'Flat
         Caption         =   "Codigo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox cmbcoincidir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3720
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvwResultado 
      Height          =   2220
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   3916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripcion"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label lblcoincidir 
      Caption         =   "Coincidir"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3765
      Width           =   975
   End
End
Attribute VB_Name = "frmBuscarCiaSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoPrincipal As New ADODB.Recordset
Dim AdoPaso As New ADODB.Recordset
Dim strSql As String
Public Paso As String
Public r As String

Private Sub cmdBuscar_Click()
Dim strSql As String
Screen.MousePointer = vbHourglass
    
    strSql = "exec Tllr_Buscar_CiaSeguros '" & gstrIdEmpresa & "','" & Me.txtCodigo & "','" & Me.txtDescripcion & "','" & Me.cmbcoincidir.ListIndex + 1 & "'"
    Llena_List_View strSql
Screen.MousePointer = vbDefault
End Sub


Sub Llena_List_View(strSql As String)
Dim strres As String
Dim Item As ListItem
Dim iFila As Double


        If Not Conexion.SendHost(strSql, AdoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    strres = MsgBox("Error en Conexion con el Host...", vbCritical, "ElisaTaller")
        End
        End If
        Me.lvwResultado.ListItems.Clear
        
        iFila = 1
        
        If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
            AdoPaso.MoveFirst
        End If
        Do While Not AdoPaso.EOF
            
            Paso = AdoPaso!Id_Compañia_Seguro
            Set Item = Me.lvwResultado.ListItems.Add(iFila)
            Item.SubItems(1) = ValorNulo(AdoPaso!Id_Compañia_Seguro)
            Item.SubItems(2) = ValorNulo(AdoPaso!Nombre)
            iFila = iFila + 1
            AdoPaso.MoveNext
            
            Me.lvwResultado.ListItems(1).Selected = True 'Selecciona la primera
            Me.lvwResultado.SetFocus
        Loop
    AdoPaso.Close

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
If Me.lvwResultado.ListItems.Count > 0 Then
    gstrBusca = Me.lvwResultado.SelectedItem.ListSubItems(1)

End If
If gstrBusca <> "" Then
    Unload Me
Else
    r = MsgBox("Primero debe selecionar un Registro", vbInformation, "Buscar Tipo Cargo")
End If
End Sub

Private Sub Form_Activate()
cmdBuscar_Click
End Sub

Private Sub Form_Load()
Dim strSql As String

Me.cmbcoincidir.AddItem "Cualquier Parte del Campo"
Me.cmbcoincidir.AddItem "Todo el Campo"
Me.cmbcoincidir.AddItem "Comienzo del Campo"
Me.cmbcoincidir.AddItem "Final del Campo"
Me.cmbcoincidir = "Cualquier Parte del Campo"

End Sub


Private Sub lvwResultado_DblClick()
gstrBusca = Me.lvwResultado.SelectedItem.ListSubItems(1)
If gstrBusca <> "" Then
    Unload Me
Else
    r = MsgBox("Primero debe selecionar un Registro", vbInformation, "Buscar Item")
End If
End Sub

Private Sub optcodigo_Click()
Me.txtCodigo = ""
Me.txtDescripcion = ""
Me.txtCodigo.SetFocus
End Sub

Private Sub optdescripcion_Click()
Me.txtCodigo = ""
Me.txtDescripcion = ""
Me.txtDescripcion.SetFocus
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Me.cmdBuscar.SetFocus
End If
End Sub

