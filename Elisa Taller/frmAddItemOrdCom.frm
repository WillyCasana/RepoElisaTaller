VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddItemOrdCom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Item Orden Compra"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmAddItemOrdCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6120
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox txtPreUni 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   2
         Top             =   840
         Width           =   1170
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   0
         Top             =   150
         Width           =   4875
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1185
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad :"
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   8
         Top             =   525
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Costo:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   870
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   6
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Sub - Total :"
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   5
         Top             =   1275
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      DisabledImageList=   "ImgBarraHerramienta"
      HotImageList    =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar a la Lista"
            ImageKey        =   "Editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar Modo Edición"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":18AC
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":19BE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":1AD0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":1BE2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":1CF4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":1E06
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":1F18
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":202A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":213C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":224E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":2360
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":2472
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":2584
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":2696
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":27A8
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":28BA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":2D0C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":315E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddItemOrdCom.frx":3270
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddItemOrdCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Nuevo"
    
    With Me
        .txtDescripcion = ""
        .txtCantidad = ""
        .txtPreUni = ""
        .txtSubTotal = ""
        .txtDescripcion.SetFocus
    End With

Case "Agregar"

    With frmEmisionOrdCom
        Set glsiItem = .lvwDetalle.ListItems.Add(, , .lvwDetalle.ListItems.Count + 1)
        glsiItem.SubItems(1) = IIf(txtDescripcion <> "", txtDescripcion, "Sin Descripción")
        glsiItem.SubItems(2) = FormatoValor(txtCantidad, "", gintDecimalesMoneda)
        glsiItem.SubItems(3) = FormatoValor(txtPreUni, "", gintDecimalesMoneda)
        glsiItem.SubItems(4) = FormatoValor(txtSubTotal, "", gintDecimalesMoneda)
    End With

Case "Cerrar"

    Unload Me

End Select

End Sub

Private Sub txtCantidad_GotFocus()
'txtCantidad = SacarFormatoValor(txtCantidad, "")
MarcaTexto txtCantidad
End Sub

Private Sub txtCantidad_LostFocus()
'txtCantidad = FormatoValor(txtCantidad, "", 0)
txtSubTotal = Val(txtCantidad) * Val(txtPreUni)


End Sub

Private Sub txtdescripcion_GotFocus()
MarcaTexto txtDescripcion
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys (Chr(9))
End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)

End Sub


Private Sub txtPreUni_GotFocus()
'txtPreUni = SacarFormatoValor(txtPreUni, "")
MarcaTexto txtPreUni
End Sub

Private Sub txtPreUni_LostFocus()

'txtPreUni = FormatoValor(txtPreUni, "", 0)
txtSubTotal = Val(txtCantidad) * Val(txtPreUni)
End Sub


Private Sub txtSubTotal_GotFocus()
'txtSubTotal = SacarFormatoValor(txtSubTotal, "")
MarcaTexto txtSubTotal
End Sub


Private Sub txtSubTotal_LostFocus()
'txtSubTotal = FormatoValor(txtSubTotal, "", 0)
End Sub


