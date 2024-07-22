VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditaTempServicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicios"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmEditaTempServicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   30
      TabIndex        =   4
      Top             =   -15
      Width           =   6360
      Begin VB.OptionButton optObjeto 
         Caption         =   "Carrocería"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4830
         TabIndex        =   15
         Top             =   885
         Width           =   1140
      End
      Begin VB.OptionButton optObjeto 
         BackColor       =   &H8000000A&
         Caption         =   "Mecánica"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3405
         TabIndex        =   14
         Top             =   885
         Width           =   1005
      End
      Begin VB.TextBox txtMarca 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1125
         MaxLength       =   25
         TabIndex        =   11
         Top             =   165
         Width           =   2595
      End
      Begin VB.TextBox txtModelo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   10
         Top             =   525
         Width           =   3225
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Left            =   3855
         TabIndex        =   9
         Top             =   1905
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   582
         ButtonWidth     =   1826
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aceptar"
               Key             =   "Aceptar"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "Cancelar"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1125
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1935
         Width           =   2000
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1125
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1575
         Width           =   2000
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1230
         Width           =   5115
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   0
         Top             =   885
         Width           =   1980
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   3300
         Top             =   1650
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
               Picture         =   "frmEditaTempServicio.frx":0442
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":09AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":0E04
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":125C
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":136E
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":1480
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":1592
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":16A4
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":17B6
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":18C8
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":19DA
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":1AEC
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":1BFE
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":1D10
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":1E22
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":1F34
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":2046
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":2158
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":226A
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":26BC
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempServicio.frx":2B0E
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   195
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   555
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Top             =   1950
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Horas:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   1605
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   1245
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   930
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmEditaTempServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrSql As String
Dim strObjeto As String * 1
Sub LimpiarServicioNuevo()
    txtCodigo = ""
    txtDescripcion = ""
    txtHoras = ""
    txtValor = ""
End Sub

Private Sub Form_Load()
'txtMarca = frmServiciosPorModelo.dtcMarca.Text
'txtModelo = frmServiciosPorModelo.dtcModelo.Text
End Sub


Private Sub optObjeto_Click(Index As Integer)
Select Case Index
    Case 0
        strObjeto = IIf(optObjeto(0).Value = True, "M", "C")
    Case 1
        strObjeto = IIf(optObjeto(1).Value = True, "C", "M")
End Select
End Sub

Private Sub tlbBotones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Aceptar"
        txtHoras_LostFocus
        
        mstrSql = "UPDATE TLLR_SERVICIO_MODELO "
        mstrSql = mstrSql & " SET Horas=" & txtHoras & " , Valor=" & txtValor & " "
        mstrSql = mstrSql & " WHERE Id_Marca= '" & frmTempServiciosMarMod.dtcMarca.BoundText & "' AND "
        mstrSql = mstrSql & " Id_Modelo='" & frmTempServiciosMarMod.dtcModelo.BoundText & "' AND "
        mstrSql = mstrSql & " Id_Servicio='" & txtCodigo & "'"
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        Set glsiItem = frmTempServiciosMarMod.lvwServicios.SelectedItem
        glsiItem.SubItems(1) = txtDescripcion
        glsiItem.SubItems(2) = txtHoras
        glsiItem.SubItems(3) = Format$(txtValor, "##,###.#0")
        glsiItem.SubItems(4) = IIf(strObjeto = "M", "MECANICA", "CARROCERIA")
        LimpiarServicioNuevo
        Unload Me
    Case "Cancelar"
        LimpiarServicioNuevo
        Unload Me
End Select

End Sub

Private Sub txtHoras_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHoras, strDot)
End Sub

Private Sub txtHoras_LostFocus()
txtValor = gcurPrecioManoObra * CDbl(txtHoras)
End Sub
