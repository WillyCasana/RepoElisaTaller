VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditaTempRepuesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repuesto"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmEditaTempRepuesto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3465
      Left            =   30
      TabIndex        =   1
      Top             =   -15
      Width           =   6135
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1875
         Width           =   2100
      End
      Begin VB.TextBox txtDescripcion 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2235
         Width           =   4620
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2610
         Width           =   2115
      End
      Begin VB.TextBox txtActividad 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1275
         Width           =   2100
      End
      Begin VB.TextBox txtServicio 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   11
         Top             =   915
         Width           =   3225
      End
      Begin VB.TextBox txtMarca 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   25
         TabIndex        =   8
         Top             =   165
         Width           =   2595
      End
      Begin VB.TextBox txtModelo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   7
         Top             =   540
         Width           =   3225
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   540
         Left            =   3870
         TabIndex        =   6
         Top             =   2865
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   953
         ButtonWidth     =   1296
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aceptar"
               Key             =   "Ok"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "Cancel"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Close"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         MaxLength       =   25
         TabIndex        =   0
         Top             =   2970
         Width           =   2115
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   5745
         Top             =   105
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
               Picture         =   "frmEditaTempRepuesto.frx":0442
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":09AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":0E04
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":125C
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":136E
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":1480
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":1592
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":16A4
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":17B6
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":18C8
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":19DA
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":1AEC
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":1BFE
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":1D10
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":1E22
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":1F34
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":2046
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":2158
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":226A
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":26BC
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempRepuesto.frx":2B0E
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   180
         X2              =   6045
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   180
         X2              =   6045
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Actividad :"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Servicio :"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   930
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modelo :"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   2655
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   3015
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   1905
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Repuesto:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   2280
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEditaTempRepuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSql As String
Dim AdoPrincipal As New ADODB.Recordset

Private Sub tlbBotones_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
Case "Ok"
    If gstrProcedencia <> "Movimientos" Then
        With frmTempServiciosMarMod
            UPDATEREPUESTO .dtcMarca.BoundText, .dtcModelo.BoundText, .lvwServicios.SelectedItem, .lvwActividades.SelectedItem, .lvwRepuestos.SelectedItem, CDbl(txtCantidad), CCur(txtValor)
            Unload Me
            .Repuestos_de_la_Actividad .dtcMarca.BoundText, .dtcModelo.BoundText, .lvwServicios.SelectedItem, .lvwActividades.SelectedItem
        End With
    Else
        frmRecepcion.lvwRepuestosMantencion.SelectedItem.SubItems(2) = FormatoValor(txtCantidad, "", 1)
        frmRecepcion.lvwRepuestosMantencion.SelectedItem.SubItems(3) = FormatoValor(txtValor, "", gintDecimalesMoneda)
        Unload Me
    End If
    
Case "Cancel"

Case "Close"
    Unload Me
End Select
End Sub



Sub UPDATEREPUESTO(strMarca As String, strModelo As String, strServicio As String, strActividad As String, strRepuesto As String, intCantidad As Double, curValor As Currency)

    mstrSql = "UPDATE Tllr_Actividad_Repuesto"
    mstrSql = mstrSql & " SET Cantidad = " & intCantidad & " , Valor= " & curValor & " "
    mstrSql = mstrSql & " WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' AND Id_Actividad = '" & strActividad & "' AND Id_Item = '" & strRepuesto & "' "
    
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
       ' MsgBox "Si"
    Else
        MsgBox "No"
    End If
    
End Sub
