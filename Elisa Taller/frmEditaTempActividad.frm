VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditaTempActividad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actividad Nueva"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmEditaTempActividad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2910
      Left            =   30
      TabIndex        =   4
      Top             =   -15
      Width           =   6360
      Begin VB.TextBox txtEspecialidad 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2535
         Width           =   2250
      End
      Begin VB.TextBox txtServicio 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   14
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txtMarca 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1275
         MaxLength       =   25
         TabIndex        =   11
         Top             =   225
         Width           =   2595
      End
      Begin VB.TextBox txtModelo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   10
         Top             =   570
         Width           =   3225
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Left            =   3960
         TabIndex        =   9
         Top             =   2520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ButtonWidth     =   1826
         ButtonHeight    =   582
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
         Left            =   1275
         MaxLength       =   25
         TabIndex        =   3
         Top             =   2205
         Width           =   2000
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1275
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1890
         Width           =   2000
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1575
         Width           =   5025
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1275
         MaxLength       =   25
         TabIndex        =   0
         Top             =   1260
         Width           =   1980
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
               Picture         =   "frmEditaTempActividad.frx":0442
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":09AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":0E04
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":125C
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":136E
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":1480
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":1592
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":16A4
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":17B6
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":18C8
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":19DA
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":1AEC
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":1BFE
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":1D10
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":1E22
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":1F34
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":2046
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":2158
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":226A
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":26BC
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditaTempActividad.frx":2B0E
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Especialidad :"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   16
         Top             =   2550
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Servicio :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   930
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modelo :"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Top             =   2235
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Horas:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   1605
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1275
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmEditaTempActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoPrincipal As New ADODB.Recordset
Dim mstrSql As String


Sub LimpiarNuevaActividad()
With Me
    .txtCodigo = ""
    .txtNombre = ""
    .txtHoras = ""
    .txtValor = ""
    .txtEspecialidad.Text = ""
End With
End Sub


Private Sub tlbBotones_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
    Case "Aceptar"
        With frmTempServiciosMarMod
        mstrSql = "UPDATE TLLR_ACTIVIDAD_SERVICIO_MODELO "
        mstrSql = mstrSql & " SET Horas=" & txtHoras & ", Valor=" & txtValor & " "
        mstrSql = mstrSql & " WHERE Id_Marca ='" & .dtcMarca.BoundText & "' AND "
        mstrSql = mstrSql & " Id_Modelo ='" & .dtcModelo.BoundText & "' AND "
        mstrSql = mstrSql & " Id_Servicio='" & .lvwServicios.SelectedItem & "' AND "
        mstrSql = mstrSql & " Id_Actividad='" & txtCodigo & "' "
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        '/////////////////////////////////////
        Set glsiItem = .lvwActividades.SelectedItem
        glsiItem.SubItems(2) = txtHoras
        glsiItem.SubItems(3) = Format$(txtValor, "##,##0")
        LimpiarNuevaActividad
        Unload Me
        End With
    Case "Cancelar"
        LimpiarNuevaActividad
        Unload Me
End Select

End Sub


'Sub FillEspecialidades()
'    mstrSql = "SELECT Id_Especialidad AS Codigo, Descripcion AS Nombre FROM Tllr_Especialidad WHERE Vigencia = 'S' order by Descripcion"
'    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
'        With datEspecialidad
'            Set .Recordset = adoPrincipal
'            If Not .Recordset.BOF And Not .Recordset.EOF Then
'                .Recordset.MoveFirst
'                dtcEspecialidad.ListField = "Nombre"
'                dtcEspecialidad.BoundColumn = "Codigo"
'                dtcEspecialidad.BoundText = .Recordset!codigo
'                If .Recordset.RecordCount < 2 Then dtcEspecialidad.Enabled = False
'            End If
'        End With
'    End If ' por el otro
'    Set adoPrincipal = New ADODB.Recordset
'    Conexion.CloseHost adoPrincipal
'End Sub
Private Sub txtHoras_Change()

End Sub

Private Sub txtHoras_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtValor.SetFocus
End If
End Sub

Private Sub txtHoras_LostFocus()
    txtValor.Text = CDbl(txtHoras) * gcurPrecioManoObra
End Sub
