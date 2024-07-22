VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEditTempCptoVsCiaSeg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editar Valor"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmEditTempCptoVsCiaSeg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbBotones 
      Height          =   540
      Left            =   4590
      TabIndex        =   11
      Top             =   1710
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   953
      ButtonWidth     =   1296
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      DisabledImageList=   "ImgBarraHerramienta"
      HotImageList    =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar Cambios y Cierra Modo Edición"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancela Cambios y Cierrar Modo Edición"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle Valor"
      Height          =   2340
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6180
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1770
         TabIndex        =   10
         Top             =   1575
         Width           =   1590
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1770
         TabIndex        =   9
         Top             =   2415
         Width           =   1290
      End
      Begin VB.Label lblPartePieza 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1770
         TabIndex        =   8
         Top             =   1155
         Width           =   4200
      End
      Begin VB.Label lblConcepto 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1785
         TabIndex        =   7
         Top             =   735
         Width           =   4200
      End
      Begin VB.Label lblCompañia 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1785
         TabIndex        =   6
         Top             =   330
         Width           =   4200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   5
         Top             =   1620
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Horas"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   2505
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Compañía de Seguro:"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   3
         Top             =   375
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parte - Pieza"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   795
         Width           =   690
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   15
      Top             =   2370
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
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":000C
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":011E
            Key             =   "Menos"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":0576
            Key             =   "Mas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":09CE
            Key             =   "Persona"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":0E26
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":0F38
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":104A
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":115C
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":126E
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":1380
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":1492
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":15A4
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":16B6
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":17C8
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":18DA
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":19EC
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":1AFE
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":1C10
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":1D22
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":1E34
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":2286
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTempCptoVsCiaSeg.frx":26D8
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEditTempCptoVsCiaSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function GrabarTempario(strCia As String, strConcepto As String, strPartePieza As String, strValor As String) As Boolean
Dim strSql As String
Dim AdoTemp As New ADODB.Recordset

strSql = "SELECT * FROM Tllr_CiaSeguro_Concepto_Parte_Pieza WHERE Id_Compañia_Seguro = '" & strCia & "' AND Id_Concepto = '" & strConcepto & "' AND Id_Parte_Pieza = '" & strPartePieza & "'"
If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoTemp
    If Not .BOF And Not .EOF Then
        If strValor = "Valor" Then
            strSql = "Update Tllr_CiaSeguro_Concepto_Parte_Pieza Set Valor=" & CCur(Val(Trim(txtValor))) & " WHERE Id_Compañia_Seguro = '" & strCia & "' AND Id_Concepto = '" & strConcepto & "' AND Id_Parte_Pieza = '" & strPartePieza & "'"
            Conexion.SendHost strSql, , , , gcTiempoEspera
        Else
            strSql = "Update Tllr_CiaSeguro_Concepto_Parte_Pieza Set Horas=" & CCur(Val(Trim(txtValor))) & " WHERE Id_Compañia_Seguro = '" & strCia & "' AND Id_Concepto = '" & strConcepto & "' AND Id_Parte_Pieza = '" & strPartePieza & "'"
            Conexion.SendHost strSql, , , , gcTiempoEspera
        End If
    Else
        If strValor = "Valor" Then
            strSql = "Insert into Tllr_CiaSeguro_Concepto_Parte_Pieza (Id_Compañia_Seguro, Id_Concepto, Id_Parte_Pieza, Valor, Horas) Values ('" & strCia & "' ,'" & strConcepto & "' ,'" & strPartePieza & "'," & CCur(Val(Trim(txtValor))) & ",0)"
            Conexion.SendHost strSql, , , , gcTiempoEspera
        Else
            strSql = "Insert into Tllr_CiaSeguro_Concepto_Parte_Pieza (Id_Compañia_Seguro, Id_Concepto, Id_Parte_Pieza, Valor, Horas) Values ('" & strCia & "' ,'" & strConcepto & "' ,'" & strPartePieza & "',0," & CCur(Val(Trim(txtValor))) & ")"
            Conexion.SendHost strSql, , , , gcTiempoEspera
        End If
    End If
    End With
End If

Conexion.CloseHost AdoTemp

End Function


Sub UpdateTempario()

End Sub

Private Sub tlbBotones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Guardar"
    With Me
        GrabarTempario .lblCompañia.Tag, .lblConcepto.Tag, .lblPartePieza.Tag, .Label1(4).Caption
    End With
    frmTempCiaSeguro.HFlexGrid.Text = txtValor
    Unload Me
Case "Cancelar"
    Unload Me
End Select

End Sub

