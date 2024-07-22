VERSION 5.00
Begin VB.Form frmReemplazaVinxPatente 
   Caption         =   "Reemplaza Vin por Placa"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "frmReemplazaVinxPatente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frFilaCol 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtVin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vin:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Patente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4080
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aplicar"
      Height          =   300
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmReemplazaVinxPatente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim lstrSql As String
Dim Crear As String

Crear = "S"

If txtVin = "" Then
    MsgBox "El Vin debe contener un valor...", vbInformation, "Advertencia"
    txtVin.SetFocus
    Exit Sub
End If
If txtPatente = "" Then
    MsgBox "La " & gstrNombrePatente & " debe contener un valor...", vbInformation, "Advertencia"
    txtPatente.SetFocus
    Exit Sub
End If
    
Screen.MousePointer = vbHourglass

'//Verifica si existe la patente
Dim AdoTemp As New ADODB.Recordset
lstrSql = "Select Patente from Tllr_Vehiculo_Cliente Where Patente='" & Me.txtPatente & "'"
If Conexion.SendHost(lstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not AdoTemp.BOF And Not AdoTemp.EOF Then
        Crear = "N"
    End If
End If
Conexion.CloseHost AdoTemp
    
'verifica si existe el Vin
lstrSql = "Select Patente from Tllr_Vehiculo_Cliente Where Patente='" & Me.txtVin & "'"
If Conexion.SendHost(lstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If AdoTemp.BOF And AdoTemp.EOF Then
        MsgBox "El Vin No Existe...   Verifique", vbExclamation, "Advertencia"
        Exit Sub
    End If
End If
Conexion.CloseHost AdoTemp
    
lstrSql = "Exec Tllr_Actualiza_Vin_Placa '" & Me.txtVin & "','" & Me.txtPatente & "','" & gstrIdUsuario & "','" & Format(Now, "DD/MM/YYYY") & "','" & Crear & "'"
If Conexion.SendHost(lstrSql, , , , gcTiempoEspera) = apOk Then
    MsgBox "Proceso Completado con Exito", vbInformation, "Información"
End If

Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not Atributos("Glbl", "Tllr_20_0140", False, False, False, False) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If '/////////ojo

    Caption = "Reemplaza Vin Por " & gstrNombrePatente
    Label2.Caption = gstrNombrePatente
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
    End Select
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'If gstrValidaPatente = "S" Then
'    KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'End If
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub

Private Sub txtVin_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub
