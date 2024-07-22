VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAsignacionRecursosColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ajuste de Colores"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2685
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsignacionRecursosColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdAcepar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdForeColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   315
   End
   Begin VB.CommandButton cmdForeColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   13
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton cmdForeColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   12
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton cmdForeColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   11
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdForeColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdBackColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   1920
      TabIndex        =   9
      Top             =   1560
      Width           =   315
   End
   Begin VB.CommandButton cmdBackColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   315
   End
   Begin VB.CommandButton cmdBackColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton cmdBackColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdBackColor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   315
   End
   Begin MSComDlg.CommonDialog cdColores 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblColores 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lblColores 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Día No Laboral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label lblColores 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Día Domingo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label lblColores 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Día Sábado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1740
   End
   Begin VB.Label lblColores 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Día Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1740
   End
End
Attribute VB_Name = "frmAsignacionRecursosColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcepar_Click()
    GrabarDatos
    Unload Me
End Sub

Private Sub cmdBackColor_Click(Index As Integer)
    With Me.cdColores
        Err.Clear
        On Error GoTo Error
        .CancelError = True
        .DialogTitle = "BackColor " & Me.lblColores(Index).Caption
        .Color = Me.lblColores(Index).BackColor
        .ShowColor
        Me.lblColores(Index).BackColor = .Color
    End With
    Exit Sub
Error:
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdForeColor_Click(Index As Integer)
    With Me.cdColores
        Err.Clear
        On Error GoTo Error
        .CancelError = True
        .DialogTitle = "ForeColor " & Me.lblColores(Index).Caption
        .Color = Me.lblColores(Index).ForeColor
        .ShowColor
        Me.lblColores(Index).ForeColor = .Color
    End With
    Exit Sub
Error:
End Sub

Private Sub Form_Load()
    CargaColores
    Screen.MousePointer = vbDefault
End Sub
Private Sub CargaColores()
    Dim strSql As String
    Dim AdoTemp As New ADODB.Recordset

    strSql = "select * from Tllr_Hoja_Recursos_Colores where Id_Empresa='" & gstrIdEmpresa & "' and id_usuario='" & gstrIdUsuario & "'"
    If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            With AdoTemp
                Me.lblColores(0).BackColor = !BackColorNormal
                Me.lblColores(0).ForeColor = !ForeColorNormal

                Me.lblColores(1).BackColor = !BackColorSabado
                Me.lblColores(1).ForeColor = !ForeColorSabado

                Me.lblColores(2).BackColor = !BackColorDomingo
                Me.lblColores(2).ForeColor = !ForeColorDomingo

                Me.lblColores(3).BackColor = !BackColorFestivos
                Me.lblColores(3).ForeColor = !ForeColorFestivos

                Me.lblColores(4).BackColor = !BackColorTotales
                Me.lblColores(4).ForeColor = !ForeColorTotales
            End With
        Else
            '//Crea colores predeterminados...
        End If
    End If
End Sub
Private Sub GrabarDatos()
    Dim strSql As String
    Dim AdoTemp As New ADODB.Recordset
    
    strSql = "update Tllr_Hoja_Recursos_Colores set "
    strSql = strSql & "BackColorNormal = " & Me.lblColores(0).BackColor & ", "
    strSql = strSql & "ForeColorNormal = " & Me.lblColores(0).ForeColor & ", "
    strSql = strSql & "BackColorSabado = " & Me.lblColores(1).BackColor & ", "
    strSql = strSql & "ForeColorSabado = " & Me.lblColores(1).ForeColor & ", "
    strSql = strSql & "BackColorDomingo = " & Me.lblColores(2).BackColor & ", "
    strSql = strSql & "ForeColorDomingo = " & Me.lblColores(2).ForeColor & ", "
    strSql = strSql & "BackColorFestivos = " & Me.lblColores(3).BackColor & ", "
    strSql = strSql & "ForeColorFestivos = " & Me.lblColores(3).ForeColor & ", "
    strSql = strSql & "BackColorTotales = " & Me.lblColores(4).BackColor & ", "
    strSql = strSql & "ForeColorTotales = " & Me.lblColores(4).ForeColor & " "
    strSql = strSql & " where id_empresa='" & gstrIdEmpresa & "' and id_usuario='" & gstrIdUsuario & "'"
    
    If Conexion.SendHost(strSql, , , , 10) <> apOk Then
        MsgBox "Los datos no fueron actualizados...", vbInformation, "Advertencia"
    Else
        frmAsignacionRecursos.Tag = "S"
    End If
End Sub
