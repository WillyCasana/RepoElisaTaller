VERSION 5.00
Begin VB.Form frmAddOpciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones de Sistema"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCodigos 
      Height          =   2760
      Left            =   4950
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   75
      Width           =   1020
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cerrar"
      Height          =   420
      Index           =   1
      Left            =   3345
      TabIndex        =   2
      Top             =   2895
      Width           =   1215
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Seleccionar"
      Height          =   420
      Index           =   0
      Left            =   2055
      TabIndex        =   1
      Top             =   2895
      Width           =   1215
   End
   Begin VB.ListBox lstOpciones 
      Height          =   2760
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   75
      Width           =   4455
   End
End
Attribute VB_Name = "frmAddOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As ADODB.Recordset
Dim mstrSql As String
Dim intIndice As Integer

Private Sub cmdBotones_Click(Index As Integer)
Select Case Index
Case 0
    For intIndice = 0 To lstOpciones.ListCount - 1
        If lstOpciones.Selected(intIndice) = True Then
            If frmMantenedorPerfil.VerificaPrivilegio(Me.lstCodigos.List(intIndice), frmMantenedorPerfil.lstCodigos) = True Then
                frmMantenedorPerfil.AgregaPrivilegio frmMantenedorPerfil.lstCodigos.ListCount, frmAddOpciones.lstCodigos.List(intIndice), frmAddOpciones.lstOpciones.List(intIndice), "S", "S", "S", "S", "S"
            End If
        End If
    Next
Case 1
    Unload Me
End Select

End Sub

Private Sub Form_Activate()
LlenaOpciones
End Sub

Private Sub LlenaOpciones()
    Set adoPrincipal = New ADODB.Recordset
    mstrSql = "SELECT Id_Opcion, Descripcion From Tllr_Opcion_Sistema Where Vigencia = 'S' "
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                .MoveLast: .MoveFirst
                While Not .EOF
                    lstOpciones.AddItem !Descripcion
                    lstCodigos.AddItem !id_opcion
                    .MoveNext
                Wend
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Private Sub lstOpciones_Click()
If lstOpciones.ListIndex <> -1 Then
    lstCodigos.ListIndex = lstOpciones.ListIndex
    If lstOpciones.Selected(lstOpciones.ListIndex) = True Then
        lstCodigos.Selected(lstOpciones.ListIndex) = True
    Else
        lstCodigos.Selected(lstOpciones.ListIndex) = False
    End If
End If
End Sub


