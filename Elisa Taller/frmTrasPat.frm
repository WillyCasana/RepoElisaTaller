VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrasPat 
   Caption         =   "Patente to Rut"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pgbEstado 
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel And Exit"
      Height          =   495
      Left            =   2490
      TabIndex        =   1
      Top             =   1185
      Width           =   2070
   End
   Begin VB.CommandButton cmdExe 
      Caption         =   "Execute Convert"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2070
   End
End
Attribute VB_Name = "frmTrasPat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExe_Click()
Dim strDig As String
Dim strRut As String

gstrSql = "Select Patente from Tllr_Vehiculo_Cliente"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
With gadoPrincipal
    If Not .BOF And Not .EOF Then
        pgbEstado.Max = .RecordCount
        .MoveFirst
        While Not .EOF
            'Call CheckPatente(ValorNulo(!Patente), strDig, strRut)
            gstrSql = "UPDATE Tllr_Vehiculo_Cliente"
            gstrSql = gstrSql & " SET RutVehiculo= '" & strRut & "' "
            gstrSql = gstrSql & " where Patente='" & !Patente & "' "
            Conexion.SendHost gstrSql, , , , gcTiempoEspera
            pgbEstado.Value = pgbEstado.Value + 1
            .MoveNext
        Wend
    End If
End With
End If

End Sub

Private Sub Form_Load()
pgbEstado.Min = 0

End Sub
