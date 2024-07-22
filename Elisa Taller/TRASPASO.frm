VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TRASPASO 
   Caption         =   "TRASPASO"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar BARRA 
      Height          =   405
      Left            =   60
      TabIndex        =   1
      Top             =   735
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EJECUTAR"
      Height          =   495
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   3225
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7830
      TabIndex        =   3
      Top             =   120
      Width           =   210
   End
   Begin VB.Label PORCENTAJE 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3405
      TabIndex        =   2
      Top             =   105
      Width           =   4320
   End
End
Attribute VB_Name = "TRASPASO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MSTRSQL As String
Dim ADOMODELO As New ADODB.Recordset
Dim ADOSERVICIO As New ADODB.Recordset
Dim ADOSERVICIOMODELO As New ADODB.Recordset

Function AVANCE(TOTAL As Long, ACTUAL As Long) As Single


End Function


Private Sub Command1_Click()
TRASPASO
End Sub

Sub TRASPASO()
BARRA.Min = 0
BARRA.Value = 0
MSTRSQL = "SELECT ID_SERVICIO FROM TLLR_SERVICIO ORDER BY ID_SERVICIO"
If Conexion.SendHost(MSTRSQL, ADOSERVICIO, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not ADOSERVICIO.BOF And Not ADOSERVICIO.EOF Then
        ADOSERVICIO.MoveFirst
        MSTRSQL = "SELECT ID_MARCA,ID_MODELO FROM GLBL_MODELO ORDER BY ID_MARCA,ID_MODELO"
        If Conexion.SendHost(MSTRSQL, ADOMODELO, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            If Not ADOMODELO.BOF And Not ADOMODELO.EOF Then
                ADOMODELO.MoveFirst
                Dim TOTAL As Long
                
                TOTAL = ADOMODELO.RecordCount * ADOSERVICIO.RecordCount
                BARRA.Max = TOTAL
                
                While Not ADOMODELO.EOF
                    ADOSERVICIO.MoveFirst
                    While Not ADOSERVICIO.EOF
                        MSTRSQL = "INSERT INTO TLLR_SERVICIO_MODELO (ID_MARCA,ID_MODELO, ID_SERVICIO,VALOR,HORAS) VALUES('" & ADOMODELO!ID_MARCA & "','" & ADOMODELO!ID_MODELO & "','" & ADOSERVICIO!ID_SERVICIO & "',0,0)"
                        If Conexion.SendHost(MSTRSQL, , , , gcTiempoEspera) = apOk Then
                            DoEvents
                            BARRA.Value = BARRA.Value + 1
                            Me.Caption = "REGISTROS ACTUALIZADOS :" & Format(BARRA.Value, "###,###")
                            PORCENTAJE.Caption = Format$((BARRA.Value * 100) / TOTAL)
                            ADOSERVICIO.MoveNext
                        Else
                            Exit Sub
                        End If
                    Wend
                    ADOMODELO.MoveNext
                Wend
            End If
        End If
    End If
End If
End Sub

