VERSION 5.00
Begin VB.Form FrmIngtareas 
   Caption         =   "Cronómetro de Tareas"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   Icon            =   "frmIngtareas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdinicio 
      Appearance      =   0  'Flat
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Cmdsuspension 
      Appearance      =   0  'Flat
      Caption         =   "Suspención"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Cmdtermino 
      Appearance      =   0  'Flat
      Caption         =   "Termino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frmtareas 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.TextBox textarea 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   315
         Index           =   1
         Left            =   4080
         Picture         =   "frmIngtareas.frx":179A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Horas :"
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
         Left            =   360
         TabIndex        =   24
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblTotalHoras 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   23
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Hora Termino :"
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
         TabIndex        =   22
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Termino :"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Hora Inicio :"
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
         Left            =   480
         TabIndex        =   20
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Inicio :"
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
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblHoraTermino 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblFechaTermino 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblHoraInicio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblFechaInicio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblMecanico 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Servicio :"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblServicio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label lblestado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado :"
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
         Left            =   720
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "OT :"
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
         Left            =   840
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Tarea :"
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
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Mecanico :"
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
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmIngtareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codi_mecanico As String
Dim servicio As String
Dim Cod_Servicio As String
Dim estado As String
Dim Nom_Mecanico As String
Dim Cod_Mecanico As String
Dim cod_tarea As String
Dim estado_tarea As String
Dim tipo_tarea As String
Dim Hora_Inicio As String
Dim calcula_minutos As String
Dim calcula_horas As Double
Dim strIdItem As String


Sub LimpiaCampos()
    Me.textarea.Text = ""
    Me.lblMecanico = ""
    Me.lblestado = ""
    Me.lblot = ""
    Me.lblServicio = ""
    Me.lblFechaInicio = ""
    Me.lblHoraInicio = ""
    Me.lblFechaTermino = ""
    Me.lblHoraTermino = ""
    Me.lblTotalHoras = ""
End Sub

Private Sub BarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            'LimpiaConsultaStock
        Case "Buscar"
            'ConsultarStock
        Case "Imprimir"
            'If ImprimirConsultaStock Then
            'End If
        Case "Configuracion"
            'frmConfiguraciondeVistaStock.Show vbModal
        Case "Cerrar"
            Unload Me
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
frmbuscarTareas.Show vbModal
End Sub

Private Sub cmdinicio_Click()
    Hora_Inicio = Format$(Time, "HH:mm")
    graba_en_ot "I"
    graba_en_ordentrabajo "I"
    MsgBox "Esta tarea a sido activada con fecha " & Format(Now, "DD/MM/YYYY") & " a las " & Format(Time, "HH:MM")
    LimpiaCampos
    BloqueaBotones "N"
End Sub

Sub graba_en_ordentrabajo(pstrEstado As String)
Dim lstrSql As String

lstrSql = "INSERT INTO tllr_ordenes_trabajo " & " (id_tarea,id_item,estado,fech_inicio,fech_termino,hora_inicio,hora_termino,total_horas)" _
        & "VALUES ('" & cod_tarea & "', '" & CorrelativoItem & "','" & pstrEstado & "','" & Format(Now, "DD/MM/YYYY") & "','" & Format(Now, "DD/MM/YYYY") & "','" & Hora_Inicio & "','" & Hora_Inicio & "'," & 0 & ")"
        
Conexion.SendHost lstrSql, , , , gcTiempoEspera
End Sub

Sub graba_en_ot(pstrEstado As String)
    'tempario
    strSql = "UPDATE TLLR_MECANICA_OT SET estado_tarea = '" & pstrEstado & "',"
    strSql = strSql & " HorasReales=" & HorasRealesTarea
    strSql = strSql & " WHERE tllr_mecanica_ot.id_tarea = '" & cod_tarea & "' and tllr_mecanica_ot.mecanico_designado =  '" & Cod_Mecanico & "'"
    Conexion.SendHost strSql, , , , gcTiempoEspera
    
    'otros servicios
    strSql = "UPDATE TLLR_OTRO_OT SET estado_tarea = '" & pstrEstado & "',"
    strSql = strSql & " HorasReales=" & HorasRealesTarea
    strSql = strSql & " WHERE tllr_Otro_ot.id_tarea = '" & cod_tarea & "' and tllr_otro_ot.mecanico_asignado =  '" & Cod_Mecanico & "'"
    Conexion.SendHost strSql, , , , gcTiempoEspera
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Cmdsuspension_Click()
Dim hora_final As String
Dim minutos_final As String
Dim Hora_Ini As String
Dim Minutos_ini As String
Dim Total_Horas As Double
Dim Hora_Termino As String
Dim Fecha_Termino As Date

Hora_Termino = Format$(Time, "HH:mm")

hora_final = Hour(Hora_Termino)
minutos_final = Minute(Hora_Termino)

Hora_Ini = Hour(Hora_Inicio)
Minutos_ini = Minute(Hora_Inicio)

calcula_horas = (hora_final - Hora_Ini) * 60
calcula_minutos = (minutos_final - Minutos_ini)

Total_Horas = Round((calcula_horas + calcula_minutos) / 60, 2)

'Actualiza tabla de tareas
Actualiza_Ordenes_Trabajo "S", Hora_Termino, Format(Now, "DD/MM/YYYY"), Total_Horas
graba_en_ot "S"
MsgBox "Esta tarea a sido suspendida con fecha " & Format(Now, "DD/MM/YYYY") & " a las " & Format$(Time, "HH:mm")
LimpiaCampos
BloqueaBotones "N"
End Sub

Private Sub Cmdtermino_Click()
Dim hora_final As String
Dim minutos_final As String
Dim Hora_Ini As String
Dim Minutos_ini As String
Dim Total_Horas As Double
Dim Hora_Termino As String
Dim Fecha_Termino As Date

If tipo_tarea = "S" Then
    graba_en_ot "T"
    BloqueaBotones "T"
    'actualiza estado de ordenes de trabajo
    strSql = "UPDATE TLLR_Ordenes_Trabajo SET estado = 'T'"
    strSql = strSql & " WHERE id_tarea = '" & Me.textarea & "' and id_item='" & strIdItem & "'"
    Conexion.SendHost strSql, , , , gcTiempoEspera
Else
    Hora_Termino = Format$(Time, "HH:mm")
    'FrmordenesTrabajo.IvwDetalleordenes.SelectedItem.SubItems(5) = hora_termino
    
    hora_final = Hour(Hora_Termino)
    minutos_final = Minute(Hora_Termino)
    
    Hora_Ini = Hour(Hora_Inicio)
    Minutos_ini = Minute(Hora_Inicio)
    
    calcula_horas = (hora_final - Hora_Ini) * 60
    calcula_minutos = (minutos_final - Minutos_ini)
    
    Total_Horas = Round((calcula_horas + calcula_minutos) / 60, 2)
    
    'Actualiza tabla de tareas
    Actualiza_Ordenes_Trabajo "T", Hora_Termino, Format(Now, "DD/MM/YYYY"), Total_Horas
    graba_en_ot "T"
    MsgBox " Esta tarea a sido terminada con fecha " & Format(Now, "dd/mm/yyyy") & " y su tiempo de duracion fue " & Total_Horas
    LimpiaCampos
    BloqueaBotones "N"
End If
End Sub

Private Sub Form_Activate()
    
    If Not Atributos("Glbl", "Tllr_20_0130", False, False, False, False) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If '/////////ojo
    
    cmdinicio.Enabled = False
    Cmdsuspension.Enabled = False
    Cmdtermino.Enabled = False

End Sub

Private Sub textarea_KeyPress(KeyAscii As Integer)
Dim AdoTemp As New ADODB.Recordset
Dim mstrSql As String
Dim SW As Integer
Dim esta As String
Dim ExisteTarea As Boolean
Dim Hora_Termino As String
Dim Fecha_Inicio As String
Dim Fecha_Termino As String

If KeyAscii = 13 Then
    Screen.MousePointer = vbHourglass
    
    ExisteTarea = False
    lblestado = ""
    lblot = ""
    lblServicio = ""
    lblMecanico = ""
    lblFechaInicio = ""
    lblFechaTermino = ""
    lblHoraInicio = ""
    lblHoraTermino = ""
    lblTotalHoras = ""
    
    cod_tarea = textarea
   
    mstrSql = "SELECT Id_ot,id_servicio,estado_tarea,mecanico_designado FROM Tllr_mecanica_OT where id_tarea='" & cod_tarea & "'"
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            lblot = ValorNulo(AdoTemp!Id_OT)
            Cod_Servicio = ValorNulo(AdoTemp!Id_servicio)
            tipo_tarea = ValorNulo(AdoTemp!estado_tarea)
            Cod_Mecanico = AdoTemp!mecanico_designado
            lblMecanico = Retorna_Valor_General("Select nombre from Tllr_Mecanicos where id_empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "' And Id_Mecanico='" & AdoTemp!mecanico_designado & "'", gcdynamic)
            ExisteTarea = True
        End If
    End If
    Conexion.CloseHost AdoTemp
    
    mstrSql = "SELECT Id_ot,estado_tarea,mecanico_asignado,Descripcion_Otro FROM Tllr_Otro_OT where id_tarea='" & cod_tarea & "'"
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            lblot = ValorNulo(AdoTemp!Id_OT)
            lblServicio = AdoTemp!Descripcion_Otro
            Cod_Servicio = ""
            tipo_tarea = ValorNulo(AdoTemp!estado_tarea)
            Cod_Mecanico = AdoTemp!Mecanico_Asignado
            lblMecanico = Retorna_Valor_General("Select nombre from Tllr_Mecanicos where id_empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "' And Id_Mecanico='" & AdoTemp!Mecanico_Asignado & "'", gcdynamic)
            ExisteTarea = True
        End If
    End If
    Conexion.CloseHost AdoTemp
    
    If ExisteTarea Then
        If Cod_Servicio <> "" Then
            mstrSql = "SELECT descripcion FROM Tllr_Servicio where id_servicio='" & Cod_Servicio & "'"
            If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not AdoTemp.BOF And Not AdoTemp.EOF Then
                    servicio = ValorNulo(AdoTemp!Descripcion)
                    lblServicio = AdoTemp!Descripcion
                End If
            End If
            Conexion.CloseHost AdoTemp
        End If
        
        mstrSql = "SELECT top 1 * FROM Tllr_ordenes_Trabajo where id_tarea='" & cod_tarea & "' order by id_Item desc"
        If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            If Not AdoTemp.BOF And Not AdoTemp.EOF Then
                esta = ValorNulo(AdoTemp!estado)
                cod_tarea = ValorNulo(AdoTemp!Id_tarea)
                tipo_tarea = ValorNulo(AdoTemp!estado)
                Hora_Inicio = ValorNulo(AdoTemp!Hora_Inicio)
                strIdItem = ValorNulo(AdoTemp!Id_Item)
                
                If tipo_tarea = "I" Then
                    Me.lblFechaInicio = ValorNulo(AdoTemp!fech_inicio)
                    Me.lblHoraInicio = Format(ValorNulo(AdoTemp!Hora_Inicio), "HH:MM")
                ElseIf tipo_tarea = "S" Then
                    Me.lblFechaInicio = ValorNulo(AdoTemp!fech_inicio)
                    Me.lblHoraInicio = Format(ValorNulo(AdoTemp!Hora_Inicio), "HH:MM")
                    Me.lblFechaTermino = ValorNulo(AdoTemp!fech_termino)
                    Me.lblHoraTermino = Format(ValorNulo(AdoTemp!Hora_Termino), "HH:MM")
                    Me.lblTotalHoras = ValorNulo(AdoTemp!Total_Horas)
                ElseIf tipo_tarea = "T" Then
                    Me.lblFechaInicio = ValorNulo(AdoTemp!fech_inicio)
                    Me.lblHoraInicio = Format(ValorNulo(AdoTemp!Hora_Inicio), "HH:MM")
                    Me.lblFechaTermino = ValorNulo(AdoTemp!fech_termino)
                    Me.lblHoraTermino = Format(ValorNulo(AdoTemp!Hora_Termino), "HH:MM")
                    Me.lblTotalHoras = ValorNulo(AdoTemp!Total_Horas)
                End If
            End If
        End If
        Conexion.CloseHost AdoTemp
        
        BloqueaBotones esta
    Else
        BloqueaBotones "N"
        MsgBox "Esta numero de Tarea no esta Asignada", vbExclamation, "Tareas"
    End If
    Screen.MousePointer = vbDefault
End If
    
End Sub

Sub BloqueaBotones(pstrEstado As String)
    If pstrEstado = "I" Then
        cmdinicio.Enabled = False
        Cmdsuspension.Enabled = True
        Cmdtermino.Enabled = True
        lblestado = "INICIADO"
    ElseIf pstrEstado = "T" Then
        Me.cmdinicio.Enabled = False
        Me.Cmdsuspension.Enabled = False
        Me.Cmdtermino.Enabled = False
        Me.lblestado = "TERMINADO"
    ElseIf pstrEstado = "S" Then
        Me.cmdinicio.Enabled = True
        Me.Cmdsuspension.Enabled = False
        Me.Cmdtermino.Enabled = True
        Me.lblestado = "SUSPENDIDO"
    ElseIf pstrEstado = "N" Then
        Me.cmdinicio.Enabled = False
        Me.Cmdsuspension.Enabled = False
        Me.Cmdtermino.Enabled = False
        Me.lblestado = ""
    Else
        Me.cmdinicio.Enabled = True
        Me.Cmdsuspension.Enabled = False
        Me.Cmdtermino.Enabled = False
        Me.lblestado = "VIGENTE"
    End If
End Sub
Function CorrelativoItem() As Integer
Dim strSql As String
Dim AdoTemp As New ADODB.Recordset

    strSql = "Select max(Id_Item) as item From Tllr_Ordenes_Trabajo where Id_tarea='" & Me.textarea & "'"
    If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            CorrelativoItem = IIf(IsNull(AdoTemp!Item), 1, AdoTemp!Item + 1)
        End If
    End If
    
Conexion.CloseHost AdoTemp
End Function

Sub Actualiza_Ordenes_Trabajo(pstrEstado As String, ptHoraTermino As String, pdfechaTermino As Date, pTotalHoras As Double)
    strSql = "UPDATE TLLR_Ordenes_Trabajo SET estado = '" & pstrEstado & "',"
    strSql = strSql & " Hora_Termino='" & ptHoraTermino & "',"
    strSql = strSql & " Fech_Termino='" & pdfechaTermino & "',"
    strSql = strSql & " Total_horas=" & pTotalHoras
    strSql = strSql & " WHERE id_tarea = '" & Me.textarea & "' and id_item='" & strIdItem & "'"
    Conexion.SendHost strSql, , , , gcTiempoEspera

End Sub
Function HorasRealesTarea() As Double
Dim strSql As String
Dim AdoTemp As New ADODB.Recordset

    strSql = "Select sum(Total_Horas) as HorasReales From Tllr_Ordenes_Trabajo where Id_tarea='" & Me.textarea & "'"
    If Conexion.SendHost(strSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not AdoTemp.BOF And Not AdoTemp.EOF Then
            HorasRealesTarea = IIf(IsNull(AdoTemp!HorasReales), 0, Round(AdoTemp!HorasReales, 2))
        End If
    End If
    
Conexion.CloseHost AdoTemp
End Function

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            'LimpiaConsultaStock
        Case "Buscar"
            'ConsultarStock
        Case "Imprimir"
            'If ImprimirConsultaStock Then
            'End If
        Case "Configuracion"
            'frmConfiguraciondeVistaStock.Show vbModal
        Case "Cerrar"
            Unload Me
    End Select
    Screen.MousePointer = vbDefault
End Sub


