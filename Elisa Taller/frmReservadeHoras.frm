VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReservadeHoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor Reservas de Horas"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmReservadeHoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVin 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   2880
      Width           =   2025
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   9615
      Begin VB.TextBox lblNroRecepcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   2100
      End
      Begin MSComCtl2.DTPicker pckFechaAtencion 
         Height          =   315
         Left            =   4830
         TabIndex        =   4
         Top             =   165
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94502913
         CurrentDate     =   36776
      End
      Begin VB.Label lblCorrelativo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reserva Nº :"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Atención"
         Height          =   195
         Index           =   9
         Left            =   3540
         TabIndex        =   7
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label lblEstadoOT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6540
         TabIndex        =   6
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblEstadoOTValor 
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
         Left            =   7545
         TabIndex        =   5
         Top             =   165
         Width           =   1815
      End
   End
   Begin VB.Frame fmePat 
      Height          =   4815
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   9615
      Begin VB.TextBox txtChasis 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtTaxiDestino 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         TabIndex        =   74
         Top             =   4440
         Width           =   2775
      End
      Begin VB.TextBox txtHora 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8520
         TabIndex        =   72
         Top             =   3000
         Width           =   870
      End
      Begin VB.CommandButton cmdSinPatente 
         Caption         =   "Sin Placa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   53
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSinPatente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Sin Placa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   52
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtRecepcionista 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   50
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   3600
         Width           =   9255
      End
      Begin VB.TextBox txtFonos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   16
         Top             =   2445
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtRut 
         Height          =   315
         Left            =   7290
         MaxLength       =   50
         TabIndex        =   15
         Top             =   5850
         Width           =   1410
      End
      Begin VB.TextBox txtComuna 
         Height          =   315
         Left            =   3870
         MaxLength       =   50
         TabIndex        =   14
         Top             =   5850
         Width           =   3195
      End
      Begin VB.TextBox txtDir 
         Height          =   315
         Left            =   180
         MaxLength       =   50
         TabIndex        =   13
         Top             =   5805
         Width           =   3330
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   915
         MaxLength       =   10
         TabIndex        =   0
         Top             =   345
         Width           =   1200
      End
      Begin VB.TextBox txtAño 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8760
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   12
         Top             =   990
         Width           =   600
      End
      Begin VB.TextBox txtKilAct 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4095
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1410
         Width           =   2160
      End
      Begin VB.ComboBox cboHora 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8445
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3315
         Visible         =   0   'False
         Width           =   990
      End
      Begin MSComCtl2.DTPicker pckFecVta 
         Height          =   315
         Left            =   7905
         TabIndex        =   17
         Top             =   1410
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   94502913
         CurrentDate     =   36796
      End
      Begin MSComCtl2.DTPicker pckFechaEntrega 
         Height          =   315
         Left            =   5895
         TabIndex        =   18
         Top             =   2955
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94502913
         CurrentDate     =   36733
      End
      Begin MSComctlLib.Toolbar tlbPatente 
         Height          =   330
         Left            =   2205
         TabIndex        =   19
         Top             =   315
         Visible         =   0   'False
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo Patente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Patente"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcRecepcionista 
         Bindings        =   "frmReservadeHoras.frx":038A
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Top             =   2955
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datRecepcionista 
         Height          =   330
         Left            =   2790
         Top             =   2955
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar tlbLlamadoTelefono 
         Height          =   330
         Left            =   8280
         TabIndex        =   49
         Top             =   2430
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Llamar"
               Object.ToolTipText     =   "Hacer Llamada Telefonica"
               ImageIndex      =   20
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtSucursal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   51
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblTaxi 
         Caption         =   "Taxi Destino"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Reserva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   14
         Left            =   7275
         TabIndex        =   25
         Top             =   3000
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. OT"
         Height          =   255
         Left            =   7200
         TabIndex        =   48
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNumeroOt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7920
         TabIndex        =   47
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   180
         TabIndex        =   45
         Top             =   3375
         Width           =   1590
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   6
         X1              =   180
         X2              =   9360
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   3
         X1              =   135
         X2              =   9360
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   2
         X1              =   135
         X2              =   9360
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kms. Act."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   33
         Left            =   3240
         TabIndex        =   44
         Top             =   1455
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   4
         X1              =   135
         X2              =   9360
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Venta"
         Height          =   195
         Index           =   6
         Left            =   6915
         TabIndex        =   43
         Top             =   1455
         Width           =   915
      End
      Begin VB.Label lblMotor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4095
         TabIndex        =   42
         Top             =   1875
         Width           =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   210
         X2              =   9360
         Y1              =   2310
         Y2              =   2295
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fonos"
         Height          =   195
         Index           =   32
         Left            =   6000
         TabIndex        =   41
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label lblFono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6960
         TabIndex        =   40
         Top             =   2400
         Width           =   2370
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIN"
         Height          =   195
         Index           =   29
         Left            =   6960
         TabIndex        =   39
         Top             =   1920
         Width           =   270
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chasis"
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   38
         Top             =   1890
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asesor de Servicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   37
         Top             =   3045
         Width           =   1605
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   855
         TabIndex        =   36
         Top             =   2445
         Width           =   4230
      End
      Begin VB.Label lblColorE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   855
         TabIndex        =   35
         Top             =   1410
         Width           =   2295
      End
      Begin VB.Label lblModelo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4095
         TabIndex        =   34
         Top             =   990
         Width           =   3540
      End
      Begin VB.Label lblMarca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   855
         TabIndex        =   33
         Top             =   990
         Width           =   2295
      End
      Begin VB.Label lblPat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   31
         Top             =   990
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   30
         Top             =   1035
         Width           =   525
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
         Height          =   195
         Index           =   3
         Left            =   8310
         TabIndex        =   29
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   28
         Top             =   1455
         Width           =   360
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2475
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fec.Reserva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   13
         Left            =   4725
         TabIndex        =   26
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Motor"
         Height          =   195
         Index           =   21
         Left            =   3240
         TabIndex        =   24
         Top             =   1875
         Width           =   705
      End
      Begin VB.Label lblIdMarca 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   900
         TabIndex        =   23
         Top             =   990
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblIdModelo 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5760
         TabIndex        =   22
         Top             =   990
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblIdCliente 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3825
         TabIndex        =   21
         Top             =   2445
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   5
         X1              =   135
         X2              =   9360
         Y1              =   885
         Y2              =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seccion"
      Height          =   615
      Left            =   12720
      TabIndex        =   69
      Top             =   3240
      Width           =   2415
      Begin VB.OptionButton optCarroceria 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Carroceria"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   71
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optMecanica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Mecanica"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   54
      Top             =   5880
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Servicios Mecánica"
      TabPicture(0)   =   "frmReservadeHoras.frx":03A9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblHorasMecanica"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTotalHoras(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvwServiciosMecanica"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAgregarServiciosMecanica"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdBorrarServiciosMecanica"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Otros Servicios"
      TabPicture(1)   =   "frmReservadeHoras.frx":03C5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "lblHorasOtrosServicios"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "lblTotalHoras(1)"
      Tab(1).Control(4)=   "lvwOtrosServicios"
      Tab(1).Control(5)=   "cmdBorrarOtrosServicios"
      Tab(1).Control(6)=   "cmdCrearOtrosServicios"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdCrearOtrosServicios 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   -74880
         TabIndex        =   60
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdBorrarOtrosServicios 
         Caption         =   "&Borrar"
         Height          =   255
         Left            =   -73680
         TabIndex        =   59
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdBorrarServiciosMecanica 
         Caption         =   "&Borrar"
         Height          =   255
         Left            =   1320
         TabIndex        =   57
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregarServiciosMecanica 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1920
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwServiciosMecanica 
         Height          =   1455
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Servicio"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Horas"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lvwOtrosServicios 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   58
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Servicio"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Horas"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lblTotalHoras 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   315
         Index           =   0
         Left            =   8520
         TabIndex        =   68
         Top             =   1875
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas:"
         Height          =   195
         Left            =   7560
         TabIndex        =   67
         Top             =   1875
         Width           =   870
      End
      Begin VB.Label lblHorasMecanica 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   315
         Left            =   6480
         TabIndex        =   66
         Top             =   1875
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas Servicios Mecanica:"
         Height          =   195
         Left            =   3960
         TabIndex        =   65
         Top             =   1875
         Width           =   2310
      End
      Begin VB.Label lblTotalHoras 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   315
         Index           =   1
         Left            =   -66960
         TabIndex        =   64
         Top             =   1875
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas:"
         Height          =   195
         Left            =   -68040
         TabIndex        =   63
         Top             =   1875
         Width           =   870
      End
      Begin VB.Label lblHorasOtrosServicios 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         Height          =   315
         Left            =   -69000
         TabIndex        =   62
         Top             =   1875
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas Otros Servicios:"
         Height          =   195
         Left            =   -71520
         TabIndex        =   61
         Top             =   1875
         Width           =   1980
      End
   End
   Begin Crystal.CrystalReport rptReserva 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro (Ctrl+G)"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
            ImageKey        =   "Cancelar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro (Ctrl+B)"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
            ImageKey        =   "Primero"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
            ImageKey        =   "Siguiente"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro (Ctrl+U)"
            ImageKey        =   "Ultimo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Anular Reserva"
            ImageKey        =   "Borrar"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Confirmar"
            Object.ToolTipText     =   "Confirmar Hora"
            ImageKey        =   "Seleccion"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Activar"
            Object.ToolTipText     =   "Activar Reserva"
            ImageKey        =   "Seleccion1"
         EndProperty
      EndProperty
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   0
         X2              =   8820
         Y1              =   1980
         Y2              =   1980
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":03E1
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":04F3
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0605
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0717
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0829
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":093B
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0A4D
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0B5F
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0C71
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0D83
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0E95
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":0FA7
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":10B9
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":11CB
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":12DD
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":13EF
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":1501
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":1953
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":1DA5
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":1EB7
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":2013
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":216F
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":22CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":2427
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":2EF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":3347
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":34AB
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":3907
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":3A63
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":4D6F
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":530B
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":5467
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":55C3
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":5917
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":5C6B
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":5FBF
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":6313
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":6667
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":6BAB
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":6CBD
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":7011
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":7365
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":7479
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":77CD
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":78DF
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReservadeHoras.frx":7C33
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReservadeHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoPrincipal As New ADODB.Recordset
Dim mstrSQL As String
Dim mstrWhere As String
Dim mstrOrderBy As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Public mblnSW As Boolean
Dim itmAux As ListItem
Dim intIndice As Integer
Dim mstrCargo As String
Dim mblnBloqueo As Boolean
Dim KilometrajeEntrada As String
Private Sub cmdAgregarServiciosMecanica_Click()
    gstrProcedencia = "Reserva_Horas"
    frmAddServiciosMarMod.Show 1
    
    ActualizaTotales
End Sub

Private Sub cmdBorrarOtrosServicios_Click()
    Dim i As Integer
    If Me.lvwOtrosServicios.ListItems.Count = 0 Then
        Exit Sub
    End If

    If MsgBox("¿ Desea eliminar los servicios seleccionados ?", vbInformation + vbYesNo + vbDefaultButton2, "Advertencia") = vbYes Then
        For i = 1 To Me.lvwOtrosServicios.ListItems.Count
            If Me.lvwOtrosServicios.ListItems(i).Selected Then
                Me.lvwOtrosServicios.ListItems.Remove i
            End If
            If i > Me.lvwOtrosServicios.ListItems.Count Then
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cmdBorrarServiciosMecanica_Click()
    Dim i As Integer
    If Me.lvwServiciosMecanica.ListItems.Count = 0 Then
        Exit Sub
    End If

    If MsgBox("¿ Desea eliminar los servicios seleccionados ?", vbInformation + vbYesNo + vbDefaultButton2, "Advertencia") = vbYes Then
        For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
            If Me.lvwServiciosMecanica.ListItems(i).Selected Then
                Me.lvwServiciosMecanica.ListItems.Remove i
            End If
            If i > Me.lvwServiciosMecanica.ListItems.Count Then
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cmdCrearOtrosServicios_Click()
    Dim strServicios As String
    Dim strHoras As String
    Dim Item As ListItem
    
    strServicios = ""
    Do
        strServicios = InputBox("Ingrese el Servicio", "Otros Servicios", strServicios)
        If strServicios = "" Then
            Exit Sub
        End If
        Exit Do
    Loop
    
    Do
        strHoras = "0"
        strHoras = InputBox("Ingrese la cantidad de Horas", "Otros Servicios", strHoras)
        If strHoras = "" Then
            Exit Sub
        End If
        strHoras = Replace(strHoras, ",", "")
        If IsNumeric(strHoras) Then
'            If CDbl(strHoras) > 0 Then
            'kjcv 12.11.13
            If CDbl(strHoras) > 0 Or CDbl(strHoras) = 0 Then
                Exit Do
'            Else
'                MsgBox "La cantidad de horas debe ser mayor a Cero...", vbInformation, "Advertencia"
            End If
        End If
    Loop
    
    '//Crea Otro Servicio...
    Set Item = Me.lvwOtrosServicios.ListItems.Add(, , UCase(strServicios))
    Item.SubItems(1) = Format(CDbl(strHoras), "#,##0.0")
    
    ActualizaTotales
End Sub

Private Sub cmdSinPatente_Click()
If Me.Tag = "Crear" Then
    LimpiaCampos
    ValoresporDefecto
    frmIngresaDatosReservaHora.Show vbModal
End If
End Sub
Private Sub Form_Load()
'Dim Sql As String
Dim sqlRecep As String
'Dim gstrIdEmpleado As String
'Dim gstrIdMecanico As String

     
        
    
    mblnSW = True
    Me.lblPat.Caption = gstrNombrePatente
    Me.cmdSinPatente.Caption = "Sin " & gstrNombrePatente
    
    gstrIdMecanico = CodigoMecanico(gstrIdEmpleado, gstrIdEmpresa, gstrIdSucursal)
  FillRecepcionista dtcRecepcionista, datRecepcionista
    'kjcv 03.07.14 se agrega recepcionista por default el que ingresa
    Me.dtcRecepcionista.BoundText = gstrIdMecanico
    
   
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gintProcedencia = 0
End Sub



Private Sub lblIdCliente_Change()
    If DatosCliente(lblIdCliente) Then DoEvents
End Sub
Private Sub lblNroRecepcion_DblClick()
    gstrBusca = InputBox("Ingrese El Numero de RESERVA Deseado :", "Ir a....", CStr(Val(lblNroRecepcion)))
    gstrBusca = Format(gstrBusca, "00000")
    If gstrBusca <> "" Then
        mstrWhere = " WHERE Tllr_RESERVAHORA.ID_RESERVA=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        mstrOrderBy = " ORDER BY Tllr_RESERVAHORA.Id_RESERVA"
        gstrSql = letSql(mstrWhere, mstrOrderBy)
        If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost AdoPrincipal
    End If
    Me.SetFocus
End Sub




Private Sub pckFechaEntrega_Click()
    frmConsultaReservaHoras.Show vbModal
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            AgregarRegistro
        Case "Grabar"
            GrabarRegistro
        Case "Cancelar"
            CancelarAgregaRegistro
        Case "Borrar"
            BorrarRegistro
        Case "Buscar"
            BuscarRegistro
        Case "Imprimir"
            PrintReserva
        Case "Primero"
            PrimerRegistro
        Case "Anterior"
            RegistroAnterior
        Case "Siguiente"
            RegistroSiguiente
        Case "Ultimo"
            UltimoRegistro
        Case "Confirmar"
            GenerarOTReserva
        Case "Activar"
            ActivarReserva
        Case "Renovar"
            Renovar
        Case "Cerrar"
            CerrarSalir
    End Select
Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()

    If mblnSW Then
        mblnSW = False
        If Not Atributos("Glbl", "Tllr_20_0070", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If '/////////ojo
        
        FillRecepcionista dtcRecepcionista, datRecepcionista
'        FillTime gintHoraInicio, 20, cboHora
        
        '//Crear registro por defecto...
        If gapAccion = apcrear Then
           AgregarRegistro
           lblNroRecepcion = gstrBusca
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
        '//Editar registro por defecto...
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrWhere = " WHERE Tllr_ReservaHora.ID_Reserva='" & gstrBusca & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"""
                mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva"
                gstrSql = letSql(mstrWhere, mstrOrderBy)
                If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            Me.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If gapAccion = apninguno Then
           Renovar
        End If
        
    End If
    gapAccion = apninguno
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            CancelarAgregaRegistro
        Case 14 And tlbBarraHerramientas.Buttons.Item("Crear").Enabled
            KeyAscii = 0
            AgregarRegistro
        Case 7 And tlbBarraHerramientas.Buttons.Item("Grabar").Enabled
            KeyAscii = 0
            GrabarRegistro
        Case 4 And tlbBarraHerramientas.Buttons.Item("Borrar").Enabled
            KeyAscii = 0
            BorrarRegistro
        Case 2 And tlbBarraHerramientas.Buttons.Item("Buscar").Enabled
            KeyAscii = 0
            BuscarRegistro
        Case 9 And tlbBarraHerramientas.Buttons.Item("Imprimir").Enabled
            KeyAscii = 0
            PrintReserva
        Case 16 And tlbBarraHerramientas.Buttons.Item("Primero").Enabled
            KeyAscii = 0
            PrimerRegistro
        Case 1 And tlbBarraHerramientas.Buttons.Item("Anterior").Enabled
            KeyAscii = 0
            RegistroAnterior
        Case 19 And tlbBarraHerramientas.Buttons.Item("Siguiente").Enabled
            KeyAscii = 0
            RegistroSiguiente
        Case 21 And tlbBarraHerramientas.Buttons.Item("Ultimo").Enabled
            KeyAscii = 0
            UltimoRegistro
        Case 18 And tlbBarraHerramientas.Buttons.Item("Renovar").Enabled
            KeyAscii = 0
            Renovar
        Case 3 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub
Public Sub AgregarRegistro()
    Me.Tag = "Crear"
    lblEstadoOTValor = ""
    lblNroRecepcion = ""
'    If frmRecordatorioServicio.mblSWRecordatorio = False Then
    DesactivaBotones
'    Else
'    ActivaBotones
'    End If
    
    LimpiaCampos
    ValoresporDefecto
    Bloqueo "V"
    If fmePat.Enabled = True Then
        txtPatente.SetFocus
    End If
    cmdSinPatente.Enabled = True
    Me.Tag = "Crear"
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones                                                                       'AND Tllr_OT.ID_OT = Tllr_OT.ID_OT >'" & Trim(lblNroRecepcion) & "'
    mstrWhere = " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva DESC"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            mstrWhere = " WHERE Tllr_ReservaHora.ID_Reserva < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva"
            gstrSql = letSql(mstrWhere, mstrOrderBy)
            If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub GrabarRegistro()
    If Not validacion() Then
        Exit Sub
    End If
    
    If Me.Tag = "Crear" Then
        lblNroRecepcion = TraeCorrelativoReserva(gstrIdEmpresa, gstrIdSucursal)
        lblEstadoOTValor = "VIGENTE"
        If Me.optSinPatente.Value = False Then
            mstrSQL = "INSERT INTO Tllr_ReservaHora "
            mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal, "
            mstrSQL = mstrSQL & " Id_Reserva , Patente, RealizadoPor,"
            mstrSQL = mstrSQL & " Estado,Fecha_Emision, "
            mstrSQL = mstrSQL & " Fecha_Reserva, Hora_Reserva, Seccion_OT,"
'            mstrSQL = mstrSQL & " Reparacion, Total_Mecanica, Total_Otros, Total_Repuestos, Total_Reserva,Recepcionista )"
'kjcv 11.09.14
            mstrSQL = mstrSQL & " Reparacion, Total_Mecanica, Total_Otros, Total_Repuestos, Total_Reserva,Taxi_destino,Recepcionista )"
            mstrSQL = mstrSQL & " VALUES ("
'            mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & txtSucursal & "',"
            'kjcv 30.10.13
            mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
            mstrSQL = mstrSQL & " '" & lblNroRecepcion & "',"
            mstrSQL = mstrSQL & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
            mstrSQL = mstrSQL & " 'V','" & CDate(pckFechaAtencion.Value) & "', "
'            mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , '" & IIf(Me.optMecanica.Value = True, "M", "C") & "',"
            'kjcv 13.11.13
            mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & txtHora.Text & "' , '" & IIf(Me.optMecanica.Value = True, "M", "C") & "',"
            mstrSQL = mstrSQL & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/R") & "',"
'            mstrSQL = mstrSQL & " 0,0,0,0,'" & txtRecepcionista & "')"
'kjcv 11.09.14
            mstrSQL = mstrSQL & " 0,0,0,0,'" & IIf(Trim(txtTaxiDestino.Text) <> "", UCase(Trim(txtTaxiDestino.Text)), ".") & "','" & txtRecepcionista & "')"
        Else
            mstrSQL = "INSERT INTO Tllr_ReservaHora "
            mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal, "
            mstrSQL = mstrSQL & " Id_Reserva , Patente, RealizadoPor,"
            mstrSQL = mstrSQL & " Estado,Fecha_Emision, "
            mstrSQL = mstrSQL & " Fecha_Reserva, Hora_Reserva, Seccion_OT,"
'            mstrSQL = mstrSQL & " Reparacion, Recepcionista, SinPatente, Nombre, Vehiculo, Telefono )"
            'kjcv 11.09.14
            mstrSQL = mstrSQL & " Reparacion,Taxi_destino ,Recepcionista, SinPatente, Nombre, Vehiculo, Telefono )"
            mstrSQL = mstrSQL & " VALUES ("
'            mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & txtSucursal & "',"
            'kjcv 30.10.13
            mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
            mstrSQL = mstrSQL & " '" & lblNroRecepcion & "',"
            mstrSQL = mstrSQL & " '" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "',"
            mstrSQL = mstrSQL & " 'V','" & CDate(pckFechaAtencion.Value) & "', "
'            mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , '" & IIf(Me.optMecanica.Value = True, "M", "C") & "',"
            'kjcv 13.11.13
            mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & txtHora.Text & "' , '" & IIf(Me.optMecanica.Value = True, "M", "C") & "',"
            mstrSQL = mstrSQL & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/R") & "',"
            'kjcv 11.09.14
            mstrSQL = mstrSQL & " '" & IIf(Trim(txtTaxiDestino.Text) <> "", UCase(Trim(txtTaxiDestino.Text)), ".") & "',"
            mstrSQL = mstrSQL & " '" & txtRecepcionista & "',"
            mstrSQL = mstrSQL & " 'S',"
            mstrSQL = mstrSQL & " '" & lblCliente & "',"
            mstrSQL = mstrSQL & " '" & lblModelo & "',"
            mstrSQL = mstrSQL & " '" & lblFono & "')"
        End If
    Else
        If Me.optSinPatente.Value = False Then
            mstrSQL = "UPDATE Tllr_ReservaHora "
            mstrSQL = mstrSQL & " SET Patente='" & txtPatente.Text & "', "
            mstrSQL = mstrSQL & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
            mstrSQL = mstrSQL & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
            mstrSQL = mstrSQL & " Fecha_Reserva='" & CDate(pckFechaEntrega) & "', "
'            mstrSQL = mstrSQL & " Hora_Reserva='" & cboHora.Text & "', "
            'kjcv 13.11.13
            mstrSQL = mstrSQL & " Hora_Reserva='" & txtHora.Text & "', "
            mstrSQL = mstrSQL & " Seccion_OT='" & IIf(Me.optMecanica.Value = True, "M", "C") & "', "
            mstrSQL = mstrSQL & " Reparacion='" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), ".") & "',"
            'kjcv 11.09.14
            mstrSQL = mstrSQL & " Taxi_destino='" & IIf(Trim(txtTaxiDestino.Text) <> "", UCase(Trim(txtTaxiDestino.Text)), ".") & "',"
            mstrSQL = mstrSQL & " Total_Mecanica= 0,"
            mstrSQL = mstrSQL & " Total_Otros= 0,"
            mstrSQL = mstrSQL & " Total_Repuestos= 0,"
            mstrSQL = mstrSQL & " Total_Reserva= 0, "
            mstrSQL = mstrSQL & " Recepcionista='" & txtRecepcionista & "',"
            mstrSQL = mstrSQL & " SinPatente='" & IIf(Me.optSinPatente.Value = True, "S", "N") & "',"
            mstrSQL = mstrSQL & " Id_Sucursal='" & txtSucursal & "'"
            mstrSQL = mstrSQL & " WHERE Id_Empresa ='" & gstrIdEmpresa & "' And Id_Reserva ='" & Trim(Trim(lblNroRecepcion)) & "' "
        Else
            mstrSQL = "UPDATE Tllr_ReservaHora "
            mstrSQL = mstrSQL & " SET Patente='" & txtPatente.Text & "', "
            mstrSQL = mstrSQL & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
            mstrSQL = mstrSQL & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
            mstrSQL = mstrSQL & " Fecha_Reserva='" & CDate(pckFechaEntrega) & "', "
'            mstrSQL = mstrSQL & " Hora_Reserva='" & cboHora.Text & "', "
            'kjcv 13.11.13
            mstrSQL = mstrSQL & " Hora_Reserva='" & txtHora.Text & "', "
            mstrSQL = mstrSQL & " Seccion_OT='" & IIf(Me.optMecanica.Value = True, "M", "C") & "', "
            mstrSQL = mstrSQL & " Reparacion='" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), ".") & "',"
            'kjcv 11.09.14
            mstrSQL = mstrSQL & " Taxi_destino='" & IIf(Trim(txtTaxiDestino.Text) <> "", UCase(Trim(txtTaxiDestino.Text)), ".") & "',"
            mstrSQL = mstrSQL & " Recepcionista='" & txtRecepcionista & "',"
            mstrSQL = mstrSQL & " Id_Sucursal='" & txtSucursal & "',"
            mstrSQL = mstrSQL & " SinPatente='S',"
            mstrSQL = mstrSQL & " Nombre='" & lblCliente & "',"
            mstrSQL = mstrSQL & " Vehiculo='" & lblModelo & "',"
            mstrSQL = mstrSQL & " Telefono='" & lblFono & "'"
            mstrSQL = mstrSQL & " WHERE Id_Empresa ='" & gstrIdEmpresa & "' And Id_Reserva ='" & Trim(Trim(lblNroRecepcion)) & "' "
        End If
    End If
    
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
        'cmdSinPatente.Enabled = True
        Me.Tag = ""
    End If '//////////////
    
    '//LREYES graba servicios de mecanica y otros servicios...
    Dim i As Integer
    Dim strSql As String
    strSql = "delete from Tllr_ReservaHora_Mecanica where id_empresa = '" & gstrIdEmpresa & "' and id_sucursal = '" & gstrIdSucursal & "' and id_reserva = '" & Trim(lblNroRecepcion) & "'"
    If Conexion.SendHost(strSql, , adOpenKeyset, adLockOptimistic, 10) = apOk Then
        For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
            strSql = "insert into Tllr_ReservaHora_Mecanica (Id_Empresa, Id_Sucursal, Id_Reserva, Id_Item, Id_Servicio, Horas)"
            strSql = strSql & "values('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', '" & Trim(lblNroRecepcion) & "', " & i & ", '" & Me.lvwServiciosMecanica.ListItems(i) & "', " & CDbl(Me.lvwServiciosMecanica.ListItems(i).SubItems(2)) & ")"
            If Conexion.SendHost(strSql, , , , 10) = apOk Then
                
            End If
        Next
    End If

    strSql = "delete from Tllr_ReservaHora_Otros_Servicios where id_empresa = '" & gstrIdEmpresa & "' and id_sucursal = '" & gstrIdSucursal & "' and id_reserva = '" & Trim(lblNroRecepcion) & "'"
    If Conexion.SendHost(strSql, , adOpenKeyset, adLockOptimistic, 10) = apOk Then
        For i = 1 To Me.lvwOtrosServicios.ListItems.Count
            strSql = "insert into Tllr_ReservaHora_Otros_Servicios (Id_Empresa, Id_Sucursal, Id_Reserva, Id_Item, Servicio, Horas)"
            strSql = strSql & "values('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', '" & Trim(lblNroRecepcion) & "', " & i & ", '" & Me.lvwOtrosServicios.ListItems(i) & "', " & CDbl(Me.lvwOtrosServicios.ListItems(i).SubItems(1)) & ")"
            If Conexion.SendHost(strSql, , , , 10) = apOk Then
                
            End If
        Next
    End If

End Sub
Private Sub BorrarRegistro()
Dim mstrMotivoAnula As String

    Screen.MousePointer = vbDefault
    If MsgBox("¿ Esta Seguro de Anular esta reserva ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        '////////////////////////////////ELIMINAR SERVICIOS DE MECANICA///////////////////////////////////
'        mstrSql = "DELETE FROM Tllr_Mecanica_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
'        Conexion.SendHost mstrSql, , , , gcTiempoEspera
'        '////////////////////////////////ELIMINAR SERVICIOS DE CARRPCERIA///////////////////////////////////
'        mstrSql = "DELETE FROM Tllr_Carroceria_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
'        Conexion.SendHost mstrSql, , , , gcTiempoEspera
'        '////////////////////////////////////ELIMINAR INENTARIO///////////////////////////////
'        mstrSql = "DELETE FROM Tllr_Inventario_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT='" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
'        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        '//////////////////////////////////////ENCABEZADO/////////////////////////////
        
        mstrMotivoAnula = InputBox("Ingrese el Motivo de Anulación", "Por que Anula...")
        gstrSql = "UPDATE TLLR_ReservaHora SET ESTADO = 'N' ,"
        gstrSql = gstrSql & "Fecha_Anulacion = '" & CDate(pckFechaAtencion.Value) & "', "
        gstrSql = gstrSql & "Quien_Anula = '" & gstrIdUsuario & "', "
        gstrSql = gstrSql & "MotivoAnula = '" & mstrMotivoAnula & "' "
        gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_ReservaHora.Id_Reserva = '" & lblNroRecepcion & "' "
        If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
            lblEstadoOTValor = "NULA"
            tlbBarraHerramientas.Buttons.Item(2).Enabled = False    'guardar
            tlbBarraHerramientas.Buttons.Item(18).Enabled = False    'Confirmar
            tlbBarraHerramientas.Buttons.Item(17).Enabled = False    'Anular
            tlbBarraHerramientas.Buttons.Item(19).Enabled = True    'Activar
            Bloqueo "N"
        End If
        MsgBox "La Reserva Nº " & lblNroRecepcion & " Fue Anulada"
    End If
End Sub
Private Sub BuscarRegistro()
Screen.MousePointer = 1
frmBuscaReserva.Show vbModal
Screen.MousePointer = 1
If gstrBusca <> "" Then
    mstrWhere = " WHERE Tllr_ReservaHora.ID_Reserva=  '" & gstrBusca & "' And Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End If
Me.SetFocus

End Sub
Private Sub PrimerRegistro()
    mstrWhere = " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'" & " And Estado='V'"
    mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroAnterior()
    mstrWhere = " WHERE Tllr_ReservaHora.Id_Reserva < '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva DESC"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroSiguiente()
    mstrWhere = " WHERE Tllr_ReservaHora.Id_Reserva > '" & Trim(lblNroRecepcion) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva "
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub UltimoRegistro()
    mstrWhere = " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva DESC"
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub Renovar()
    mstrWhere = " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'" & " And Estado= 'V'"
    mstrOrderBy = " ORDER BY Tllr_ReservaHora.Id_Reserva "
    gstrSql = letSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        VerificaTablaVacia
        ActivaBotones
        If Not mblnTablaVacia Then
            PrimerRegistro
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
    txtPatente.Enabled = False
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
        .Item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoEditar, True, False))
        .Item("Cancelar").Enabled = False
        .Item("Borrar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoBorrar, True, False))
        .Item("Buscar").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Imprimir").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoImprimir, True, False))
        .Item("Primero").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Anterior").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Siguiente").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Ultimo").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Renovar").Enabled = True
        .Item("Cerrar").Enabled = True
        .Item("Confirmar").Enabled = IIf(Me.lblEstadoOTValor = "CONFIRMADA", False, True)
    End With
End Sub
Private Sub DesactivaBotones()
    txtPatente.Enabled = True
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = False
        If swActivateRecorda Then
        .Item("Grabar").Enabled = True
        Else
        .Item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
        End If
        
        .Item("Cancelar").Enabled = True
        .Item("Borrar").Enabled = False
        .Item("Buscar").Enabled = False
        .Item("Imprimir").Enabled = False
        .Item("Primero").Enabled = False
        .Item("Anterior").Enabled = False
        .Item("Siguiente").Enabled = False
        .Item("Ultimo").Enabled = False
        .Item("Renovar").Enabled = False
        .Item("Cerrar").Enabled = True
        .Item("Confirmar").Enabled = False
    End With
End Sub
Private Sub VerificaTablaVacia()
    If (Not AdoPrincipal.BOF And Not AdoPrincipal.EOF) And AdoPrincipal.RecordCount > 0 Then
        mblnTablaVacia = False
    Else
        mblnTablaVacia = True
        LimpiaCampos
        MsgBox "La tabla no contiene registros...", vbInformation, "Advertencia"
        Bloqueo "V"
    End If
End Sub
Private Sub LeerCampos()
If mblnTablaVacia Then
    LimpiaCampos
    Exit Sub
End If
With AdoPrincipal

    If !SinPatente <> "S" Or IsNull(!SinPatente) Then
        txtPatente = ValorNulo(!Patente)
        If ValorNulo(!Patente) <> "" Then DatosVehiculo !Patente
        txtPatente.Enabled = False
        optSinPatente.Value = False
        cmdSinPatente.Enabled = False
    Else
        LimpiaCampos
        optSinPatente = True
        lblCliente = ValorNulo(!Nombre)
        lblModelo = ValorNulo(!Vehiculo)
        lblFono = ValorNulo(!Telefono)
        fmePat.Enabled = True
        txtPatente.Enabled = True
        txtPatente.SetFocus
        cmdSinPatente.Enabled = True
    End If

    lblNroRecepcion.Text = !Id_Reserva
    lblNumeroOt = IIf(IsNull(!Id_OT), "Sin OT", !Id_OT)
    If !Seccion_OT = "C" Then
        optCarroceria.Value = True
    Else
        optMecanica.Value = True
    End If
    dtcRecepcionista.BoundText = !RealizadoPor
    pckFechaAtencion.Value = !Fecha_Emision
    pckFechaEntrega.Value = !Fecha_Reserva
'    cboHora.Text = ValorNulo(!Hora_Reserva)
    'kjcv 13.11.13
    txtHora.Text = ValorNulo(!Hora_Reserva)
    
    txtComentario = !Reparacion
    txtRecepcionista = IIf(IsNull(!Recepcionista), gstrMecanicoDefectoSecMec, !Recepcionista)
    txtTaxiDestino = IIf(IsNull(!Taxi_destino), "", !Taxi_destino)
    'kjcv 14.11.13
'    gstrNombreRecepcionista = NombreRecepcionista(!Recepcionista)
'    gstrNombreRecepLlamado = NombreRecepcionista(!RealizadoPor)
    txtSucursal = IIf(IsNull(!Id_Sucursal), gstrIdSucursal, !Id_Sucursal)
    
    If Not IsNull(!estado) Then
        lblEstadoOTValor.Caption = IIf(!estado = "V", "VIGENTE", IIf(!estado = "C", "CONFIRMADA", IIf(!estado = "N", "NULA", IIf(!estado = "E", "CANCELADA", IIf(!estado = "R", "RECEPCIONADA", "")))))
        tlbBarraHerramientas.Buttons.Item(2).Enabled = IIf(!estado = "V", True, IIf(!estado = "L", False, IIf(!estado = "N", True, IIf(!estado = "F" Or !estado = "B", True, False))))
        tlbBarraHerramientas.Buttons.Item(18).Enabled = IIf(!estado = "V", True, IIf(!estado = "C", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, False))))    'Confirmar
        tlbBarraHerramientas.Buttons.Item(17).Enabled = IIf(!estado = "V", True, IIf(!estado = "C", False, IIf(!estado = "N", False, IIf(!estado = "F" Or !estado = "B", False, False))))    'Anular
        tlbBarraHerramientas.Buttons.Item(19).Enabled = IIf(!estado = "V", False, IIf(!estado = "C", False, IIf(!estado = "N", True, IIf(!estado = "F" Or !estado = "B", False, False))))    'Activar
'        tlbBarraHerramientas.Buttons.Item(15).Enabled = IIf(!Estado = "V", True, IIf(!Estado = "L", False, IIf(!Estado = "N", False, IIf(!Estado = "F" Or !Estado = "B", True, False))))    'LIQUIDAR
        Bloqueo !estado
    End If
    
    
    '//...Leer detalles
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    Dim Item As ListItem
    
    Me.lvwServiciosMecanica.ListItems.Clear
    strSql = "SELECT     isnull(dbo.Tllr_ReservaHora_Mecanica.Id_Servicio,'') as Id_Servicio, isnull(dbo.Tllr_Servicio.Descripcion,'') as Descripcion, isnull(dbo.Tllr_ReservaHora_Mecanica.Horas,0) as Horas "
    strSql = strSql & "FROM         dbo.Tllr_ReservaHora_Mecanica INNER JOIN dbo.Tllr_Servicio ON dbo.Tllr_ReservaHora_Mecanica.Id_Servicio = dbo.Tllr_Servicio.Id_Servicio "
    strSql = strSql & "where id_empresa = '" & gstrIdEmpresa & "' and id_sucursal = '" & gstrIdSucursal & "' and id_reserva = '" & Trim(Me.lblNroRecepcion) & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            While Not adoTemp.EOF
                Set Item = Me.lvwServiciosMecanica.ListItems.Add(, , adoTemp!Id_servicio)
                Item.SubItems(1) = adoTemp!Descripcion
                Item.SubItems(2) = adoTemp!Horas
                adoTemp.MoveNext
            Wend
        End If
    End If
    Conexion.CloseHost adoTemp
    
    Me.lvwOtrosServicios.ListItems.Clear
    strSql = "select isnull(Servicio,'') as Servicio, isnull(horas,0) as horas from Tllr_ReservaHora_Otros_Servicios "
    strSql = strSql & "where id_empresa = '" & gstrIdEmpresa & "' and id_sucursal = '" & gstrIdSucursal & "' and id_reserva = '" & Trim(Me.lblNroRecepcion) & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            While Not adoTemp.EOF
                Set Item = Me.lvwOtrosServicios.ListItems.Add(, , adoTemp!servicio)
                Item.SubItems(1) = adoTemp!Horas
                adoTemp.MoveNext
            Wend
        End If
    End If
    Conexion.CloseHost adoTemp

    ActualizaTotales
End With
End Sub
Private Sub LimpiaCampos()
    txtPatente.Text = ""
    lblMarca = ""
    lblModelo = ""
    txtAño = ""
    txtKilAct = ""
    'lblChasis = ""
    txtChasis.Text = ""
    lblMotor = ""
    'lblVin = ""
    txtVin.Text = ""
    lblCliente = ""
    lblIdCliente = ""
    lblColorE = ""
    lblFono = ""
    dtcRecepcionista.Text = ""
    txtTaxiDestino.Text = ""
    txtComentario = ""
'    cboHora = ""
    'kjcv 13.11.13
    txtHora = ""
    optCarroceria.Value = False
    optMecanica.Value = False '7622225
    lblNumeroOt = ""
    optSinPatente.Value = False
    Me.lblHorasMecanica = "0.0"
    Me.lblHorasOtrosServicios = "0.0"
    Me.lblTotalHoras(0) = "0.0"
    Me.lblTotalHoras(1) = "0.0"

    Me.lvwServiciosMecanica.ListItems.Clear
    Me.lvwOtrosServicios.ListItems.Clear
End Sub
Private Sub ValoresporDefecto()
  Me.pckFechaAtencion = Date
  Me.pckFechaEntrega = Date
  txtRecepcionista = gstrMecanicoDefectoSecMec
  Me.dtcRecepcionista.BoundText = gstrIdMecanico
End Sub
Private Function validacion() As Boolean
    validacion = True
    If txtPatente = "" Then
        If Me.optSinPatente.Value = False Then
            MsgBox "La " & gstrNombrePatente & " debe contener un valor...", vbInformation, "Advertencia"
            txtPatente.SetFocus
            validacion = False
            Exit Function
        End If
    End If
            
    optMecanica.Value = True
'    If optMecanica.Value = False Then
'        If optCarroceria.Value = False Then
'            MsgBox "Seleccione un Valor de Mecanica o Carroceria ", vbExclamation, "Advertencia"
'            Validacion = False
'            Exit Function
'        End If
'    End If
    
    If Me.dtcRecepcionista.Text = "" Then
        MsgBox "El Recepcionista debe contener un valor...", vbInformation, "Advertencia"
        Me.dtcRecepcionista.SetFocus
        validacion = False
        Exit Function
    End If
       
    If Me.pckFechaEntrega = "" Then
        MsgBox "La Fecha de Reserva Debe Contener un Valor...", vbInformation, "Advertencia"
        Me.pckFechaEntrega.SetFocus
        validacion = False
        Exit Function
    End If
           
    If CDate(pckFechaEntrega) < Date Then
        MsgBox "La Fecha de Reserva No Puede ser Menor a la Actual", vbInformation, "Advertencia"
        Me.pckFechaEntrega.SetFocus
        validacion = False
        Exit Function
    End If
           
'    If Me.cboHora = "" Then
'        MsgBox "Debe Seleccionar una Hora De Reserva ...", vbInformation, "Advertencia"
'        Me.cboHora.SetFocus
'        Validacion = False
'        Exit Function
'    End If
'kjcv 13.11.13
    If Me.txtHora = "" Then
        MsgBox "Debe Seleccionar una Hora De Reserva ...", vbInformation, "Advertencia"
'        Me.txtHora.SetFocus
        validacion = False
        Exit Function
    End If
           
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As New ADODB.Recordset
        mstrSQL = "select ID_Reserva from TLLR_ReservaHora where ID_Reserva ='" & lblNroRecepcion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción "
                validacion = False
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
       
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMantenedorVehiculosPropios = Nothing
    gstrBusca = txtPatente.Text
End Sub
Private Sub RevizaAtributos()
'    mblnAccesoCrear = rsUsuarios!OPC_AUXILIAR_CREAR
'    mblnAccesoEditar = rsUsuarios!OPC_AUXILIAR_EDITAR
'    mblnAccesoBorrar = rsUsuarios!OPC_AUXILIAR_BORRAR
'    mblnAccesoImprimir = rsUsuarios!OPC_AUXILIAR_IMPRIMIR

    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True

End Sub

Private Sub tlbLlamadoTelefono_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Llamar"
        LlamarTelefono "Reserva De Horas", Me.lblCliente, Me.lblFono, Me.txtComentario
End Select
End Sub

Private Sub tlbPatente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Me.Tag = "Crear" Then
    Select Case Button.Key
    Case "Nuevo"
        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apcrear)
        DatosVehiculo txtPatente
    Case "Buscar"
        gstrProcedencia = "ReservaHora"
        frmBuscaVehiculo.Show vbModal
        Me.txtPatente = gstrBusca
        txtPatente_KeyDown 13, 0
    End Select
Else
    Select Case Button.Key
    Case "Nuevo"
        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apeditar)
        DatosVehiculo txtPatente
    End Select
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
'kjcv 13.11.13 Valida Comilla Simple
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtPatente_Click()
If optSinPatente.Value = True Then
    LimpiaCampos
    ValoresporDefecto
    optSinPatente = False
End If
End Sub

Private Sub txtPatente_KeyDown(KeyCode As Integer, Shift As Integer)
'If Me.Tag = "Crear" Then
Dim str1 As String
Dim str2 As String
    If KeyCode = 13 Then
        If txtPatente <> "" Then
            If Len(txtPatente) = 6 And lblPat.Caption = gstrNombrePatente Then
                If ConsultaVehiculo(txtPatente) = True Then
                'kjcv 30.10.15
                   If ConsultaPatente(txtPatente) = True Then
                        MsgBox "No hay Cupo en el Taller...", vbCritical, "Elisa"
                        Call DatosVehiculo(txtPatente)
                    Else
                        Call DatosVehiculo(txtPatente)
                    End If
                
'                    Call DatosVehiculo(txtPatente)
                    optSinPatente.Value = False
                Else
                    gstrProcedencia = "ReservaHora"
                    gapAccion = apcrear
                    frmMantenedorVehiculoCliente.Show vbModal
                End If
            
            Else
                MsgBox LoadResString(326)
            End If
            
        Else
            MsgBox LoadResString(327)
        End If
    End If
    
'End If
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'If gstrValidaPatente = "S" Then
'    KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'End If
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub

Sub DatosVehiculo(strPatente As String)
If strPatente <> "" Then
    mstrSQL = "SELECT Tllr_Vehiculo_Cliente.Patente,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Marca AS IDMARCA,"
    mstrSQL = mstrSQL & " Glbl_Marca.Descripcion AS MARCA,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Modelo AS IDMODELO,"
    mstrSQL = mstrSQL & " Glbl_Modelo.Descripcion AS MODELO,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Año,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Color_Exterior AS IDCOLOR,"
    mstrSQL = mstrSQL & " Glbl_Color_Exterior.Descripcion AS COLOR,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Kilometros_Actuales AS KILACT,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Nro_Motor AS MOTOR,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Nro_Chasis AS CHASIS,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.VIN AS VIN,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor AS IDCLI,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Fecha_Venta AS FECVTA,"
    mstrSQL = mstrSQL & " Tllr_Vehiculo_Cliente.Concesionario AS CONCES"
    mstrSQL = mstrSQL & " FROM Glbl_Cliente_Proveedor RIGHT OUTER JOIN Glbl_Color_Exterior RIGHT OUTER JOIN Tllr_Vehiculo_Cliente ON Glbl_Color_Exterior.Id_Color_Exterior = Tllr_Vehiculo_Cliente.Id_Color_Exterior LEFT OUTER JOIN Glbl_Modelo LEFT OUTER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca ON Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor"
    mstrSQL = mstrSQL & " WHERE Tllr_Vehiculo_Cliente.Patente='" & strPatente & "'"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            With AdoPrincipal
                lblMarca = ValorNulo(!Marca)
                lblIdMarca = ValorNulo(!IdMarca)
                lblModelo = ValorNulo(!Modelo)
                lblIdModelo = ValorNulo(!IdModelo)
                'lblChasis = ValorNulo(!chasis)
                txtChasis.Text = ValorNulo(!chasis)
                lblMotor = ValorNulo(!motor)
                'lblVin = ValorNulo(!VIN)
                txtVin.Text = ValorNulo(!VIN)
                txtAño = ValorNulo(!Año)
                lblColorE = ValorNulo(!Color)
                'lblCliente = ValorNulo(!idCLI)
                pckFecVta.Value = IIf(Not IsNull(!FECVTA), !FECVTA, Now)
                txtKilAct = IIf(Not IsNull(!kilact), !kilact, "0")
                lblIdCliente = ValorNulo(!idCLI)
                KilometrajeEntrada = txtKilAct 'Variable de ileiva 07/02/2001
            End With
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End If
End Sub

Function letSql(strWhere As String, strOrder As String) As String
mstrSQL = "SELECT top 1 * From Tllr_ReservaHora"
letSql = mstrSQL & " " & strWhere & " " & strOrder

End Function

Sub PrintReserva()

Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
'kjcv 14.11.13 se agrega datos de recepcionista
gstrNombreRecepcionista = NombreRecepcionista(txtRecepcionista)
gstrNombreRecepLlamado = NombreRecepcionista(dtcRecepcionista.BoundText)

On Error GoTo Solucion
    
    'Devuelve la ruta del directorio Windows
    Dim rc As Long
    Dim WinPath As String
    WinPath = Space$(300)
    rc = GetWindowsDirectory(WinPath, 300)
    GcamBaseTem = Trim$(WinPath)
    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
    '---------------------------------------
    
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
'    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
'    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
'    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Reparacion Memo)"
    Dbsnueva.Execute "CREATE TABLE T_PARAMECANICA (IdServicio text,Descripcion text,Horas text)"
    Dbsnueva.Execute "CREATE TABLE T_PARASERVICIO (Servicio text,Horas text)"
    
'    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
'    Tabla.AddNew
'    Tabla!Reparacion = Me.txtComentario
'    Tabla.Update
'    Tabla.Close

    'kjcv 13.11.13
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAMECANICA")
    For i = 1 To lvwServiciosMecanica.ListItems.Count
        Set lvwServiciosMecanica.SelectedItem = lvwServiciosMecanica.ListItems(i)
        Tabla.AddNew
        Tabla!idServicio = lvwServiciosMecanica.ListItems(i)
        Tabla!Descripcion = IIf(lvwServiciosMecanica.SelectedItem.SubItems(1) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(1))
        Tabla!Horas = IIf(lvwServiciosMecanica.SelectedItem.SubItems(2) = "", " ", lvwServiciosMecanica.SelectedItem.SubItems(2))
        Tabla.Update
    Next i
    Tabla.Close
    
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARASERVICIO")
    For i = 1 To lvwOtrosServicios.ListItems.Count
        Set lvwOtrosServicios.SelectedItem = lvwOtrosServicios.ListItems(i)
        Tabla.AddNew
        Tabla!servicio = lvwOtrosServicios.ListItems(i)
        Tabla!Horas = IIf(lvwOtrosServicios.SelectedItem.SubItems(1) = "", " ", lvwOtrosServicios.SelectedItem.SubItems(1))
        Tabla.Update
    Next i
    Tabla.Close
    Dbsnueva.Close
   With rptReserva
        .ReportFileName = gstrPathReporte & "\RESERVAHORAS.rpt"
'        .ReportFileName = gstrPathReporte & "\ReservaHorasPrueba.rpt"
        .WindowTitle = "Informe Reserva de Horas"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='RESERVA DE HORA PARA SERVICIO   N°" & Me.lblNroRecepcion & "'"
        .Formulas(2) = "Empresa='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        
        .Formulas(5) = "RazonSocial='" & Me.lblCliente & "'"
'        .Formulas(6) = "RutCliente='" & FormatoRut(Me.lblIdCliente) & "'"
        'kjcv 13.11.13
        .Formulas(6) = "RutCliente='" & (Me.lblIdCliente) & "'"
        .Formulas(7) = "DireccionCliente='" & Me.txtDir & "'"
        .Formulas(8) = "Comuna='" & Me.txtComuna & "'"
        .Formulas(9) = "Telefono='" & Me.lblFono & "'"

        .Formulas(10) = "Patente='" & Me.txtPatente & "'"
        .Formulas(11) = "Ano='" & Me.txtAño & "'"
        .Formulas(12) = "Marca='" & Me.lblMarca & "'"
        .Formulas(13) = "Modelo='" & Me.lblModelo & "'"
        .Formulas(14) = "Color='" & Me.lblColorE & "'"
        .Formulas(15) = "Kilometros='" & Me.txtKilAct & "'"
        .Formulas(16) = "FechaVenta='" & Format(Me.pckFecVta, "dd/mm/yyyy") & "'"
        .Formulas(17) = "NumeroChasis='" & Me.txtChasis.Text & "'"
        .Formulas(18) = "NumeroMotor='" & Me.lblMotor & "'"
        .Formulas(19) = "NumeroVin='" & Me.txtVin.Text & "'"

        .Formulas(20) = "FechaEmision='" & Format(Me.pckFechaAtencion, "long date") & "'"
        .Formulas(21) = "FechaReserva='" & Format(Me.pckFechaEntrega, "Long date") & "'"
'        .Formulas(22) = "HoraReserva='" & Me.cboHora.Text & "'"
        'kjcv 13.11.13
        .Formulas(22) = "HoraReserva='" & Me.txtHora.Text & "'"

        .Formulas(23) = "NombreRut='" & gstrNombreRut & "'"
        .Formulas(24) = "NombreComuna='" & gstrNombreComuna & "'"
        .Formulas(25) = "NombrePatente='" & gstrNombrePatente & "'"
        .Formulas(26) = "Recepcionista='" & gstrNombreRecepcionista & "'"
        .Formulas(27) = "RecepLlamado='" & gstrNombreRecepLlamado & "'"
        .Formulas(28) = "Observaciones='" & Me.txtComentario & "'"
        
        .Destination = crptToWindow
        .Action = True
   End With
   
   Screen.MousePointer = 1

Solucion:
    If Err.Number <> 0 Then
        MsgBox "Impresión Cancelada por el usuario", vbExclamation, "Imprimir"
        Screen.MousePointer = 1
        Exit Sub
    End If
End Sub

Sub Bloqueo(pstrEstado As String)
If pstrEstado = "V" Then
    'fmePat.Enabled = True
    mblnBloqueo = False
Else
    'fmePat.Enabled = False
    mblnBloqueo = True
End If
End Sub

Function DatosCliente(strIdCliente As String) As Boolean
If strIdCliente <> "" Then
    mstrSQL = "SELECT Glbl_Cliente_Proveedor.Razon_Social as NOMBRE, Glbl_Cliente_Proveedor.Direccion AS DIREC, Glbl_Comuna.Descripcion AS COMUNA, Glbl_Cliente_Proveedor.Rut AS RUT ,Glbl_Cliente_Proveedor.Telefono AS FONO FROM Glbl_Cliente_Proveedor INNER JOIN Glbl_Comuna ON Glbl_Cliente_Proveedor.Id_Comuna = Glbl_Comuna.Id_Comuna "
    mstrSQL = mstrSQL & " Where Glbl_Cliente_Proveedor.Id_Cliente_Proveedor='" & strIdCliente & "'"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoPrincipal
            If Not .BOF And Not .EOF Then
                lblCliente = IIf(Not IsNull(!Nombre), !Nombre, "")
                txtDir = IIf(Not IsNull(!DirEC), !DirEC, "")
                txtComuna = IIf(Not IsNull(!Comuna), !Comuna, "")
                txtRut = IIf(Not IsNull(!rut), !rut, "")
                lblFono = ValorNulo(!FONO)
            End If
        End With
    End If
    Conexion.CloseHost AdoPrincipal
End If
End Function
Sub GenerarOTReserva()
If Me.optSinPatente.Value = False Then
    GrabaReserva
    GrabaEncabezadoOt
    '//LREYES...
    GrabaServiciosMecanicaOtros
    gstrBusca = OrdenesdeTrabajo(Conexion, gstrIdUsuario, "Tllr", "", gstrIdEmpresa, gstrPathReporte, "R-" & lblNroRecepcion, apeditar)
    gstrImpresion = "O"
    gstrProcedencia = "Movimientos"
Else
    MsgBox "No puede confirmar una Hora sin Tener " & gstrNombrePatente, vbExclamation, "Advertencia"
End If
End Sub
Sub ActivarReserva()
    gstrSql = "UPDATE TLLR_ReservaHora SET ESTADO = 'V' ,"
    gstrSql = gstrSql & "Fecha_Activacion = '" & Date & "' , "
    gstrSql = gstrSql & "Quien_Activa = '" & gstrIdUsuario & "' "
    gstrSql = gstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' AND Tllr_ReservaHora.Id_Reserva = '" & lblNroRecepcion & "'"
    If Conexion.SendHost(gstrSql, , adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
        lblEstadoOTValor = "VIGENTE"
        tlbBarraHerramientas.Buttons.Item(2).Enabled = True     'guardar
        tlbBarraHerramientas.Buttons.Item(19).Enabled = False   'ACTIVAR
        tlbBarraHerramientas.Buttons.Item(18).Enabled = True    'ANULAR
        tlbBarraHerramientas.Buttons.Item(17).Enabled = True    'LIQUIDAR
    End If
    MsgBox "La Reserva Nº " & lblNroRecepcion & " Fue Activada"
    Bloqueo "V"
End Sub
Sub GrabaEncabezadoOt()
    mstrSQL = "INSERT INTO Tllr_OT "
    mstrSQL = mstrSQL & " (Id_Empresa, Id_Sucursal, "
    mstrSQL = mstrSQL & " Id_OT , Seccion_OT, "
    mstrSQL = mstrSQL & " Id_Garantia, Folio_Garantia, "
    mstrSQL = mstrSQL & " Id_Tipo_Cono, Nro_Cono, "
    mstrSQL = mstrSQL & " Patente, RealizadoPor,"
    mstrSQL = mstrSQL & " Kilometros_Recepcion, Id_Compañia_seguro,"
    mstrSQL = mstrSQL & " Fecha_Proxima_Visita, "                           'Fecha_Liquidacion,"
    mstrSQL = mstrSQL & " Estado,Fecha_Emision, "
    mstrSQL = mstrSQL & " Entrega_Estimada, Hora_Entrega, "
    mstrSQL = mstrSQL & " Nro_Factura_Emitida,Nro_Presupuesto_Origen,"
    mstrSQL = mstrSQL & " Nro_Siniestro, Nro_Poliza, Liquidador, "
    mstrSQL = mstrSQL & " Comentario, Solicitado_Por,"
    mstrSQL = mstrSQL & " Deducible_UF , Deducible_Pesos, "
    mstrSQL = mstrSQL & " Total_Mecanica,Total_Carroceria,"
    mstrSQL = mstrSQL & " Total_Desabolladura,Total_Pintura,"
    mstrSQL = mstrSQL & " Total_Terceros,Total_Repuestos,"
    mstrSQL = mstrSQL & " Total_Materiales,Total_Insumos, "
    mstrSQL = mstrSQL & " Total_Otros,Total_Ot,"
    mstrSQL = mstrSQL & " Total_OT_Iva,Total_IVA,Id_Cliente_Proveedor ) "
    mstrSQL = mstrSQL & " VALUES ("
    mstrSQL = mstrSQL & " '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "',"
    mstrSQL = mstrSQL & " '" & "R-" & lblNroRecepcion & "', '" & IIf(optMecanica = True, "M", "C") & "',"
    mstrSQL = mstrSQL & " '" & gstrIdTipoOtDefecto & "','" & "S/F" & "',"
    mstrSQL = mstrSQL & " '" & "01" & "', " & 0 & ","
    mstrSQL = mstrSQL & " '" & txtPatente.Text & "','" & txtRecepcionista & "',"
    mstrSQL = mstrSQL & " " & CLng(txtKilAct) & ", '" & "00" & "',"   'OJO
    mstrSQL = mstrSQL & " '" & CDate(DateAdd("d", 365, pckFechaAtencion.Value)) & "', "
    mstrSQL = mstrSQL & " 'R','" & CDate(pckFechaEntrega.Value) & "', "
'    mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & cboHora.Text & "' , "
    'kjcv 13.11.13
    mstrSQL = mstrSQL & " '" & CDate(pckFechaEntrega) & "' , '" & txtHora.Text & "' , "
    mstrSQL = mstrSQL & " '" & "S/N" & "', '" & "S/N" & "',"
    mstrSQL = mstrSQL & " '" & "S/N" & " ','" & "S/N" & "','" & "S/L" & "' , "
    mstrSQL = mstrSQL & " '" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), "S/C") & "' , '" & "S/S" & "' ,"
    mstrSQL = mstrSQL & " " & 0 & " , " & 0 & " ,"
    mstrSQL = mstrSQL & " " & 0 & " ," & 0 & ","
    mstrSQL = mstrSQL & " " & 0 & "," & 0 & ","
    mstrSQL = mstrSQL & " " & 0 & "," & 0 & ","
    mstrSQL = mstrSQL & " " & 0 & ", " & 0 & ", "
    mstrSQL = mstrSQL & " " & 0 & ", " & 0 & " ,"
    mstrSQL = mstrSQL & " " & 0 & " ," & 0 & ",'" & lblIdCliente & "')"
    
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
      '/////////////////////////////// AQUI GUARDAR DATOS DEL VEHICULO
    '//////////////////////////////////
      mblnTablaVacia = False
      'ActivaBotones
      Me.Tag = ""
    End If
End Sub
Sub GrabaReserva()

    mstrSQL = "UPDATE Tllr_ReservaHora "
    mstrSQL = mstrSQL & " SET Patente='" & txtPatente.Text & "', "
    mstrSQL = mstrSQL & " Estado='C', "
    mstrSQL = mstrSQL & " RealizadoPor='" & dtcRecepcionista.BoundText & "',"
    mstrSQL = mstrSQL & " Fecha_Emision='" & CDate(pckFechaAtencion) & "', "
    mstrSQL = mstrSQL & " Fecha_Reserva='" & CDate(pckFechaEntrega) & "', "
'    mstrSQL = mstrSQL & " Hora_Reserva='" & cboHora.Text & "', "
    'kjcv 13.11.13
    mstrSQL = mstrSQL & " Hora_Reserva='" & txtHora.Text & "', "
    mstrSQL = mstrSQL & " Seccion_OT='" & IIf(Me.optMecanica = True, "M", "C") & "', "
    mstrSQL = mstrSQL & " Fecha_Confirmacion='" & Date & "', "
    mstrSQL = mstrSQL & " Quien_Confirma='" & gstrUsuario & "', "
    mstrSQL = mstrSQL & " Reparacion='" & IIf(Trim(txtComentario.Text) <> "", UCase(Trim(txtComentario.Text)), ".") & "',"
    mstrSQL = mstrSQL & " Total_Mecanica= 0,"
    mstrSQL = mstrSQL & " Total_Otros= 0,"
    mstrSQL = mstrSQL & " Total_Repuestos= 0,"
    mstrSQL = mstrSQL & " Total_Reserva= 0 "
    mstrSQL = mstrSQL & " WHERE Id_Empresa ='" & gstrIdEmpresa & "' And Id_Sucursal ='" & gstrIdSucursal & "' And Id_Reserva ='" & Trim(Trim(lblNroRecepcion)) & "' "
    
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        Me.lblEstadoOTValor = "CONFIRMADA"
        ActivaBotones
        Bloqueo "C"
        Me.Tag = ""
    End If '//////////////

End Sub
Sub LlamarTelefono(strPrograma As String, strNombre As String, strFono As String, strComentario As String)
    Dim ValDev&
    Dim i As Integer
    Dim strFono1 As String
    Dim gstrTomarLinea As String
    
    strFono = Trim(strFono)
    strFono1 = ""
    For i = 1 To Len(strFono)
        If Mid(strFono, i, 1) = "" Then
            Exit For
        End If
        strFono1 = strFono1 & Mid(strFono, i, 1)
    Next
    strFono1 = gstrTomarLinea & strFono1
    ValDev = tapiRequestMakecall(strFono1, App.ProductName, strNombre, strComentario)
End Sub
Private Sub ActualizaTotales()
    Dim i As Integer
    Dim dblHorasMecanica As Double
    Dim dblHorasOtrosServicios As Double
    
    dblHorasMecanica = 0
    dblHorasOtrosServicios = 0
    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
        dblHorasMecanica = dblHorasMecanica + CDbl(Me.lvwServiciosMecanica.ListItems(i).SubItems(2))
    Next

    dblHorasOtrosServicios = 0
    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
        dblHorasOtrosServicios = dblHorasOtrosServicios + CDbl(Me.lvwOtrosServicios.ListItems(i).SubItems(1))
    Next


    Me.lblHorasMecanica = Format(dblHorasMecanica, "#,##0.0")
    Me.lblHorasOtrosServicios = Format(dblHorasOtrosServicios, "#,##0.0")
    Me.lblTotalHoras(0) = Format(dblHorasMecanica + dblHorasOtrosServicios, "#,##0.0")
    Me.lblTotalHoras(1) = Format(dblHorasMecanica + dblHorasOtrosServicios, "#,##0.0")
   
End Sub
Private Sub GrabaServiciosMecanicaOtros()
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    Dim i As Integer
    
    
    
    For i = 1 To Me.lvwServiciosMecanica.ListItems.Count
        strSql = "insert into Tllr_Mecanica_Ot (Id_Empresa, Id_Sucursal, Id_OT, Seccion_OT, Id_Marca, Id_Modelo, Id_Servicio, Id_Tipo_Cargo, "
        strSql = strSql & "Mecanico_Designado, Horas, Valor, SubTotal, Porcentaje_Descuento, Monto_Descuento, Facturado, Id_Tarea, Estado_Tarea, horasReales) "
        strSql = strSql & "values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', '" & "R-" & lblNroRecepcion & "', '" & IIf(optMecanica = True, "M", "C") & "', "
        strSql = strSql & "'" & Me.lblIdMarca & "', '" & Me.lblIdModelo & "', '" & Me.lvwServiciosMecanica.ListItems(i) & "', '" & gstrIdCargoDefecto & "', '" & gstrMecanicoDefectoSecMec & "', "
        strSql = strSql & " " & CDbl(Me.lvwServiciosMecanica.ListItems(i).SubItems(2)) & ", " & ValorHora(gstrIdEmpresa, gstrIdSucursal) & ", " & Round(ValorHora(gstrIdEmpresa, gstrIdSucursal) * CDbl(Me.lvwServiciosMecanica.ListItems(i).SubItems(2)), gintDecimalesMoneda) & ", 0, 0, 'N', NULL, NULL, 0) "
        
        
        If Conexion.SendHost(strSql, , , , gcTiempoEspera) = apOk Then
            '/////////////////////////////// AQUI GUARDAR DATOS DEL VEHICULO
            '//////////////////////////////////
          mblnTablaVacia = False
          'ActivaBotones
          Me.Tag = ""
        End If
    Next

    For i = 1 To Me.lvwOtrosServicios.ListItems.Count
    
        strSql = "insert into Tllr_Otro_Ot (Id_Empresa, Id_Sucursal, Id_OT, Seccion_OT, Id_Otro_Servicio, Id_Tipo_Cargo, "
        strSql = strSql & "Mecanico_Asignado, Horas, Valor, SubTotal, Porcentaje_Descuento, Monto_Descuento, Descripcion_Otro, "
        strSql = strSql & "Facturado, HorasReales, Id_Tarea, Estado_Tarea) "
        strSql = strSql & "values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', '" & "R-" & lblNroRecepcion & "', '" & IIf(optMecanica = True, "M", "C") & "', '" & i & "', "
        strSql = strSql & "'" & gstrIdCargoDefecto & "', '" & gstrMecanicoDefectoSecMec & "', " & CDbl(Me.lvwOtrosServicios.ListItems(i).SubItems(1)) & ", "
        strSql = strSql & ValorHora(gstrIdEmpresa, gstrIdSucursal) & ", " & Round(ValorHora(gstrIdEmpresa, gstrIdSucursal) * CDbl(Me.lvwOtrosServicios.ListItems(i).SubItems(1)), gintDecimalesMoneda) & ", "
        strSql = strSql & "0, 0, '" & UCase(Trim(Me.lvwOtrosServicios.ListItems(i))) & "', 'N', 0, null, null)"


        If Conexion.SendHost(strSql, , , , gcTiempoEspera) = apOk Then
            '/////////////////////////////// AQUI GUARDAR DATOS DEL VEHICULO
            '//////////////////////////////////
          mblnTablaVacia = False
          'ActivaBotones
          Me.Tag = ""

        End If
    Next
End Sub
