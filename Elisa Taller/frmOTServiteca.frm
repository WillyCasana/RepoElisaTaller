VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOtServiteca 
   Caption         =   "Orden de Trabajo - Serviteca"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   Icon            =   "frmOTServiteca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Cliente "
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7575
      Begin MSComctlLib.Toolbar tlbCliente 
         Height          =   330
         Left            =   7140
         TabIndex        =   133
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Cliente"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   285
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Nombre del Cliente (Puede ingresar el RUT)"
         Top             =   360
         Width           =   6975
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc datMarca 
         Height          =   270
         Left            =   2280
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
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
         Caption         =   "datMarca"
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
      Begin MSDataListLib.DataCombo dbcboMarca 
         Bindings        =   "frmOTServiteca.frx":179A
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Marca"
         Text            =   "dbcboMarca"
      End
      Begin MSAdodcLib.Adodc datModelo 
         Height          =   270
         Left            =   5280
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   476
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
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
         Caption         =   "datModelo"
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
      Begin MSDataListLib.DataCombo dbcboModelo 
         Bindings        =   "frmOTServiteca.frx":17B1
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Modelo"
         Text            =   "dbcboModelo"
      End
      Begin VB.Label Label3 
         Caption         =   "Modelo Vehiculo"
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Marca Vehiculo"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Placa"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   7800
      TabIndex        =   11
      Top             =   480
      Width           =   3855
      Begin VB.TextBox txtEstadoOt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "VIGENTE"
         Top             =   1320
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc datAtendiodoPor 
         Height          =   270
         Left            =   1200
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   476
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
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
         Caption         =   "datAtendidoPor"
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
      Begin MSDataListLib.DataCombo dbcboAtendidoPor 
         Bindings        =   "frmOTServiteca.frx":17C9
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   880
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Mecanico"
         Text            =   "dbcboAtendidoPro"
      End
      Begin VB.Label Label11 
         Caption         =   "Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1370
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Atendido por:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   680
         Width           =   1095
      End
      Begin VB.Label lblNumeroOt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Orden de Trabajo N°"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   400
         Width           =   1575
      End
   End
   Begin VB.Frame frmFechasHoras 
      Height          =   495
      Left            =   6360
      TabIndex        =   135
      Top             =   2160
      Width           =   5295
      Begin VB.Label lblUltModif 
         Alignment       =   1  'Right Justify
         Caption         =   "00/00/0000 00:00"
         Height          =   255
         Left            =   3840
         TabIndex        =   139
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ult. Modificación:"
         Height          =   195
         Left            =   2520
         TabIndex        =   138
         Top             =   195
         Width           =   1230
      End
      Begin VB.Label lblApertura 
         Caption         =   "00/00/0000 00:00"
         Height          =   255
         Left            =   840
         TabIndex        =   137
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Apertura:"
         Height          =   195
         Left            =   120
         TabIndex        =   136
         Top             =   195
         Width           =   645
      End
   End
   Begin VB.TextBox txtTotalOT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0"
      Top             =   7080
      Width           =   2055
   End
   Begin TabDlg.SSTab tab 
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Revisión de Seguridad"
      TabPicture(0)   =   "frmOTServiteca.frx":17E7
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Servicios Solicitados"
      TabPicture(1)   =   "frmOTServiteca.frx":1803
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lsvServicios"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "tlbServicios"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtTotalServicios"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Repuestos Solicitados"
      TabPicture(2)   =   "frmOTServiteca.frx":181F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTotalRepuestos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lsvRepuestos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "tlbRepuestos"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label12"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "General"
      TabPicture(3)   =   "frmOTServiteca.frx":183B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton Command20 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1020
            TabIndex        =   34
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command17 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   3660
            TabIndex        =   35
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton Command14 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   37
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command8 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   38
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command11 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   2940
            TabIndex        =   36
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   2940
            TabIndex        =   39
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command19 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   780
            TabIndex        =   45
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command16 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   3420
            TabIndex        =   44
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton Command13 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   43
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command10 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   2700
            TabIndex        =   42
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command7 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   41
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   2700
            TabIndex        =   40
            Top             =   600
            Width           =   255
         End
         Begin VB.Label txtNeumaTD 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   47
            Top             =   1200
            Width           =   300
         End
         Begin VB.Label txtNeumaDD 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   49
            Top             =   600
            Width           =   300
         End
         Begin VB.Label txtNeumaTI 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   2400
            TabIndex        =   48
            Top             =   1200
            Width           =   300
         End
         Begin VB.Label txtNeumaDI 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   2400
            TabIndex        =   50
            Top             =   600
            Width           =   300
         End
         Begin VB.Label txtEstadoNeuma 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label txtNeumaR 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3120
            TabIndex        =   57
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Repuesto"
            Height          =   255
            Left            =   2400
            TabIndex        =   56
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Neumáticos"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   990
            Width           =   1380
         End
         Begin VB.Label Label55 
            Caption         =   "Delantero I."
            Height          =   255
            Left            =   2400
            TabIndex        =   54
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label Label56 
            Caption         =   "Delantero D."
            Height          =   255
            Left            =   3960
            TabIndex        =   53
            Top             =   405
            Width           =   975
         End
         Begin VB.Label Label57 
            Caption         =   "Trasero I."
            Height          =   255
            Left            =   2400
            TabIndex        =   52
            Top             =   1005
            Width           =   1095
         End
         Begin VB.Label Label58 
            Caption         =   "Trasero D."
            Height          =   255
            Left            =   3960
            TabIndex        =   51
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label Label87 
            Caption         =   "NEUMATICOS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1575
         Left            =   -68640
         TabIndex        =   102
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton Command68 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   111
            Top             =   1155
            Width           =   255
         End
         Begin VB.CommandButton Command62 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   109
            Top             =   675
            Width           =   255
         End
         Begin VB.CommandButton Command65 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   2940
            TabIndex        =   110
            Top             =   1155
            Width           =   255
         End
         Begin VB.CommandButton Command59 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   2940
            TabIndex        =   108
            Top             =   675
            Width           =   255
         End
         Begin VB.CommandButton Command56 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   103
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton Command55 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   104
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton Command53 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   105
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command52 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   106
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command50 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   107
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command49 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   112
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command67 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   116
            Top             =   1155
            Width           =   255
         End
         Begin VB.CommandButton Command64 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   2700
            TabIndex        =   115
            Top             =   1155
            Width           =   255
         End
         Begin VB.CommandButton Command61 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   114
            Top             =   675
            Width           =   255
         End
         Begin VB.CommandButton Command58 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   2700
            TabIndex        =   113
            Top             =   675
            Width           =   255
         End
         Begin VB.Label txtAmorTI 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   2400
            TabIndex        =   117
            Top             =   1155
            Width           =   300
         End
         Begin VB.Label txtAmorTD 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   118
            Top             =   1155
            Width           =   300
         End
         Begin VB.Label txtAmorDD 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   119
            Top             =   675
            Width           =   300
         End
         Begin VB.Label txtAmorDI 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   2400
            TabIndex        =   120
            Top             =   675
            Width           =   300
         End
         Begin VB.Label Label53 
            Caption         =   "Trasero D."
            Height          =   255
            Left            =   3960
            TabIndex        =   132
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label52 
            Caption         =   "Trasero I."
            Height          =   255
            Left            =   2400
            TabIndex        =   131
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label51 
            Caption         =   "Delantero D."
            Height          =   255
            Left            =   3960
            TabIndex        =   130
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label50 
            Caption         =   "Delantero I."
            Height          =   255
            Left            =   2400
            TabIndex        =   129
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label54 
            Caption         =   "AMORTIGUADORES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   128
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label txtFreTD 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   127
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label txtFreTra 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   126
            Top             =   840
            Width           =   300
         End
         Begin VB.Label txtFreDe 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   125
            Top             =   600
            Width           =   300
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tren Delant."
            Height          =   195
            Left            =   120
            TabIndex        =   124
            Top             =   1080
            Width           =   885
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Traseros"
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delanteros"
            Height          =   195
            Left            =   120
            TabIndex        =   122
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label41 
            Caption         =   "FRENOS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   59
         Top             =   2520
         Width           =   5055
         Begin VB.CommandButton Command32 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   68
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton Command31 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   69
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton Command29 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   70
            Top             =   1320
            Width           =   255
         End
         Begin VB.CommandButton Command28 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   71
            Top             =   1320
            Width           =   255
         End
         Begin VB.CommandButton Command26 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   72
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton Command25 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   73
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton Command23 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   74
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command22 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   75
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   1620
            TabIndex        =   76
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   1380
            TabIndex        =   77
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command47 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   60
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton Command46 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   61
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton Command44 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   62
            Top             =   1320
            Width           =   255
         End
         Begin VB.CommandButton Command43 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   63
            Top             =   1320
            Width           =   255
         End
         Begin VB.CommandButton Command41 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   64
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton Command40 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   65
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton Command38 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   66
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command37 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   67
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command35 
            Appearance      =   0  'Flat
            Caption         =   "M"
            Height          =   255
            Left            =   4500
            TabIndex        =   78
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton Command34 
            Appearance      =   0  'Flat
            Caption         =   "B"
            Height          =   255
            Left            =   4260
            TabIndex        =   79
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label40 
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   101
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label txtDir 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   100
            Top             =   1560
            Width           =   300
         End
         Begin VB.Label Label38 
            Caption         =   "LUBRICACION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   99
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label txtLubLi 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   98
            Top             =   1320
            Width           =   300
         End
         Begin VB.Label txtLubPa 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   97
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label txtLubRa 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   96
            Top             =   840
            Width           =   300
         End
         Begin VB.Label txtLubAc 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   3960
            TabIndex        =   95
            Top             =   600
            Width           =   300
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liq. Frenos"
            Height          =   195
            Left            =   3000
            TabIndex        =   94
            Top             =   1320
            Width           =   780
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parabrisas"
            Height          =   195
            Left            =   3000
            TabIndex        =   93
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radiador"
            Height          =   195
            Left            =   3000
            TabIndex        =   92
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aceite Motor"
            Height          =   195
            Left            =   3000
            TabIndex        =   91
            Top             =   600
            Width           =   900
         End
         Begin VB.Label txtBatCa 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   90
            Top             =   1560
            Width           =   300
         End
         Begin VB.Label txtBatCo 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   89
            Top             =   1320
            Width           =   300
         End
         Begin VB.Label txtBatAn 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   88
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label txtBatVo 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   87
            Top             =   840
            Width           =   300
         End
         Begin VB.Label txtBatDe 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1080
            TabIndex        =   86
            Top             =   600
            Width           =   300
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calcomanía"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Correa"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anclaje"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltaje"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Densidad"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label15 
            Caption         =   "BATERIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtTotalRepuestos 
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
         Height          =   285
         Left            =   -65640
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtTotalServicios 
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
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   11175
         Begin MSComCtl2.UpDown updDiasLLamado 
            Height          =   285
            Left            =   3720
            TabIndex        =   134
            Top             =   1500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtDiasLLamado"
            BuddyDispid     =   196730
            OrigLeft        =   3720
            OrigTop         =   1500
            OrigRight       =   3960
            OrigBottom      =   1785
            Max             =   365
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDiasLLamado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "10"
            Top             =   1500
            Width           =   495
         End
         Begin VB.TextBox txtObservaciones 
            Appearance      =   0  'Flat
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   480
            Width           =   10935
         End
         Begin VB.Label Label9 
            Caption         =   "días más."
            Height          =   255
            Left            =   4080
            TabIndex        =   21
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Se sugiere llamar al Cliente en por lo menos"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   3135
         End
         Begin VB.Label Label7 
            Caption         =   "Observaciones"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSComctlLib.Toolbar tlbServicios 
         Height          =   540
         Left            =   120
         TabIndex        =   10
         Top             =   3910
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   953
         ButtonWidth     =   1217
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               ImageKey        =   "Seleccion1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Editar"
               ImageKey        =   "Editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               ImageKey        =   "Borrar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvServicios 
         Height          =   3375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo Servicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Servicio Especifico"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Descuento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Mecánico Asignado"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Id_Tipo_Servicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Id_Servicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Id_Mecanico"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lsvRepuestos 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo Parte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Solicitado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Despachado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Descuento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Id_Item"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbRepuestos 
         Height          =   540
         Left            =   -74880
         TabIndex        =   23
         Top             =   3915
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   953
         ButtonWidth     =   1217
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Editar"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Servicios Solicitados"
         Height          =   255
         Left            =   -68160
         TabIndex        =   32
         Top             =   4125
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Servicios Solicitados"
         Height          =   255
         Left            =   6840
         TabIndex        =   25
         Top             =   4125
         Width           =   2415
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
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
            Object.Visible         =   0   'False
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Registro (Ctrl+D)"
            ImageKey        =   "Borrar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro (Ctrl+B)"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
            ImageKey        =   "Primero"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
            ImageKey        =   "Siguiente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro (Ctrl+U)"
            ImageKey        =   "Ultimo"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+Q)"
            ImageKey        =   "Cerrar"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Liquidar"
            Object.ToolTipText     =   "Estados Orden de Trabajo"
            ImageKey        =   "Seleccion"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Liquidar"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cmdialogo 
      Left            =   -120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
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
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1857
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1969
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1A7B
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1B8D
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1C9F
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1DB1
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1EC3
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":1FD5
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":20E7
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":21F9
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":230B
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":241D
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":252F
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":2641
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":2753
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":2865
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":2977
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":2DC9
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOTServiteca.frx":321B
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNumDoc 
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
      Height          =   285
      Left            =   120
      TabIndex        =   140
      Top             =   7080
      Width           =   5655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Orden de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   30
      Top             =   7080
      Width           =   2415
   End
End
Attribute VB_Name = "frmOtServiteca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Item As ListItem

Dim adoPrincipal As New ADODB.Recordset

Dim mstrSql As String
Dim mblnTablaVacia As Boolean

Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Dim mblnSW As Boolean

Private Sub Grabar_Correlativo()
Dim tablaVerifica As New ADODB.Recordset
Dim tablaCorrelativo As New ADODB.Recordset
Dim sql As String
Dim lstrLaTabla As String
Dim ldblNumero As Double

lstrLaTabla = "Srvt_Correlativo_OT"
ldblNumero = CDbl(Me.lblNumeroOt.Caption)

sql = ""
sql = "SELECT * FROM " & lstrLaTabla
sql = sql & " WHERE Ultimo_Numero = " & ldblNumero & " "
sql = sql & "AND " & lstrLaTabla & ".Id_Empresa = '" & gstrIdEmpresa & "' "
sql = sql & "AND " & lstrLaTabla & ".Id_Sucursal = '" & gstrIdSucursal & "' "
If Conexion.SendHost(sql, tablaVerifica, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If tablaVerifica.EOF = False And tablaVerifica.BOF = False Then
        Conexion.CloseHost tablaVerifica
        Exit Sub
    End If
End If
Conexion.CloseHost tablaVerifica

sql = ""
sql = "INSERT INTO " & lstrLaTabla
sql = sql & " (Id_Empresa, Id_Sucursal, Ultimo_Numero)"
sql = sql & " Values ("
sql = sql & "'" & gstrIdEmpresa & "', "
sql = sql & "'" & gstrIdSucursal & "', "
sql = sql & CDbl(ldblNumero) & ") "
Conexion.SendHost sql, , , , gcTiempoEspera

End Sub

Function ExistePatente(Patente As String) As Boolean
Dim tablaPatente As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT * FROM Tllr_Vehiculo_Cliente WHERE Patente='" & Patente & "'"
If Conexion.SendHost(sql, tablaPatente, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaPatente.EOF = False And tablaPatente.BOF = False Then
        ExistePatente = True
    Else
        ExistePatente = False
    End If
Else
    MsgBox "Problemas de conexión con el Servidor."
End If
Conexion.CloseHost tablaPatente

End Function

Public Sub GrabePatente()
Dim Tabla As New ADODB.Recordset
Dim sql As String

If ExistePatente(Me.txtPatente.Text) = False Then
    sql = ""
    sql = "INSERT INTO Tllr_Vehiculo_Cliente (Id_Marca, Id_Modelo, Id_Cliente_Proveedor, Patente, Fecha_Ingreso, Usr_Id, Usr_Fecha,Id_color_Exterior) "
    sql = sql & "VALUES ("
    sql = sql & "'" & Me.dbcboMarca.BoundText & "', "
    sql = sql & "'" & Me.dbcboModelo.BoundText & "', "
    sql = sql & "'" & Me.txtCliente.Tag & "', "
    sql = sql & "'" & Me.txtPatente.Text & "', "
    sql = sql & "'" & Format(Date, "dd/mm/yyyy") & "', "
    sql = sql & "'" & gstrIdUsuario & "', "
    sql = sql & "'" & Format(Date, "dd/mm/yyyy") & "','01')"
Else
    sql = ""
    sql = sql & "UPDATE Tllr_Vehiculo_Cliente SET "
    sql = sql & "Patente='" & Me.txtPatente.Text & "', "
    sql = sql & "Id_Marca='" & Me.dbcboMarca.BoundText & "', "
    sql = sql & "Id_Modelo='" & Me.dbcboModelo.BoundText & "', "
    sql = sql & "Id_Cliente_Proveedor='" & Me.txtCliente.Tag & "' "
    sql = sql & " WHERE Patente='" & Me.txtPatente.Text & "'"
End If
If Conexion.SendHost(sql, Tabla, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apAbort Then
    MsgBox "Problemas de conexión con el Servidor. No se ha guardado la totalidad de los registros."
End If
Conexion.CloseHost Tabla

End Sub

Public Sub TraeMarcaModelo(Patente As String)
Dim Tabla As New ADODB.Recordset
Dim sql As String
Dim lintPregunta As Integer

sql = ""
sql = "SELECT Id_Marca, Id_Modelo FROM Tllr_Vehiculo_Cliente WHERE Patente='" & Patente & "'"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then
        Me.dbcboMarca.BoundText = Tabla!Id_Marca
        LLena_Modelo Tabla!Id_Marca, Tabla!Id_Modelo
        Me.dbcboModelo.Refresh
        Me.dbcboModelo.BoundText = Tabla!Id_Modelo
    Else
        lintPregunta = MsgBox("La " & gstrNombrePatente & " " & Me.txtPatente.Text & " no se encuentra en los registros." & Chr(13) & "¿Desea ingresarla ahora?", vbQuestion + vbYesNo + vbDefaultButton2, "Maestro de " & gstrNombrePatente)
        If lintPregunta = 6 Then
            Load frmMantenedorVehiculoCliente
            frmMantenedorVehiculoCliente.txtPatente.Text = Me.txtPatente.Text
            frmMantenedorVehiculoCliente.Show vbModal
            TraeMarcaModelo (Me.txtPatente.Text)
        End If
    End If
Else
    MsgBox "Problemas de conexión con el Servidor."
End If
Conexion.CloseHost Tabla

End Sub

Function SumaServicios() As Double
Dim ldblCont As Double

SumaServicios = 0

For ldblCont = 1 To Me.lsvServicios.ListItems.Count
    SumaServicios = SumaServicios + CDbl(SacarFormatoValor(Me.lsvServicios.ListItems(ldblCont).SubItems(6), gstrMonedaLocal))
Next ldblCont

End Function

Function SumaRepuestos() As Double
Dim ldblCont As Double

SumaRepuestos = 0

For ldblCont = 1 To Me.lsvRepuestos.ListItems.Count
    SumaRepuestos = SumaRepuestos + CDbl(SacarFormatoValor(Me.lsvRepuestos.ListItems(ldblCont).SubItems(7), gstrMonedaLocal))
Next ldblCont

End Function

Function SumaOT() As Double

SumaOT = 0

SumaOT = SumaOT + CDbl(SacarFormatoValor(Me.txtTotalServicios.Text, gstrMonedaLocal))
SumaOT = SumaOT + CDbl(SacarFormatoValor(Me.txtTotalRepuestos.Text, gstrMonedaLocal))



End Function

Public Sub LLena_Mecanico()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = sql & "SELECT Id_Mecanico, Nombre FROM Tllr_Mecanicos "
sql = sql & "WHERE Es_Recepcionista='S' "
sql = sql & "AND Vigencia='S'"
sql = sql & "AND Tllr_Mecanicos.Id_Empresa = '" & gstrIdEmpresa & "' "
sql = sql & "AND Tllr_Mecanicos.Id_Sucursal = '" & gstrIdSucursal & "' "
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datAtendiodoPor.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Marcas()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT * FROM Glbl_Marca WHERE Vigencia='S' Order by Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datMarca.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Modelos(IdMarca As String)
Dim Tabla As New ADODB.Recordset
Dim sql As String

Me.dbcboModelo.Text = ""
sql = ""
sql = "SELECT * FROM Glbl_Modelo WHERE Id_Marca='" & IdMarca & "' AND Vigencia='S' order by Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datModelo.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Modelo(IdMarca As String, IdModelo As String)
Dim Tabla As New ADODB.Recordset
Dim sql As String

Me.dbcboModelo.Text = ""
sql = ""
sql = "SELECT * FROM Glbl_Modelo WHERE Id_Marca='" & IdMarca & "' AND Id_Modelo='" & IdModelo & "' AND Vigencia='S'"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datModelo.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Private Sub dbcboMarca_Click(Area As Integer)
If Area = 2 Then
    LLena_Modelos (Me.dbcboMarca.BoundText)
End If
End Sub

Private Sub dbcboModelo_Click(Area As Integer)
If Area = 1 Then
    LLena_Modelos (Me.dbcboMarca.BoundText)
End If
End Sub

Private Sub Form_Load()
    mblnSW = True
    Me.Label1.Caption = gstrNombrePatente
    Retorno = Space$(128)
    tam = Len(Retorno)
    Valido = GetPrivateProfileString("TLLR", "PROXIMOLLAMADO", "", Retorno, tam, "AutoPro.ini")
    gstrDiasProximoLLamado = Trim$(Left$(Retorno, Valido))
    
    LLena_Marcas
    LLena_Mecanico
    LimpiaCampos
    ActivaBotones
    Me.tab.tab = 1
    Me.tab.TabVisible(0) = False
 '   Me.UltimoRegistro
    
End Sub

Private Sub lblNumeroOt_Change()

Me.tlbBarraHerramientas.Buttons(2).Enabled = True
Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(1).Enabled = True
Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(2).Enabled = True
Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(3).Enabled = True
If Mid$(Me.txtEstadoOt.Text, 1, 1) = "L" Then
    Me.tlbBarraHerramientas.Buttons(2).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(1).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(2).Enabled = True
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(3).Enabled = True
End If
If Mid$(Me.txtEstadoOt.Text, 1, 1) = "V" Then
    Me.tlbBarraHerramientas.Buttons(2).Enabled = True
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(1).Enabled = True
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(2).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(3).Enabled = True
End If
If Mid$(Me.txtEstadoOt.Text, 1, 1) = "N" Then
    Me.tlbBarraHerramientas.Buttons(2).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(1).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(2).Enabled = True
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(3).Enabled = False
End If
If Mid$(Me.txtEstadoOt.Text, 1, 1) = "F" Then
    Me.tlbBarraHerramientas.Buttons(2).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(1).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(2).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(3).Enabled = False
End If
If Mid$(Me.txtEstadoOt.Text, 1, 1) = "B" Then
    Me.tlbBarraHerramientas.Buttons(2).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(1).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(2).Enabled = False
    Me.tlbBarraHerramientas.Buttons(18).ButtonMenus(3).Enabled = False
End If

If Mid$(Me.txtEstadoOt.Text, 1, 1) = "L" Or Mid$(Me.txtEstadoOt.Text, 1, 1) = "F" Or Mid$(Me.txtEstadoOt.Text, 1, 1) = "B" Or Mid$(Me.txtEstadoOt.Text, 1, 1) = "N" Then
    Me.txtCliente.Enabled = False
    Me.txtPatente.Enabled = False
    Me.dbcboMarca.Enabled = False
    Me.dbcboModelo.Enabled = False
    Me.dbcboAtendidoPor.Enabled = False
    Me.tlbRepuestos.Buttons(1).Enabled = False
    Me.tlbRepuestos.Buttons(2).Enabled = False
    Me.tlbRepuestos.Buttons(3).Enabled = False
    Me.lsvRepuestos.Enabled = False
    Me.tlbServicios.Buttons(1).Enabled = False
    Me.tlbServicios.Buttons(2).Enabled = False
    Me.tlbServicios.Buttons(3).Enabled = False
    Me.lsvServicios.Enabled = False
    Me.tlbCliente.Buttons(1).Enabled = False
    Me.updDiasLLamado.Enabled = False
    Me.txtObservaciones.Enabled = False
Else
    If Me.lsvRepuestos.ListItems.Count = 0 Then
        Me.tlbRepuestos.Buttons(2).Enabled = False
        Me.tlbRepuestos.Buttons(3).Enabled = False
    Else
        Me.tlbRepuestos.Buttons(2).Enabled = True
        Me.tlbRepuestos.Buttons(3).Enabled = True
    End If
    
    If Me.lsvServicios.ListItems.Count = 0 Then
        Me.tlbServicios.Buttons(2).Enabled = False
        Me.tlbServicios.Buttons(3).Enabled = False
    Else
        Me.tlbServicios.Buttons(2).Enabled = True
        Me.tlbServicios.Buttons(3).Enabled = True
    End If
    Me.txtCliente.Enabled = True
    Me.txtPatente.Enabled = True
    Me.dbcboMarca.Enabled = True
    Me.dbcboModelo.Enabled = True
    Me.dbcboAtendidoPor.Enabled = True
    Me.tlbRepuestos.Buttons(1).Enabled = True
    Me.lsvRepuestos.Enabled = True
    Me.tlbServicios.Buttons(1).Enabled = True
    Me.lsvServicios.Enabled = True
    Me.tlbCliente.Buttons(1).Enabled = True
    Me.updDiasLLamado.Enabled = True
    Me.txtObservaciones.Enabled = True
End If
End Sub

Private Sub lsvRepuestos_DblClick()
If Me.lsvRepuestos.ListItems.Count > 0 Then
    frmEditaRepuesto.Show 1
    Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
End If
End Sub

Private Sub lsvRepuestos_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim ldblCont As Double

If Me.lsvRepuestos.ListItems.Count = 0 Then
    Me.tlbRepuestos.Buttons(2).Enabled = False
    Me.tlbRepuestos.Buttons(3).Enabled = False
Else
    Me.tlbRepuestos.Buttons(2).Enabled = True
    Me.tlbRepuestos.Buttons(3).Enabled = True
End If

For ldblCont = 1 To Me.lsvRepuestos.ListItems.Count
    If ldblCont > Me.lsvRepuestos.ListItems.Count Then
        Me.tlbRepuestos.Buttons(3).Enabled = False
        Exit For
    End If
    If Me.lsvRepuestos.ListItems(ldblCont).Selected = True Then
        If Me.lsvRepuestos.ListItems(ldblCont).SubItems(5) = "0" Then
            Me.tlbRepuestos.Buttons(3).Enabled = True
        Else
            Me.tlbRepuestos.Buttons(3).Enabled = False
        End If
    End If
Next ldblCont
End Sub

Private Sub lsvServicios_DblClick()
If Me.lsvServicios.ListItems.Count > 0 Then
    gblnNuevo = False
    frmEditaServicio.Show 1
    Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
End If
End Sub

Private Sub lsvServicios_ItemClick(ByVal Item As MSComctlLib.ListItem)

If Me.lsvServicios.ListItems.Count = 0 Then
    Me.tlbServicios.Buttons(2).Enabled = False
    Me.tlbServicios.Buttons(3).Enabled = False
Else
    Me.tlbServicios.Buttons(2).Enabled = True
    Me.tlbServicios.Buttons(3).Enabled = True
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lintRespuesta As Integer
Dim lstrTmp As String

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
        If Not Validacion() Then
            Exit Sub
        End If
        If Trim$(Me.lblNumeroOt.Caption) <> "?" And Trim$(Me.lblNumeroOt.Caption) <> "" Then
            Renovar
        End If
        lintRespuesta = MsgBox("¿Desea Imprimir la Orden de Trabajo N°" & Me.lblNumeroOt.Caption & "?", 36, "ServiPro")
        If lintRespuesta = 6 Then
            GrabarRegistro
            On Error Resume Next
            Cmdialogo.Flags = &H80000 Or &H40000 Or &H1
            Cmdialogo.CancelError = True
            Cmdialogo.Action = 5
            ImprimirOT
            If Err.Number <> 0 Then
                MsgBox "Impresión Cancelada.", vbInformation, "ServiPro"
            End If
        End If
    Case "Primero"
        PrimerRegistro
    Case "Anterior"
        RegistroAnterior
    Case "Siguiente"
        RegistroSiguiente
    Case "Ultimo"
        UltimoRegistro
    Case "Renovar"
        Renovar
    Case "Cerrar"
        CerrarSalir
        Screen.MousePointer = vbDefault
        Exit Sub
End Select
lstrTmp = Me.lblNumeroOt.Caption
Me.lblNumeroOt.Caption = ""
Me.lblNumeroOt.Caption = lstrTmp

Screen.MousePointer = vbDefault

End Sub
Private Sub Form_Activate()
    'If Me.lblNumeroOt.Caption = "0" Then AgregarRegistro
    If gstrBusca = "" Then
        AgregarRegistro
    End If
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
            ImprimirOT
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
            'Renovar
        Case 17 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub

Private Sub LiquidarOT()
Dim Tabla As New ADODB.Recordset
Dim sql As String
Dim lintRespuesta As Integer
Dim lstrTmp As String

lintRespuesta = MsgBox("Está a punto de Liquidar la Orden de Trabajo N°" & Me.lblNumeroOt.Caption & Chr(13) & "¿Desea Continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Liquidación de OT")
If lintRespuesta = 6 Then
    sql = ""
    sql = "UPDATE Srvt_OT SET "
    sql = sql & "Estado='L', "
    sql = sql & "Usr_Fecha = '" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "' "
    sql = sql & "WHERE Id_OT=" & CDbl(Me.lblNumeroOt.Caption)
    sql = sql & " AND Id_Empresa='" & gstrIdEmpresa & "' "
    sql = sql & " AND Id_Sucursal='" & gstrIdSucursal & "' "
    If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apAbort Then
        MsgBox "Problemas de conexión con el Servidor." & Chr(13) & "No se han actualizado los datos.", vbCritical, "ServiPro"
    Else
        Me.txtEstadoOt.Text = "LIQUIDADA"
        Me.lblUltModif.Caption = ValorNulo(Format$(Now, "dd/MM/yyyy HH:mm"))
    End If
End If

lstrTmp = Me.lblNumeroOt.Caption
Me.lblNumeroOt.Caption = ""
Me.lblNumeroOt.Caption = lstrTmp

lintRespuesta = MsgBox("¿Desea Imprimir la Orden de Trabajo N°" & Me.lblNumeroOt.Caption & "?", 36, "ServiPro")
If lintRespuesta = 6 Then
    On Error Resume Next
    Cmdialogo.Flags = &H80000 Or &H40000 Or &H1
    Cmdialogo.CancelError = True
    Cmdialogo.Action = 5
    ImprimirOT
    If Err.Number <> 0 Then
        MsgBox "Impresión Cancelada.", vbInformation, "ServiPro"
    End If
End If

End Sub

Private Sub ActivarOT()
Dim Tabla As New ADODB.Recordset
Dim sql As String
Dim lintRespuesta As Integer
Dim lstrTmp As String

lintRespuesta = MsgBox("Está a punto de Activar la Orden de Trabajo N°" & Me.lblNumeroOt.Caption & Chr(13) & "¿Desea Continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Liquidación de OT")
If lintRespuesta = 6 Then
    sql = ""
    sql = "UPDATE Srvt_OT SET "
    sql = sql & "Estado='V', "
    sql = sql & "Usr_Fecha = '" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "' "
    sql = sql & "WHERE Id_OT=" & CDbl(Me.lblNumeroOt.Caption)
    sql = sql & " AND Id_Empresa='" & gstrIdEmpresa & "' "
    sql = sql & " AND Id_Sucursal='" & gstrIdSucursal & "' "
    If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apAbort Then
        MsgBox "Problemas de conexión con el Servidor." & Chr(13) & "No se han actualizado los datos.", vbCritical, "ServiPro"
    Else
        Me.txtEstadoOt.Text = "VIGENTE"
        Me.lblUltModif.Caption = ValorNulo(Format$(Now, "dd/MM/yyyy HH:mm"))
    End If
End If

lstrTmp = Me.lblNumeroOt.Caption
Me.lblNumeroOt.Caption = ""
Me.lblNumeroOt.Caption = lstrTmp

lintRespuesta = MsgBox("¿Desea Imprimir la Orden de Trabako N°" & Me.lblNumeroOt.Caption & "?", 36, "ServiPro")
If lintRespuesta = 6 Then
    On Error Resume Next
    Cmdialogo.Flags = &H80000 Or &H40000 Or &H1
    Cmdialogo.CancelError = True
    Cmdialogo.Action = 5
    ImprimirOT
    If Err.Number <> 0 Then
        MsgBox "Impresión Cancelada.", vbInformation, "ServiPro"
    End If
End If

End Sub

Private Sub AnularOT()
Dim Tabla As New ADODB.Recordset
Dim sql As String
Dim lintRespuesta As Integer
Dim lstrTmp As String

lintRespuesta = MsgBox("Está a punto de Anular la Orden de Trabajo N°" & Me.lblNumeroOt.Caption & Chr(13) & "¿Desea Continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Liquidación de OT")
If lintRespuesta = 6 Then
    sql = ""
    sql = "UPDATE Srvt_OT SET "
    sql = sql & "Estado='N', "
    sql = sql & "Usr_Fecha = '" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "' "
    sql = sql & "WHERE Id_OT=" & CDbl(Me.lblNumeroOt.Caption)
    sql = sql & " AND Id_Empresa='" & gstrIdEmpresa & "' "
    sql = sql & " AND Id_Sucursal='" & gstrIdSucursal & "' "
    If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apAbort Then
        MsgBox "Problemas de conexión con el Servidor." & Chr(13) & "No se han actualizado los datos.", vbCritical, "ServiPro"
    Else
        Me.txtEstadoOt.Text = "NULA"
        Me.lblUltModif.Caption = ValorNulo(Format$(Now, "dd/MM/yyyy HH:mm"))
    End If
End If

lstrTmp = Me.lblNumeroOt.Caption
Me.lblNumeroOt.Caption = ""
Me.lblNumeroOt.Caption = lstrTmp

lintRespuesta = MsgBox("¿Desea Imprimir la Orden de Trabako N°" & Me.lblNumeroOt.Caption & "?", 36, "ServiPro")
If lintRespuesta = 6 Then
    On Error Resume Next
    Cmdialogo.Flags = &H80000 Or &H40000 Or &H1
    Cmdialogo.CancelError = True
    Cmdialogo.Action = 5
    ImprimirOT
    If Err.Number <> 0 Then
        MsgBox "Impresión Cancelada.", vbInformation, "ServiPro"
    End If
End If

End Sub

Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    If Me.txtCliente.Enabled = True Then Me.txtCliente.SetFocus
    Me.frmFechasHoras.Visible = False
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    Me.UltimoRegistro
    Me.frmFechasHoras.Visible = True
End Sub
Private Sub GrabarRegistro()
Dim Tabla As New ADODB.Recordset
Dim ldblCont As Double

If Trim$(Me.lblNumeroOt.Caption) = "?" Or Trim$(Me.lblNumeroOt.Caption) = "" Then
    Me.lblNumeroOt.Caption = TraeNumOT
End If

If Not Validacion() Then
    Exit Sub
End If

GrabePatente

If Me.Tag = "Crear" Then
    mstrSql = ""
    mstrSql = "INSERT INTO Srvt_OT (Id_Empresa, Id_Sucursal, Id_OT, Id_Cliente_Proveedor, Patente, Id_Mecanico, Fecha_Apertura, Observaciones, Dias_LLamado, Estado, Valor_OT, Usr_Id, Usr_Fecha) "
    mstrSql = mstrSql & "VALUES ("
    mstrSql = mstrSql & "'" & gstrIdEmpresa & "', "
    mstrSql = mstrSql & "'" & gstrIdSucursal & "', "
    mstrSql = mstrSql & CDbl(Me.lblNumeroOt.Caption) & ", "
    mstrSql = mstrSql & "'" & Me.txtCliente.Tag & "', "
    mstrSql = mstrSql & "'" & Me.txtPatente.Text & "', "
    mstrSql = mstrSql & "'" & Me.dbcboAtendidoPor.BoundText & "', "
    mstrSql = mstrSql & "'" & Format(Date, "dd/MM/yyyy") & " " & Format$(Time, "HH:mm:ss") & "', "
    mstrSql = mstrSql & "'" & Me.txtObservaciones.Text & "', "
    mstrSql = mstrSql & CDbl(Me.txtDiasLLamado.Text) & ", "
    mstrSql = mstrSql & "'" & Mid$(Me.txtEstadoOt.Text, 1, 1) & "', "
    mstrSql = mstrSql & CDbl(SacarFormatoValor(Me.txtTotalOT.Text, gstrMonedaLocal)) & ", "
    mstrSql = mstrSql & "'" & gstrIdUsuario & "', "
    mstrSql = mstrSql & "'" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:ss") & "')"
Else
    mstrSql = ""
    mstrSql = "UPDATE Srvt_OT SET "
    mstrSql = mstrSql & "Id_Cliente_Proveedor='" & Me.txtCliente.Tag & "', "
    mstrSql = mstrSql & "Patente='" & Me.txtPatente.Text & "', "
    mstrSql = mstrSql & "Id_Mecanico='" & Me.dbcboAtendidoPor.BoundText & "', "
    mstrSql = mstrSql & "Observaciones='" & Me.txtObservaciones.Text & "', "
    If Trim$(Me.txtDiasLLamado.Text) = "" Then Me.txtDiasLLamado.Text = "0"
    mstrSql = mstrSql & "Dias_LLamado=" & CDbl(Me.txtDiasLLamado.Text) & ", "
    mstrSql = mstrSql & "Estado='" & Mid$(Me.txtEstadoOt.Text, 1, 1) & "', "
    mstrSql = mstrSql & "Valor_OT=" & CDbl(SacarFormatoValor(Me.txtTotalOT.Text, gstrMonedaLocal)) & ", "
    mstrSql = mstrSql & "Usr_Id='" & gstrIdUsuario & "', "
    mstrSql = mstrSql & "Usr_Fecha='" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:ss") & "'"
    mstrSql = mstrSql & " WHERE Id_Empresa='" & gstrIdEmpresa & "' "
    mstrSql = mstrSql & " AND Id_Sucursal='" & gstrIdSucursal & "' "
    mstrSql = mstrSql & " AND Id_OT=" & CDbl(Me.lblNumeroOt.Caption)
End If
If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
    MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
End If



If Me.Tag = "Crear" Then
    For ldblCont = 1 To Me.lsvServicios.ListItems.Count
        mstrSql = ""
        mstrSql = "INSERT INTO Srvt_Servicios_OT (Id_Empresa, Id_Sucursal, Id_OT, Id_Concepto_Servicio, Id_Servicio, Id_Mecanico, Cantidad, Valor, Descuento, Total) "
        mstrSql = mstrSql & "VALUES ("
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "', "
        mstrSql = mstrSql & "'" & gstrIdSucursal & "', "
        mstrSql = mstrSql & CDbl(Me.lblNumeroOt.Caption) & ", "
        mstrSql = mstrSql & "'" & Me.lsvServicios.ListItems(ldblCont).SubItems(8) & "', "
        mstrSql = mstrSql & "'" & Me.lsvServicios.ListItems(ldblCont).SubItems(9) & "', "
        mstrSql = mstrSql & "'" & Me.lsvServicios.ListItems(ldblCont).SubItems(10) & "', "
        mstrSql = mstrSql & Me.lsvServicios.ListItems(ldblCont).SubItems(3) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvServicios.ListItems(ldblCont).SubItems(4), gstrMonedaLocal) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvServicios.ListItems(ldblCont).SubItems(5), gstrMonedaLocal) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvServicios.ListItems(ldblCont).SubItems(6), gstrMonedaLocal) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ")"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
            MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
        End If
    Next ldblCont
Else
    mstrSql = ""
    mstrSql = "DELETE FROM Srvt_Servicios_OT WHERE Id_OT=" & CDbl(Me.lblNumeroOt.Caption)
    If Conexion.SendHost(mstrSql, Tabla, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apAbort Then
        MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
    End If
    For ldblCont = 1 To Me.lsvServicios.ListItems.Count
        mstrSql = ""
        mstrSql = "INSERT INTO Srvt_Servicios_OT (Id_Empresa, Id_Sucursal, Id_OT, Id_Concepto_Servicio, Id_Servicio, Id_Mecanico, Cantidad, Valor, Descuento, Total) "
        mstrSql = mstrSql & "VALUES ("
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "', "
        mstrSql = mstrSql & "'" & gstrIdSucursal & "', "
        mstrSql = mstrSql & CDbl(Me.lblNumeroOt.Caption) & ", "
        mstrSql = mstrSql & "'" & Me.lsvServicios.ListItems(ldblCont).SubItems(8) & "', "
        mstrSql = mstrSql & "'" & Me.lsvServicios.ListItems(ldblCont).SubItems(9) & "', "
        mstrSql = mstrSql & "'" & Me.lsvServicios.ListItems(ldblCont).SubItems(10) & "', "
        mstrSql = mstrSql & Me.lsvServicios.ListItems(ldblCont).SubItems(3) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvServicios.ListItems(ldblCont).SubItems(4), gstrMonedaLocal) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvServicios.ListItems(ldblCont).SubItems(5), gstrMonedaLocal) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvServicios.ListItems(ldblCont).SubItems(6), gstrMonedaLocal) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ")"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
            MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
        End If
    Next ldblCont
End If



If Me.Tag = "Crear" Then
    For ldblCont = 1 To Me.lsvRepuestos.ListItems.Count
        mstrSql = ""
        mstrSql = "INSERT INTO Srvt_Repuestos_OT (Id_Empresa, Id_Sucursal, Linea, Id_OT, Id_Item, Valor_Unitario, Cant_Solicitado, Cant_Despachado, Descuento, Total, Costo) "
        mstrSql = mstrSql & "VALUES ("
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "', "
        mstrSql = mstrSql & "'" & gstrIdSucursal & "', "
        mstrSql = mstrSql & ldblCont & ", "
        mstrSql = mstrSql & CDbl(Me.lblNumeroOt.Caption) & ", "
        mstrSql = mstrSql & "'" & Me.lsvRepuestos.ListItems(ldblCont).SubItems(9) & "', "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvRepuestos.ListItems(ldblCont).SubItems(3) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(4) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(5) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(6) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvRepuestos.ListItems(ldblCont).SubItems(7), gstrMonedaLocal) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(8) & ")"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
            MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
        End If
    Next ldblCont
Else
    mstrSql = ""
    mstrSql = "DELETE FROM Srvt_Repuestos_OT WHERE Id_OT=" & CDbl(Me.lblNumeroOt.Caption)
    If Conexion.SendHost(mstrSql, Tabla, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apAbort Then
        MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
    End If
    For ldblCont = 1 To Me.lsvRepuestos.ListItems.Count
        mstrSql = ""
        mstrSql = "INSERT INTO Srvt_Repuestos_OT (Id_Empresa, Id_Sucursal, Linea, Id_OT, Id_Item, Valor_Unitario, Cant_Solicitado, Cant_Despachado, Descuento, Total, Costo) "
        mstrSql = mstrSql & "VALUES ("
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "', "
        mstrSql = mstrSql & "'" & gstrIdSucursal & "', "
        mstrSql = mstrSql & ldblCont & ", "
        mstrSql = mstrSql & CDbl(Me.lblNumeroOt.Caption) & ", "
        mstrSql = mstrSql & "'" & Me.lsvRepuestos.ListItems(ldblCont).SubItems(9) & "', "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvRepuestos.ListItems(ldblCont).SubItems(3), gstrMonedaLocal) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(4) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(5) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(6) & ", "
        mstrSql = mstrSql & SacarFormatoValor(Me.lsvRepuestos.ListItems(ldblCont).SubItems(7), gstrMonedaLocal) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto) & ", "
        mstrSql = mstrSql & Me.lsvRepuestos.ListItems(ldblCont).SubItems(8) & ")"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
            MsgBox "No se ha podido establecer conexión con el Servidor." & Chr(13) & "No se ha guardado la totalidad de los registros."
        End If
    Next ldblCont
End If

Grabar_Correlativo

mblnTablaVacia = False
ActivaBotones
Me.Tag = ""
Me.frmFechasHoras.Visible = True
End Sub
Private Sub BorrarRegistro()
'    Screen.MousePointer = vbDefault
'    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
'        mstrsql = "DELETE FROM " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtCodigo & "'"
'        If Conexion.SendHost(mstrsql, , , , gcTiempoEspera) = apOk Then
'            mstrsql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
'            If Conexion.SendHost(mstrsql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
'                    LeerCampos
'                Else
'                    mstrsql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo
'                    If Conexion.SendHost(mstrsql, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'                        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
'                            LeerCampos
'                        Else
'                            mblnTablaVacia = True
'                            LimpiaCampos
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        Conexion.CloseHost AdoPrincipal
'    End If
End Sub

Public Sub TraeDesdeFuera()
If gstrBusca <> "" Then
    If IsNumeric(gstrBusca) Then
        mstrSql = ""
        mstrSql = "SELECT Srvt_OT.*, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, Glbl_Cliente_Proveedor.Razon_Social "
        mstrSql = mstrSql & "FROM (Srvt_OT LEFT JOIN Tllr_Vehiculo_Cliente ON Srvt_OT.Patente = Tllr_Vehiculo_Cliente.Patente) LEFT JOIN Glbl_Cliente_Proveedor ON Srvt_OT.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
        mstrSql = mstrSql & " WHERE Srvt_OT.Id_OT=" & CDbl(gstrBusca) & " "
        mstrSql = mstrSql & " AND Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
        mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
        mstrSql = mstrSql & "ORDER BY Srvt_OT.Id_OT"
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos_OT
            End If
        End If
        Conexion.CloseHost adoPrincipal
        '
        Me.lsvServicios.ListItems.Clear
        mstrSql = ""
        mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
        mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
        mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(gstrBusca) & " "
        mstrSql = mstrSql & " AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
        mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
        mstrSql = mstrSql & "ORDER BY Srvt_Servicios_OT.Id_OT"
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos_Servicios
            End If
        End If
        Conexion.CloseHost adoPrincipal
        '
        Me.lsvRepuestos.ListItems.Clear
        mstrSql = ""
        mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
        mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
        mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(gstrBusca) & " "
        mstrSql = mstrSql & " AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
        mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
        mstrSql = mstrSql & "ORDER BY Srvt_Repuestos_OT.Id_OT"
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos_Repuestos
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
    Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
    Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)

End If

End Sub

Private Sub BuscarRegistro()

    gstrBusca = InputBox("Ingrese el número de la Orden de Trabajo que desea buscar:", "Buscar OT", "0")
    If gstrBusca <> "" Then
        If IsNumeric(gstrBusca) Then
            mstrSql = ""
            mstrSql = "SELECT Srvt_OT.*, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, Glbl_Cliente_Proveedor.Razon_Social "
            mstrSql = mstrSql & "FROM (Srvt_OT LEFT JOIN Tllr_Vehiculo_Cliente ON Srvt_OT.Patente = Tllr_Vehiculo_Cliente.Patente) LEFT JOIN Glbl_Cliente_Proveedor ON Srvt_OT.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
            mstrSql = mstrSql & " WHERE Srvt_OT.Id_OT=" & CDbl(gstrBusca) & " "
            mstrSql = mstrSql & " AND Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
            mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
            mstrSql = mstrSql & "ORDER BY Srvt_OT.Id_OT"
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos_OT
                End If
            End If
            Conexion.CloseHost adoPrincipal
            '
            Me.lsvServicios.ListItems.Clear
            mstrSql = ""
            mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
            mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
            mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(gstrBusca) & " "
            mstrSql = mstrSql & " AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
            mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
            mstrSql = mstrSql & "ORDER BY Srvt_Servicios_OT.Id_OT"
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos_Servicios
                End If
            End If
            Conexion.CloseHost adoPrincipal
            '
            Me.lsvRepuestos.ListItems.Clear
            mstrSql = ""
            mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
            mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
            mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(gstrBusca) & " "
            mstrSql = mstrSql & " AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
            mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
            mstrSql = mstrSql & "ORDER BY Srvt_Repuestos_OT.Id_OT"
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos_Repuestos
                End If
            End If
            Conexion.CloseHost adoPrincipal
        End If
        Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
        Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)

    End If
    Me.SetFocus
End Sub
Private Sub ImprimirOT()
Dim linea As String * 300
Dim i As Integer

Printer.FontName = "Times New Roman"
Printer.PaperSize = 1
Printer.FontBold = True
Printer.FontItalic = True
Printer.FontSize = 12

Printer.CurrentX = 200
Printer.CurrentY = 250
Printer.Print "Serviteca"

Printer.FontSize = 8
Printer.CurrentX = 7000
Printer.CurrentY = 250
Printer.Print "Fecha Impresión : " & Format(Now, "dd") + " de " + Format$(Now, "mmmm") + " de " + Format$(Now, "yyyy")
Printer.CurrentX = 7000
Printer.Print "Hora Impresión  : " & Format$(Now, "hh:mm:ss")
Printer.CurrentX = 7000
Printer.Print "Estado O.T.     : " & Me.txtEstadoOt.Text
Printer.FontSize = 14
Printer.CurrentX = 7000
Printer.Print "O.T. Nº         : " & Format$(Me.lblNumeroOt.Caption, "##,###,##0")
Printer.CurrentX = 7000
Printer.Print ""
Printer.CurrentX = 7000
Printer.Print ""
Printer.CurrentX = 3500
Printer.Print "ORDEN DE TRABAJO SERVITECA"

Printer.FontSize = 8
Printer.CurrentX = 200
Printer.Print ""
Printer.CurrentX = 200
Printer.Print "________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"

Printer.FontName = "Courier New"
Printer.FontSize = 10
Printer.FontItalic = False
Printer.FontBold = True

Printer.CurrentY = 2500
Printer.CurrentX = 200
Printer.Print "CLIENTE      : " & Trim$(Me.txtCliente.Text) & "   " & gstrNombreRut & ": " & IIf(gstrEditaRut = "S", Format$(Me.txtCliente.Tag, "@@.@@@.@@@-@"), Me.txtCliente.Tag)
Printer.CurrentX = 200
Printer.Print "VEHICULO     : " & Me.dbcboMarca.Text
Printer.CurrentX = 200
Printer.Print "MODELO       : " & Me.dbcboModelo.Text
Printer.CurrentX = 200
Printer.Print UCase(gstrNombrePatente) & "      : " & IIf(gstrValidaPatente = "S", Format$(Me.txtPatente.Text, "@@-@@@@"), Me.txtPatente.Text)
Printer.CurrentX = 200
Printer.Print ""
Printer.CurrentX = 200
Printer.Print "ATENDIDO POR : " & Me.dbcboAtendidoPor.Text
Printer.CurrentX = 200
Printer.Print ""
Printer.CurrentX = 200
Printer.Print "________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
Printer.CurrentX = 200
Printer.Print ""
Printer.FontSize = 14
Printer.CurrentX = 200
Printer.Print "SERVICIOS"
Printer.FontSize = 10
Printer.CurrentX = 200
Printer.Print ""

linea = ""
Mid$(linea, 1, 15) = "TIPO SERVICIO"
Mid$(linea, 18, 20) = "SERVICIO ESPECIFICO"
Mid$(linea, 40, 10) = "     VALOR"
Mid$(linea, 52, 8) = "CANTIDAD"
Mid$(linea, 61, 10) = " DESCUENTO"
Mid$(linea, 73, 10) = "     TOTAL"
Mid$(linea, 85, 10) = "MECANICO"

Printer.CurrentX = 200
Printer.Print linea

Printer.CurrentY = 5500
For i = 1 To Me.lsvServicios.ListItems.Count
    Printer.CurrentX = 200
    linea = ""
    Mid$(linea, 1, 15) = Me.lsvServicios.ListItems(i).SubItems(1)
    Mid$(linea, 18, 20) = Me.lsvServicios.ListItems(i).SubItems(2)
    Mid$(linea, 40, 10) = Format$(Format$(SacarFormatoValor(Me.lsvServicios.ListItems(i).SubItems(4), gstrMonedaLocal), "##,###,###"), "@@@@@@@@@@")
    Mid$(linea, 52, 8) = Format$(Format$(Me.lsvServicios.ListItems(i).SubItems(3), "###,###"), "@@@@@@@@")
    Mid$(linea, 61, 10) = Format$(Format$(Me.lsvServicios.ListItems(i).SubItems(5), "##,###,###"), "@@@@@@@@@@")
    Mid$(linea, 73, 10) = Format$(Format$(SacarFormatoValor(Me.lsvServicios.ListItems(i).SubItems(6), gstrMonedaLocal), "##,###,###"), "@@@@@@@@@@")
    Mid$(linea, 85, 10) = Me.lsvServicios.ListItems(i).SubItems(7)
    Printer.Print linea
Next i


Printer.Print ""
Printer.FontSize = 14
Printer.CurrentX = 200
Printer.Print "PRODUCTOS"
Printer.FontSize = 10
Printer.CurrentX = 200
Printer.Print ""
Printer.CurrentX = 200
linea = ""
Mid$(linea, 1, 15) = "CÓDIGO PARTE"
Mid$(linea, 18, 20) = "DESCRIPCION"
Mid$(linea, 40, 10) = "     VALOR"
Mid$(linea, 52, 8) = "SOLICIT."
Mid$(linea, 61, 8) = "DESPACH."
Mid$(linea, 73, 10) = " DESCUENTO"
Mid$(linea, 85, 10) = "     TOTAL"
Printer.Print linea

For i = 1 To Me.lsvRepuestos.ListItems.Count
    Printer.CurrentX = 200
    linea = ""
    Mid$(linea, 1, 15) = Me.lsvRepuestos.ListItems(i).SubItems(1)
    Mid$(linea, 18, 20) = Me.lsvRepuestos.ListItems(i).SubItems(2)
    Mid$(linea, 40, 10) = Format$(Format$(SacarFormatoValor(Me.lsvRepuestos.ListItems(i).SubItems(3), gstrMonedaLocal), "##,###,###"), "@@@@@@@@@@")
    Mid$(linea, 52, 8) = Format$(Format$(Me.lsvRepuestos.ListItems(i).SubItems(4), "###,###"), "@@@@@@@@")
    Mid$(linea, 61, 8) = Format$(Format$(Me.lsvRepuestos.ListItems(i).SubItems(5), "###,###"), "@@@@@@@@")
    Mid$(linea, 73, 10) = Format$(Format$(Me.lsvRepuestos.ListItems(i).SubItems(6), "##,###,###"), "@@@@@@@@@@")
    Mid$(linea, 85, 20) = Format$(Format$(SacarFormatoValor(Me.lsvRepuestos.ListItems(i).SubItems(7), gstrMonedaLocal), "##,###,###"), "@@@@@@@@@@")
    Printer.Print linea
Next i

Printer.CurrentX = 200
Printer.Print ""
Printer.CurrentX = 200
Printer.Print "COMENTARIO : " & Trim$(Me.txtObservaciones.Text)
Printer.CurrentY = 12000
Printer.CurrentX = 200
Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Printer.FontSize = 8
Printer.CurrentX = 200
Printer.Print "(Copia Bodega)"
Printer.CurrentX = 7500
Printer.Print "O.T. Nº : " & Format$(Me.lblNumeroOt.Caption, "##,###,##0")
Printer.CurrentX = 200
Printer.Print "Fecha  : " & Format(Now, "dd") & " de " + Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy") & "   -   Hora   : " & Format$(Now, "hh:mm:ss")
Printer.FontSize = 14
Printer.CurrentX = 200
Printer.Print "PRODUCTOS SOLICITADOS"
Printer.FontSize = 10
Printer.CurrentX = 200
linea = ""
Mid$(linea, 1, 15) = "CÓDIGO PARTE"
Mid$(linea, 18, 20) = "DESCRIPCION"
Mid$(linea, 40, 10) = "     VALOR"
Mid$(linea, 52, 8) = "SOLICIT."
Mid$(linea, 61, 8) = "DESPACH."
Mid$(linea, 73, 10) = " DESCUENTO"
Mid$(linea, 85, 10) = "     TOTAL"
Printer.Print linea

Printer.CurrentY = 13400
For i = 1 To Me.lsvRepuestos.ListItems.Count
    Printer.CurrentX = 200
    linea = ""
    Mid$(linea, 1, 15) = Me.lsvRepuestos.ListItems(i).SubItems(1)
    Mid$(linea, 18, 20) = Me.lsvRepuestos.ListItems(i).SubItems(2)
    Mid$(linea, 40, 10) = Format$(Format$(SacarFormatoValor(Me.lsvRepuestos.ListItems(i).SubItems(3), gstrMonedaLocal), "##,###,###"), "@@@@@@@@@@")
    Mid$(linea, 52, 8) = Format$(Format$(Me.lsvRepuestos.ListItems(i).SubItems(4), "###,###"), "@@@@@@@@")
    Mid$(linea, 61, 8) = Format$(Format$(Me.lsvRepuestos.ListItems(i).SubItems(5), "###,###"), "@@@@@@@@")
    Mid$(linea, 73, 10) = Format$(Format$(Me.lsvRepuestos.ListItems(i).SubItems(6), "##,###,###"), "@@@@@@@@@@")
    Mid$(linea, 85, 20) = Format$(Format$(SacarFormatoValor(Me.lsvRepuestos.ListItems(i).SubItems(7), gstrMonedaLocal), "##,###,###"), "@@@@@@@@@@")
    Printer.Print linea
Next i

Printer.EndDoc

End Sub
Private Sub PrimerRegistro()
    
    
    mstrSql = ""
    mstrSql = "SELECT TOP 1  Srvt_OT.*, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, Glbl_Cliente_Proveedor.Razon_Social "
    mstrSql = mstrSql & "FROM (Srvt_OT LEFT JOIN Tllr_Vehiculo_Cliente ON Srvt_OT.Patente = Tllr_Vehiculo_Cliente.Patente) LEFT JOIN Glbl_Cliente_Proveedor ON Srvt_OT.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
    mstrSql = mstrSql & "WHERE Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
    mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
    mstrSql = mstrSql & "ORDER BY Srvt_OT.Id_OT"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos_OT
        End If
    End If
    Conexion.CloseHost adoPrincipal
    
    
    Me.lsvServicios.ListItems.Clear
    mstrSql = ""
    mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
    mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
    mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
    mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
    mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
    mstrSql = mstrSql & "ORDER BY Srvt_Servicios_OT.Id_OT"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos_Servicios
        End If
    End If
    Conexion.CloseHost adoPrincipal


    Me.lsvRepuestos.ListItems.Clear
    mstrSql = ""
    mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
    mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
    mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
    mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
    mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
    mstrSql = mstrSql & "ORDER BY Srvt_Repuestos_OT.Id_OT"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos_Repuestos
        End If
    End If
    Conexion.CloseHost adoPrincipal

    Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
    Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
    
End Sub
Private Sub RegistroAnterior()
    
    mstrSql = ""
    mstrSql = "SELECT TOP 1  Srvt_OT.*, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, Glbl_Cliente_Proveedor.Razon_Social "
    mstrSql = mstrSql & "FROM (Srvt_OT LEFT JOIN Tllr_Vehiculo_Cliente ON Srvt_OT.Patente = Tllr_Vehiculo_Cliente.Patente) LEFT JOIN Glbl_Cliente_Proveedor ON Srvt_OT.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
    mstrSql = mstrSql & "WHERE Srvt_OT.Id_OT<" & Me.lblNumeroOt.Caption & " "
    mstrSql = mstrSql & "AND Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
    mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
    mstrSql = mstrSql & "ORDER BY Srvt_OT.Id_OT DESC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos_OT
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
    
    Me.lsvServicios.ListItems.Clear
    mstrSql = ""
    mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
    mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
    mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
    mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
    mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
    mstrSql = mstrSql & " ORDER BY Srvt_Servicios_OT.Id_OT DESC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos_Servicios
        End If
    End If
    Conexion.CloseHost adoPrincipal

    Me.lsvRepuestos.ListItems.Clear
    mstrSql = ""
    mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
    mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
    mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
    mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
    mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
    mstrSql = mstrSql & " ORDER BY Srvt_Repuestos_OT.Id_OT DESC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos_Repuestos
        End If
    End If
    Conexion.CloseHost adoPrincipal

    Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
    Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
    
End Sub
Private Sub RegistroSiguiente()
    
mstrSql = ""
mstrSql = "SELECT TOP 1  Srvt_OT.*, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, Glbl_Cliente_Proveedor.Razon_Social "
mstrSql = mstrSql & "FROM (Srvt_OT LEFT JOIN Tllr_Vehiculo_Cliente ON Srvt_OT.Patente = Tllr_Vehiculo_Cliente.Patente) LEFT JOIN Glbl_Cliente_Proveedor ON Srvt_OT.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
mstrSql = mstrSql & "WHERE Srvt_OT.Id_OT >'" & Me.lblNumeroOt.Caption & "' "
mstrSql = mstrSql & "AND Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & "ORDER BY Srvt_OT.Id_OT"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_OT
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal

Me.lsvServicios.ListItems.Clear
mstrSql = ""
mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & " ORDER BY Srvt_Servicios_OT.Id_OT"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_Servicios
    End If
End If
Conexion.CloseHost adoPrincipal


Me.lsvRepuestos.ListItems.Clear
mstrSql = ""
mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & " ORDER BY Srvt_Repuestos_OT.Id_OT"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_Repuestos
    End If
End If
Conexion.CloseHost adoPrincipal

Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
    
End Sub
Public Sub UltimoRegistro()
    
If Trim$(Me.lblNumeroOt.Caption) = "" Then Me.lblNumeroOt.Caption = "0"
    
mstrSql = ""
mstrSql = "SELECT TOP 1  Srvt_OT.*, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, Glbl_Cliente_Proveedor.Razon_Social "
mstrSql = mstrSql & "FROM (Srvt_OT LEFT JOIN Tllr_Vehiculo_Cliente ON Srvt_OT.Patente = Tllr_Vehiculo_Cliente.Patente) LEFT JOIN Glbl_Cliente_Proveedor ON Srvt_OT.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
mstrSql = mstrSql & "WHERE Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & " ORDER BY Srvt_OT.Id_OT DESC"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_OT
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal

Me.lsvServicios.ListItems.Clear
mstrSql = ""

mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & " ORDER BY Srvt_Servicios_OT.Id_OT DESC"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_Servicios
    End If
End If
Conexion.CloseHost adoPrincipal
    
    
Me.lsvRepuestos.ListItems.Clear
mstrSql = ""
mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & " ORDER BY Srvt_Repuestos_OT.Id_OT DESC"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_Repuestos
    End If
End If
Conexion.CloseHost adoPrincipal

Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
    
End Sub
Private Sub Renovar()

mstrSql = ""
mstrSql = "SELECT Srvt_OT.*, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, Glbl_Cliente_Proveedor.Razon_Social "
mstrSql = mstrSql & "FROM (Srvt_OT LEFT JOIN Tllr_Vehiculo_Cliente ON Srvt_OT.Patente = Tllr_Vehiculo_Cliente.Patente) LEFT JOIN Glbl_Cliente_Proveedor ON Srvt_OT.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor "
mstrSql = mstrSql & "WHERE Srvt_OT.Id_OT = '" & Me.lblNumeroOt.Caption & "' "
mstrSql = mstrSql & "AND Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_OT
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal

Me.lsvServicios.ListItems.Clear
mstrSql = ""
mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & " ORDER BY Srvt_Servicios_OT.Id_OT"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_Servicios
    End If
End If
Conexion.CloseHost adoPrincipal


Me.lsvRepuestos.ListItems.Clear
mstrSql = ""
mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
mstrSql = mstrSql & " ORDER BY Srvt_Repuestos_OT.Id_OT"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos_Repuestos
    End If
End If
Conexion.CloseHost adoPrincipal

Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()

    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = True
        .Item("Grabar").Enabled = True
        .Item("Cancelar").Enabled = False
        .Item("Borrar").Enabled = True
        .Item("Buscar").Enabled = True
        .Item("Imprimir").Enabled = True
        .Item("Primero").Enabled = True
        .Item("Anterior").Enabled = True
        .Item("Siguiente").Enabled = True
        .Item("Ultimo").Enabled = True
        .Item("Renovar").Enabled = True
        .Item("Cerrar").Enabled = True
        .Item("Liquidar").Enabled = True
    End With
End Sub
Private Sub DesactivaBotones()
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = False
        .Item("Grabar").Enabled = True
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
        .Item("Liquidar").Enabled = False
    End With
End Sub
Private Sub VerificaTablaVacia()
    If (Not adoPrincipal.BOF And Not adoPrincipal.EOF) And adoPrincipal.RecordCount > 0 Then
        mblnTablaVacia = False
    Else
        mblnTablaVacia = True
        LimpiaCampos
        MsgBox "La tabla no contiene registros...", vbInformation, "Advertencia"
    End If
End Sub
Private Sub LeerCampos_OT()

    If mblnTablaVacia Then
        LimpiaCampos
        Exit Sub
    End If

    With adoPrincipal
        Me.txtCliente.Tag = ValorNulo(!Id_Cliente_Proveedor)
        Me.txtCliente.Text = traeCLIENTE(ValorNulo(!Id_Cliente_Proveedor))
        Me.txtPatente.Text = ValorNulo(!Patente)
        Me.dbcboMarca.BoundText = ValorNulo(!Id_Marca)
        LLena_Modelo ValorNulo(!Id_Marca), ValorNulo(!Id_Modelo)
        Me.dbcboModelo.BoundText = ValorNulo(!Id_Modelo)
        Me.dbcboAtendidoPor.BoundText = ValorNulo(!Id_Mecanico)
        If Len(ValorNulo(!Fecha_Apertura)) = 10 Then
            Me.lblApertura.Caption = ValorNulo(Format$(!Fecha_Apertura, "dd/MM/yyyy"))
        Else
            Me.lblApertura.Caption = ValorNulo(Format$(!Fecha_Apertura, "dd/MM/yyyy HH:mm"))
        End If
        Me.lblUltModif.Caption = ValorNulo(Format$(!Usr_Fecha, "dd/MM/yyyy HH:mm"))
        Select Case ValorNulo(!estado)
            Case "V"
                Me.txtEstadoOt.Text = "VIGENTE"
            Case "N"
                Me.txtEstadoOt.Text = "NULA"
            Case "F"
                Me.txtEstadoOt.Text = "FACTURADA"
            Case "B"
                Me.txtEstadoOt.Text = "BOLETEADA"
            Case "C"
                Me.txtEstadoOt.Text = "CERRADA"
            Case "L"
                Me.txtEstadoOt.Text = "LIQUIDADA"
        End Select
        Me.lblNumeroOt.Caption = !Id_OT
        Me.txtObservaciones.Text = !Observaciones
        Me.txtDiasLLamado.Text = !Dias_LLamado
        Select Case Mid$(Me.txtEstadoOt.Text, 1, 1)
            Case "F"
                Me.lblNumDoc.Caption = "FACTURA N° " & ValorNulo(!Nro_Factura_Emitida)
            Case "B"
                Me.lblNumDoc.Caption = "BOLETA N° " & ValorNulo(!Nro_Factura_Emitida)
            Case Else
                Me.lblNumDoc.Caption = "s/d"
        End Select
    End With
    
    
    
'    Me.lsvServicios.ListItems.Clear
'    mstrSql = ""
'    mstrSql = "SELECT Srvt_Servicios_OT.*, Srvt_Concepto_Servicio.Descripcion AS Desc1, Srvt_Servicios.Descripcion AS Desc2, Tllr_Mecanicos.Nombre "
'    mstrSql = mstrSql & "FROM ((Srvt_Servicios_OT LEFT JOIN Srvt_Concepto_Servicio ON Srvt_Servicios_OT.Id_Concepto_Servicio = Srvt_Concepto_Servicio.Id_Concepto_Servicio) LEFT JOIN Srvt_Servicios ON Srvt_Servicios_OT.Id_Servicio = Srvt_Servicios.Id_Servicio) LEFT JOIN Tllr_Mecanicos ON Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
'    mstrSql = mstrSql & " WHERE Srvt_Servicios_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
'    mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
'    mstrSql = mstrSql & "AND Srvt_Servicios_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
'    mstrSql = mstrSql & " ORDER BY Srvt_Servicios_OT.Id_OT"
'    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
'        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
'            LeerCampos_Servicios
'        End If
'    End If
'    Conexion.CloseHost adoPrincipal
'
'
'    Me.lsvRepuestos.ListItems.Clear
'    mstrSql = ""
'    mstrSql = "SELECT Srvt_Repuestos_OT.Id_Item, Srvt_Repuestos_OT.Valor_Unitario, Srvt_Repuestos_OT.Cant_Solicitado, Srvt_Repuestos_OT.Cant_Despachado, Srvt_Repuestos_OT.Descuento, Srvt_Repuestos_OT.Total, Srvt_Repuestos_OT.Costo, Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo "
'    mstrSql = mstrSql & "FROM (Srvt_Repuestos_OT LEFT JOIN Stck_Item ON Srvt_Repuestos_OT.Id_Item = Stck_Item.Id_Item) "
'    mstrSql = mstrSql & "WHERE Srvt_Repuestos_OT.Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
'    mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
'    mstrSql = mstrSql & "AND Srvt_Repuestos_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
'    mstrSql = mstrSql & " ORDER BY Srvt_Repuestos_OT.Id_OT"
'    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
'        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
'            LeerCampos_Repuestos
'        End If
'    End If
'    Conexion.CloseHost adoPrincipal
'
'    Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
'    Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)

    
    
End Sub

Private Sub LeerCampos_Servicios()

    With adoPrincipal
        If .EOF = False And .BOF = False Then
            .MoveFirst
            While Not .EOF
                Set Item = Me.lsvServicios.ListItems.Add(, , Me.lsvServicios.ListItems.Count + 1)
                Item.SubItems(1) = !Desc1
                Item.SubItems(2) = !Desc2
                Item.SubItems(3) = !cantidad
                Item.SubItems(4) = FormatoValor(!Valor * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal, gintDecimalesMoneda)
                Item.SubItems(5) = !Descuento
                Item.SubItems(6) = FormatoValor(!Total * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal, gintDecimalesMoneda)
                If Not IsNull(!Nombre) Then
                    Item.SubItems(7) = !Nombre
                Else
                    Item.SubItems(7) = ""
                End If
                Item.SubItems(8) = ValorNulo(!Id_Concepto_Servicio)
                Item.SubItems(9) = ValorNulo(!Id_servicio)
                Item.SubItems(10) = ValorNulo(!Id_Mecanico)
                .MoveNext
            Wend
        End If
    End With
End Sub

Private Sub LeerCampos_Repuestos()

    With adoPrincipal
        If .EOF = False And .BOF = False Then
            .MoveFirst
            While .EOF = False
                Set Item = Me.lsvRepuestos.ListItems.Add(, , Me.lsvRepuestos.ListItems.Count + 1)
                Item.SubItems(1) = !prefijo & Guion & !basico & Guion & !sufijo
                Item.SubItems(2) = !Descripcion
                Item.SubItems(3) = FormatoValor(!Valor_Unitario * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal, gintDecimalesMoneda)
                Item.SubItems(4) = !Cant_Solicitado
                Item.SubItems(5) = !Cant_Despachado
                Item.SubItems(6) = !Descuento
                Item.SubItems(7) = FormatoValor(!Total * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal, gintDecimalesMoneda)
                Item.SubItems(8) = !Costo
                Item.SubItems(9) = !Id_Item
                .MoveNext
            Wend
        End If
    End With
End Sub

Private Sub LimpiaCampos()

Me.txtCliente.Text = ""
Me.txtCliente.Tag = ""
Me.txtPatente.Text = ""
Me.dbcboMarca.Text = ""
Me.dbcboModelo.Text = ""
Me.lblNumeroOt.Caption = ""
Me.dbcboAtendidoPor.Text = ""
Me.txtObservaciones.Text = ""
Me.txtDiasLLamado.Text = ""
Me.txtEstadoOt.Text = ""
Me.lsvServicios.ListItems.Clear
Me.lsvRepuestos.ListItems.Clear
Me.txtTotalRepuestos.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
Me.txtTotalServicios.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
Me.txtTotalOT.Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)

End Sub
Private Sub ValoresporDefecto()
Me.lblNumeroOt.Caption = "?" 'TraeNumOT
Me.txtDiasLLamado.Text = gstrDiasProximoLLamado
Me.txtEstadoOt.Text = "VIGENTE"
End Sub
Private Function Validacion() As Boolean
    Validacion = True
  
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As New ADODB.Recordset
        mstrSql = ""
        mstrSql = mstrSql & "SELECT * FROM Srvt_OT "
        mstrSql = mstrSql & "WHERE Id_OT=" & CDbl(Me.lblNumeroOt.Caption) & " "
        mstrSql = mstrSql & "AND Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' "
        mstrSql = mstrSql & "AND Srvt_OT.Id_Sucursal = '" & gstrIdSucursal & "' "
        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Registro ya ingresado.", vbInformation, "Advertencia"
                Validacion = False
                Exit Function
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
    
    If Me.dbcboAtendidoPor.BoundText = "" Then
        MsgBox "Debe seleccionar un Mecánico Recepcionista.", vbExclamation, "ServiPro"
        If Me.dbcboAtendidoPor.Enabled = True Then Me.dbcboAtendidoPor.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If Trim$(Me.txtPatente.Text) = "" Then
        MsgBox "Debe indicar una " & gstrNombrePatente, vbExclamation, "ServiPro"
        If Me.txtPatente.Enabled = True Then Me.txtPatente.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If Me.dbcboMarca.BoundText = "" Then
        MsgBox "Debe indicar una Marca.", vbExclamation, "ServiPro"
        If Me.dbcboMarca.Enabled = True Then Me.dbcboMarca.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If Me.dbcboModelo.BoundText = "" Then
        MsgBox "Debe indicar un Modelo.", vbExclamation, "ServiPro"
        If Me.dbcboModelo.Enabled = True Then Me.dbcboModelo.SetFocus
        Validacion = False
        Exit Function
    End If
    
End Function

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

Private Sub tlbBarraHerramientas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Index
    Case 1 'LIQUIDAR
        LiquidarOT
    Case 2 'ACTIVAR
        ActivarOT
    Case 3 'ANULAR
        AnularOT
End Select
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
'Select Case Button.Key
'    Case "Buscar"
'        gstrRutCliente = ""
'        gstrNombreCliente = ""
'        APFORM1.BuscarRegistroClientes Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa
'        'APFORM1.BuscarRegistroClientes Conexion, gstrRutCliente, gstrNombreCliente
'        If gstrRutCliente <> "" Then
'            Me.txtCliente.Text = gstrNombreCliente
'            Me.txtCliente.Tag = Trim$(gstrRutCliente)
'        End If
'End Select
End Sub

Private Sub tlbRepuestos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Tabla As New ADODB.Recordset
Dim sql As String
Dim ldblCont As Double
Dim lstrTmp As String
Dim lintRespuesta As Integer

Select Case Button.Index
    Case 1 'AGREGAR
        frmBuscar.Show 1
        If Trim$(gstrBusca) <> "" Then
            sql = ""
            sql = "SELECT Stck_Item.Descripcion, Stck_Item.Prefijo, Stck_Item.Basico, Stck_Item.Sufijo, Stck_Item.Precio_Venta, Stck_Saldos.Saldo, Stck_Item.Precio_Costo "
            sql = sql & "FROM Stck_Item LEFT JOIN Stck_Saldos ON Stck_Item.Id_Item = Stck_Saldos.Id_Item "
            sql = sql & "WHERE Stck_Item.Id_Item='" & gstrBusca & "'"
            If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Tabla.EOF = False And Tabla.BOF = False Then
                    Set Item = Me.lsvRepuestos.ListItems.Add(, , Me.lsvRepuestos.ListItems.Count + 1)
                    Item.SubItems(1) = Tabla!prefijo & Guion & Tabla!basico & Guion & Tabla!sufijo
                    Item.SubItems(2) = Tabla!Descripcion
                    Item.SubItems(3) = FormatoValor(Tabla!Precio_Venta * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal, gintDecimalesMoneda)
                    Item.SubItems(4) = "1"
                    Item.SubItems(5) = "0"
                    Item.SubItems(6) = "0"
                    Item.SubItems(7) = FormatoValor(Tabla!Precio_Venta * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal, gintDecimalesMoneda)
                    Item.SubItems(8) = Tabla!Precio_Costo
                    Item.SubItems(9) = gstrBusca
                    Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
                End If
            End If
            Conexion.CloseHost Tabla
        End If
    Case 2 'EDITAR
        If Me.lsvRepuestos.ListItems.Count > 0 Then
            frmEditaRepuesto.Show 1
            Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
        End If
    Case 3 'ELIMINAR
        For ldblCont = 1 To Me.lsvRepuestos.ListItems.Count
            If ldblCont > Me.lsvRepuestos.ListItems.Count Then
                Exit For
            End If
            If Me.lsvRepuestos.ListItems(ldblCont).Selected = True Then
                If Me.lsvRepuestos.ListItems(ldblCont).SubItems(5) = "0" Then
                    lintRespuesta = MsgBox("¿Está seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Registro")
                    If lintRespuesta = 6 Then
                        Me.lsvRepuestos.ListItems.Remove (ldblCont)
                        ldblCont = ldblCont - 1
                    End If
                Else
                    MsgBox "No es posible eliminar el registro de Repuestos seleccionado." & Chr(13) & "Stock ya capturó y rebajó el Repuesto.", vbExclamation, "ServiPro"
                End If
            End If
        Next ldblCont
        Me.txtTotalRepuestos.Text = FormatoValor(SumaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
End Select
lstrTmp = Me.lblNumeroOt.Caption
Me.lblNumeroOt.Caption = ""
Me.lblNumeroOt.Caption = lstrTmp
End Sub

Private Sub tlbServicios_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim ldblCont As Double
Dim lstrTmp As String
Dim lintRespuesta As Integer

Select Case Button.Index
    Case 1 'AGREGAR
        gblnNuevo = True
        frmEditaServicio.Show 1
        Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
    Case 2 'EDITAR
        If Me.lsvServicios.ListItems.Count > 0 Then
            gblnNuevo = False
            frmEditaServicio.Show 1
        End If
        Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
    Case 3 'ELIMINAR
        For ldblCont = 1 To Me.lsvServicios.ListItems.Count
            If ldblCont > Me.lsvServicios.ListItems.Count Then
                Exit For
            End If
            If Me.lsvServicios.ListItems(ldblCont).Selected = True Then
                lintRespuesta = MsgBox("¿Está seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Registro")
                If lintRespuesta = 6 Then
                    Me.lsvServicios.ListItems.Remove (ldblCont)
                    ldblCont = ldblCont - 1
                End If
            End If
        Next ldblCont
        Me.txtTotalServicios.Text = FormatoValor(SumaServicios, gstrMonedaLocal, gintDecimalesMoneda)
End Select
lstrTmp = Me.lblNumeroOt.Caption
Me.lblNumeroOt.Caption = ""
Me.lblNumeroOt.Caption = lstrTmp
End Sub

Private Sub txtCliente_GotFocus()
Me.txtCliente.SelStart = 0
Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim lintR As Integer

KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 8 Then
    Me.txtCliente.Text = ""
    Exit Sub
End If

If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 75 And KeyAscii <> 107 And KeyAscii <> 13 Then
        KeyAscii = 8
        Me.txtCliente.Text = ""
        Exit Sub
    End If
End If

If KeyAscii = 13 And Me.txtCliente.Text <> "" And Me.txtCliente.Text <> " " Then
    Screen.MousePointer = 11
    If RutValido(Trim$(Me.txtCliente.Text)) = False Then
        MsgBox gstrNombreRut & " inválido!" & Chr(13) & Chr(13) & "Por favor, intente nuevamente.", vbExclamation, "ServiPro"
        Me.txtCliente.SetFocus
        Me.txtCliente.SelStart = 0
        Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
        Screen.MousePointer = 1
        Exit Sub
    End If
    'Verifica si cliente existe....
    If ExisteCliente(Trim$(Me.txtCliente.Text)) = False Then
        lintR = MsgBox("Cliente no existe." & Chr(13) & "¿Desea crearlo ahora?", 36, "MAESTRO DE CLIENTES")
        If lintR = 6 Then
            gstrRutCliente = Me.txtCliente.Text
            apfFormulario.clientes Conexion, gstrUsuario, "Srvt", "", gstrIdEmpresa, gstrPathReporte, gstrRutCliente, gstrNombreCliente, apcrear
            Me.txtCliente.Text = UCase$(Trim$(gstrNombreCliente))
            Me.txtCliente.Tag = UCase$(Trim$(gstrRutCliente))
        End If
    Else
        Me.txtCliente.Tag = Me.txtCliente.Text
        Me.txtCliente.Text = traeCLIENTE(Me.txtCliente.Tag)
    End If
    Screen.MousePointer = 1
End If
End Sub

Private Sub txtCliente_LostFocus()
Dim lintR As Integer



If IsNumeric(Mid(Me.txtCliente.Text, 1, 2)) And Me.txtCliente.Text <> "" And Me.txtCliente.Text <> " " Then
    Screen.MousePointer = 11
    If RutValido(Trim$(Me.txtCliente.Text)) = False Then
        MsgBox gstrNombreRut & " inválido!" & Chr(13) & Chr(13) & "Por favor, intente nuevamente.", vbExclamation, "ServiPro"
        Me.txtCliente.SetFocus
        Me.txtCliente.SelStart = 0
        Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
        Screen.MousePointer = 1
        Exit Sub
    End If
    'Verifica si cliente existe....
    If ExisteCliente(Trim$(Me.txtCliente.Text)) = False Then
        lintR = MsgBox("Cliente no existe." & Chr(13) & "¿Desea crearlo ahora?", 36, "MAESTRO DE CLIENTES")
        If lintR = 6 Then
            gstrRutCliente = Me.txtCliente.Text
            apfFormulario.clientes Conexion, gstrUsuario, "Srvt", "", gstrIdEmpresa, gstrPathReporte, gstrRutCliente, gstrNombreCliente, apcrear, "Cliente - Proveedor", gstrIdSucursal
            Me.txtCliente.Text = UCase$(Trim$(gstrNombreCliente))
            Me.txtCliente.Tag = UCase$(Trim$(gstrRutCliente))
        End If
    Else
        Me.txtCliente.Tag = Me.txtCliente.Text
        Me.txtCliente.Text = traeCLIENTE(Me.txtCliente.Tag)
    End If
    Screen.MousePointer = 1
End If
End Sub

Private Sub txtCliente_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Me.txtCliente.Text <> "" Then
    Me.txtCliente.ToolTipText = gstrNombreRut & ": " & IIf(gstrEditaRut = "S", Format$(Trim$(Me.txtCliente.Tag), "@@.@@@.@@@-@"), Me.txtCliente.Tag)
Else
    Me.txtCliente.ToolTipText = "Nombre del cliente (puede digitar el & " & gstrNombreRut & ")"
End If
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

If KeyAscii = 13 Then
    Me.dbcboAtendidoPor.SetFocus
End If
End Sub

Private Sub txtPatente_LostFocus()
If Trim$(Me.txtPatente.Text) <> "" Then
    TraeMarcaModelo (Me.txtPatente.Text)
End If
End Sub

Private Sub txtTotalRepuestos_Change()
Me.txtTotalOT.Text = FormatoValor(SumaOT, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtTotalServicios_Change()
Me.txtTotalOT.Text = FormatoValor(SumaOT, gstrMonedaLocal, gintDecimalesMoneda)
End Sub
