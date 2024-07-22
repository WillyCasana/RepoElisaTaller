VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMantenedorParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración del Sistema"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   ClipControls    =   0   'False
   Icon            =   "frmMantenedorParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   8775
   Begin TabDlg.SSTab stbParametros 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Empresa"
      TabPicture(0)   =   "frmMantenedorParametros.frx":038A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame11(8)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Proceso"
      TabPicture(1)   =   "frmMantenedorParametros.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Valores"
      TabPicture(2)   =   "frmMantenedorParametros.frx":03C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame11 
         Caption         =   "Detalle"
         Height          =   7815
         Index           =   8
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   8295
         Begin MSDataListLib.DataCombo dtcSucursal 
            Bindings        =   "frmMantenedorParametros.frx":03DE
            Height          =   315
            Left            =   1260
            TabIndex        =   72
            Top             =   1125
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin VB.TextBox txtRazonSocial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   675
            Width           =   5000
         End
         Begin VB.TextBox txtidEmpresa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   5000
         End
         Begin VB.TextBox txtDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1560
            Width           =   5000
         End
         Begin MSAdodcLib.Adodc datSucursal 
            Height          =   330
            Left            =   4440
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "adodc1"
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
         Begin VB.TextBox txtSucursal 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1125
            Width           =   5000
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Razon Social"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   90
            TabIndex        =   37
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   36
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   35
            Top             =   1245
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7815
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Siguiente Folio para..."
         Top             =   360
         Width           =   8295
         Begin VB.Frame Frame11 
            Caption         =   "MultiMarca"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   9
            Left            =   120
            TabIndex        =   110
            Top             =   1440
            Width           =   3735
            Begin VB.CheckBox chkServiciosGenerales 
               Appearance      =   0  'Flat
               Caption         =   "Servicios Generales (Tempario)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   240
               TabIndex        =   113
               Top             =   600
               Width           =   3135
            End
            Begin VB.CommandButton cmdMarcas 
               Caption         =   "Marcas"
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
               TabIndex        =   112
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox chkPrecioMarca 
               Appearance      =   0  'Flat
               Caption         =   "Mano de Obra"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   240
               TabIndex        =   111
               Top             =   320
               Width           =   1815
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Repuestos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Index           =   0
            Left            =   120
            TabIndex        =   105
            Top             =   4440
            Width           =   3735
            Begin VB.TextBox txtDecimalesMoneda 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2280
               TabIndex        =   108
               Text            =   "0"
               Top             =   600
               Width           =   1380
            End
            Begin VB.TextBox txtMonedaLocal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2280
               TabIndex        =   106
               Text            =   "01"
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Decimales Moneda"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   109
               Top             =   600
               Width           =   1605
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Cod.Sigla Moneda Local"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   107
               Top             =   240
               Width           =   2040
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Repuestos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   6
            Left            =   120
            TabIndex        =   79
            Top             =   3720
            Width           =   3735
            Begin VB.TextBox txtDscMaxCIA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2400
               TabIndex        =   129
               Text            =   "15"
               Top             =   1080
               Width           =   1260
            End
            Begin VB.TextBox txtDescuentoMaximo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2400
               TabIndex        =   80
               Text            =   "15"
               Top             =   240
               Width           =   1260
            End
            Begin VB.Label Label19 
               Caption         =   "% Desc. Max. CIA SEG."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   128
               Top             =   1080
               Width           =   2055
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "% Descuento Máximo"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   81
               Top             =   240
               Width           =   1860
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Reserva de Horas y Hora de Entrega"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   54
            Top             =   2400
            Width           =   3735
            Begin VB.TextBox txtMinutos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2160
               MaxLength       =   2
               TabIndex        =   60
               Text            =   "30"
               Top             =   960
               Width           =   1500
            End
            Begin VB.TextBox txtHoraTermino 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2160
               MaxLength       =   2
               TabIndex        =   59
               Text            =   "20"
               Top             =   600
               Width           =   1500
            End
            Begin VB.TextBox txtHoraInicio 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2160
               MaxLength       =   2
               TabIndex        =   56
               Text            =   "08"
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label Label11 
               Caption         =   "Intervalo de Minutos"
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
               TabIndex        =   58
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label Label10 
               Caption         =   "Hora Termino"
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
               TabIndex        =   57
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label8 
               Caption         =   "Hora Inicio"
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
               TabIndex        =   55
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Otros"
            Height          =   2415
            Left            =   3960
            TabIndex        =   45
            Top             =   2400
            Width           =   4200
            Begin VB.TextBox txtValoIva 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   119
               Text            =   "18"
               Top             =   2040
               Width           =   1500
            End
            Begin VB.TextBox txtSeguroTaller 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   90
               Text            =   "0"
               Top             =   1680
               Width           =   1500
            End
            Begin VB.TextBox txtManoObraGtia 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   53
               Text            =   "0"
               Top             =   1320
               Width           =   1500
            End
            Begin VB.TextBox txtLineasRecepcion 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               MaxLength       =   2
               TabIndex        =   52
               Text            =   "9"
               Top             =   960
               Width           =   1500
            End
            Begin VB.TextBox txtValorExistencia 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   51
               Text            =   "2000000"
               Top             =   600
               Width           =   1500
            End
            Begin VB.TextBox txtHorasTrabajo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   50
               Text            =   "8"
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label Label15 
               Caption         =   "Valor IGV"
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
               TabIndex        =   120
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Seguro Taller"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   91
               Top             =   1695
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Mano Obra Garantía"
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
               TabIndex        =   49
               Top             =   1320
               Width           =   1935
            End
            Begin VB.Label Label6 
               Caption         =   "Lineas Recepción"
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
               TabIndex        =   48
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label5 
               Caption         =   "Valor Existencia"
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
               TabIndex        =   47
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "Numero Horas Trabajo"
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
               TabIndex        =   46
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Insumos"
            Height          =   975
            Left            =   3960
            TabIndex        =   44
            Top             =   1440
            Width           =   4200
            Begin VB.OptionButton optInsumos 
               Appearance      =   0  'Flat
               Caption         =   "Venta Insumos"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               TabIndex        =   97
               Top             =   600
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.TextBox txtCostoInsumos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   96
               Text            =   "0"
               Top             =   600
               Width           =   1500
            End
            Begin VB.TextBox txtInsumosMO 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               MaxLength       =   2
               TabIndex        =   93
               Text            =   "0"
               Top             =   240
               Width           =   1500
            End
            Begin VB.OptionButton optInsumosMO 
               Appearance      =   0  'Flat
               Caption         =   "% Sobre Mano Obra"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   92
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Consultas"
            Height          =   660
            Index           =   7
            Left            =   3960
            TabIndex        =   41
            Top             =   4800
            Width           =   4215
            Begin VB.TextBox txtRegistrosDefecto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   42
               Text            =   "25"
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Nº de Registros Defecto"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   75
               TabIndex        =   43
               Top             =   285
               Width           =   2055
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Costo Insumos"
            Height          =   1140
            Index           =   4
            Left            =   3960
            TabIndex        =   26
            Top             =   240
            Width           =   4200
            Begin VB.OptionButton optCostoInsumosPesos 
               Appearance      =   0  'Flat
               Caption         =   "Costo Insumos en (S/.)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   600
               Width           =   2345
            End
            Begin VB.OptionButton optCostoInsumosPorc 
               Appearance      =   0  'Flat
               Caption         =   "Costo Insumos en %"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Value           =   -1  'True
               Width           =   2175
            End
            Begin VB.TextBox txtCostoInsumosPesos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   87
               Text            =   "0"
               Top             =   600
               Width           =   1500
            End
            Begin VB.TextBox txtCostoInsumosPorc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               MaxLength       =   2
               TabIndex        =   86
               Text            =   "0"
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Margenes"
            Enabled         =   0   'False
            Height          =   1080
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   8400
            Visible         =   0   'False
            Width           =   7020
            Begin VB.TextBox txtMargenMateriales 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5475
               MaxLength       =   3
               TabIndex        =   19
               Text            =   "0"
               Top             =   585
               Width           =   675
            End
            Begin VB.TextBox txtMargenRepuestos 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5475
               MaxLength       =   3
               TabIndex        =   18
               Text            =   "0"
               Top             =   270
               Width           =   675
            End
            Begin VB.TextBox txtMargenLubricantes 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   15
               Text            =   "0"
               Top             =   555
               Width           =   675
            End
            Begin VB.TextBox txtMargenInsumos 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   14
               Text            =   "0"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   8
               Left            =   6255
               TabIndex        =   25
               Top             =   675
               Width           =   120
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   7
               Left            =   6255
               TabIndex        =   24
               Top             =   315
               Width           =   120
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   6
               Left            =   2700
               TabIndex        =   23
               Top             =   600
               Width           =   120
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   5
               Left            =   2685
               TabIndex        =   22
               Top             =   315
               Width           =   120
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Materiales"
               Height          =   195
               Index           =   4
               Left            =   3660
               TabIndex        =   21
               Top             =   615
               Width           =   720
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Repuestos"
               Height          =   195
               Index           =   4
               Left            =   3660
               TabIndex        =   20
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lubricantes"
               Height          =   195
               Index           =   2
               Left            =   75
               TabIndex        =   17
               Top             =   570
               Width           =   825
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Insumos"
               Height          =   195
               Index           =   3
               Left            =   75
               TabIndex        =   16
               Top             =   285
               Width           =   585
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Hora Hombre"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   3735
            Begin VB.TextBox txtCostoManoObra 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2040
               TabIndex        =   10
               Text            =   "0"
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtPrecioManoObra 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2040
               TabIndex        =   9
               Text            =   "0"
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Costo de Mano de Obra"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Precio Mano Obra"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   11
               Top             =   600
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   1
         ToolTipText     =   "Siguiente Folio para..."
         Top             =   360
         Width           =   8295
         Begin VB.TextBox txtCargoGtiaFabrica 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6240
            MaxLength       =   25
            TabIndex        =   126
            Top             =   5580
            Width           =   975
         End
         Begin VB.CheckBox chkAsignaRecursos 
            Appearance      =   0  'Flat
            Caption         =   "Asignación de Recursos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   125
            Top             =   5280
            Width           =   2415
         End
         Begin VB.CheckBox chkImprimeImagen 
            Appearance      =   0  'Flat
            Caption         =   "Imprime Imagen Vehículo, Inventario"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   124
            Top             =   4920
            Width           =   3615
         End
         Begin VB.CheckBox chkValidaCostoRepuestos 
            Appearance      =   0  'Flat
            Caption         =   "Valida Costo de Repuestos en la OT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   123
            Top             =   3840
            Width           =   3375
         End
         Begin VB.CheckBox chkBloqueaSubtotalRep 
            Appearance      =   0  'Flat
            Caption         =   "Bloquea Subtotal Repuestos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   122
            Top             =   4200
            Width           =   2775
         End
         Begin VB.CheckBox chkValidaServiciosCero 
            Appearance      =   0  'Flat
            Caption         =   "Valida Servicios en 0 en Liquidación"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   121
            Top             =   4560
            Width           =   3375
         End
         Begin VB.Frame Frame16 
            Caption         =   "Deducibles"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   114
            Top             =   4920
            Width           =   3495
            Begin VB.TextBox txtDeducibleMas 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               MaxLength       =   4
               TabIndex        =   116
               Top             =   120
               Width           =   900
            End
            Begin VB.TextBox txtDeducibleMenos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               MaxLength       =   4
               TabIndex        =   115
               Top             =   480
               Width           =   900
            End
            Begin VB.Label Label17 
               Caption         =   "Código Cargo Deducible (+)"
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
               TabIndex        =   118
               Top             =   240
               Width           =   2535
            End
            Begin VB.Label Label16 
               Caption         =   "Código Cargo Deducible (-)"
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
               TabIndex        =   117
               Top             =   525
               Width           =   2415
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Familia de Repuestos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   98
            Top             =   3720
            Width           =   3495
            Begin VB.TextBox txtCodigoInsumos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               MaxLength       =   4
               TabIndex        =   101
               Text            =   "0"
               Top             =   840
               Width           =   900
            End
            Begin VB.TextBox txtCodigoMateriales 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               MaxLength       =   4
               TabIndex        =   100
               Text            =   "0"
               Top             =   480
               Width           =   900
            End
            Begin VB.TextBox txtCodigoLubricantes 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               MaxLength       =   4
               TabIndex        =   99
               Text            =   "0"
               Top             =   120
               Width           =   900
            End
            Begin VB.Label Label14 
               Caption         =   "Código Familia Insumos"
               Height          =   255
               Left            =   120
               TabIndex        =   104
               Top             =   885
               Width           =   2415
            End
            Begin VB.Label Label13 
               Caption         =   "Código Familia Materiales"
               Height          =   255
               Left            =   120
               TabIndex        =   103
               Top             =   585
               Width           =   2295
            End
            Begin VB.Label Label3 
               Caption         =   "Código Familia Lubricantes"
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   260
               Width           =   2295
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Nota Presupuesto"
            Height          =   855
            Left            =   120
            TabIndex        =   84
            Top             =   6720
            Visible         =   0   'False
            Width           =   6615
            Begin VB.TextBox txtNotaPresupuesto 
               Height          =   495
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   85
               Top             =   240
               Width           =   6375
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Nota Recepción"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   82
            Top             =   5880
            Width           =   6975
            Begin VB.TextBox txtNotaRecepcion 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   83
               Top             =   240
               Width           =   6375
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Presupuestos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   77
            Top             =   3000
            Width           =   3495
            Begin VB.CheckBox chkTraspasaRepuestos 
               Appearance      =   0  'Flat
               Caption         =   "Traspasa Repuestos a OT"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Cambia Dias Habiles y Autoriza Descuentos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3840
            TabIndex        =   73
            Top             =   3000
            Width           =   3255
            Begin MSDataListLib.DataCombo dtcEncargado 
               Bindings        =   "frmMantenedorParametros.frx":03F8
               Height          =   315
               Left            =   1080
               TabIndex        =   75
               Top             =   240
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin VB.Label Label12 
               Caption         =   "Encargado"
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
               TabIndex        =   74
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Reserva de Repuestos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   67
            Top             =   1800
            Width           =   3495
            Begin VB.TextBox txtMailRepuestosFallidos 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   94
               Top             =   840
               Width           =   3015
            End
            Begin VB.CheckBox chkEnviaMail 
               Appearance      =   0  'Flat
               Caption         =   "Envia Mail a Bodega"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   2535
            End
            Begin VB.Label Label2 
               Caption         =   "Mail Repuestos Venta Fallida"
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   600
               Width           =   2055
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Tipo de Impresión de Recepción"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3840
            TabIndex        =   64
            Top             =   2040
            Width           =   3255
            Begin VB.OptionButton optEnBlanco 
               Appearance      =   0  'Flat
               Caption         =   "En Blanco"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1680
               TabIndex        =   66
               Top             =   360
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optPreImpreso 
               Appearance      =   0  'Flat
               Caption         =   "Pre Impreso"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Productividad Mecanico"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   3495
            Begin VB.OptionButton optLiqyFact 
               Appearance      =   0  'Flat
               Caption         =   "Ambas"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2280
               TabIndex        =   76
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton optFacturado 
               Appearance      =   0  'Flat
               Caption         =   "Facturado"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1200
               TabIndex        =   63
               Top             =   360
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optLiquidado 
               Appearance      =   0  'Flat
               Caption         =   "Liquidado"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   62
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Valores Por Defecto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1785
            Index           =   5
            Left            =   3810
            TabIndex        =   27
            Top             =   195
            Width           =   3255
            Begin MSDataListLib.DataCombo dtcGarantia 
               Bindings        =   "frmMantenedorParametros.frx":0412
               Height          =   315
               Left            =   1080
               TabIndex        =   69
               Top             =   360
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "NOMBRE"
               BoundColumn     =   "CODIGO"
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc datGarantia 
               Height          =   330
               Left            =   1170
               Top             =   360
               Visible         =   0   'False
               Width           =   1980
               _ExtentX        =   3493
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
            Begin MSDataListLib.DataCombo dtcCargo 
               Bindings        =   "frmMantenedorParametros.frx":042C
               Height          =   315
               Left            =   1080
               TabIndex        =   70
               Top             =   840
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "NOMBRE"
               BoundColumn     =   "CODIGO"
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc datCargo 
               Height          =   330
               Left            =   1200
               Top             =   840
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
            Begin MSDataListLib.DataCombo dtcMecanico 
               Bindings        =   "frmMantenedorParametros.frx":0443
               Height          =   315
               Left            =   1080
               TabIndex        =   71
               Top             =   1320
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "Nombre"
               BoundColumn     =   "Codigo"
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc datMecanico 
               Height          =   330
               Left            =   1200
               Top             =   1320
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
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               Caption         =   "Mecanico"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label31 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000016&
               Caption         =   "Tipo OT"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   10
               Left            =   75
               TabIndex        =   29
               Top             =   480
               Width           =   660
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Cargo"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   28
               Top             =   960
               Width           =   945
            End
         End
         Begin VB.Frame Frame11 
            Appearance      =   0  'Flat
            Caption         =   "Valores Carrocería"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   195
            Width           =   3495
            Begin VB.TextBox txtPresupuestoDyP 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1605
               TabIndex        =   4
               Text            =   "1"
               Top             =   225
               Width           =   1500
            End
            Begin VB.TextBox txtOrdenCarroceria 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1605
               TabIndex        =   3
               Text            =   "1"
               Top             =   540
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Nº Presupuesto"
               Height          =   195
               Index           =   2
               Left            =   75
               TabIndex        =   6
               Top             =   270
               Width           =   1110
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº Orden de Trabajo"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   5
               Top             =   555
               Visible         =   0   'False
               Width           =   1470
            End
         End
         Begin VB.Label Label18 
            Caption         =   "Código Cargo Gtia. Fábrica"
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
            Left            =   3840
            TabIndex        =   127
            Top             =   5640
            Width           =   2535
         End
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Nueva Sucursal"
            ImageKey        =   "Crear"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registros"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   0
      Top             =   480
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
            Picture         =   "frmMantenedorParametros.frx":045D
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":056F
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":0681
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":0793
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":08A5
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":09B7
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":0AC9
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":0BDB
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":0CED
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":0DFF
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":0F11
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":1023
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":1135
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":1247
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":1359
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":146B
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":157D
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":19CF
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":1E21
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":1F33
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":208F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":21EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":2347
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":24A3
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":2F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":33C3
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":3527
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":3983
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":3ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":4DEB
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":5387
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":54E3
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":563F
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":5993
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":5CE7
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":603B
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":638F
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":66E3
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":6C27
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":6D39
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":708D
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":73E1
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":74F5
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":7849
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":795B
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorParametros.frx":7CAF
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantenedorParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset
Dim mblnSW As Boolean
Dim mstrSql As String
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean


Private Sub cmdMarcas_Click()
frmPreciosporMarca.Show vbModal
End Sub

Private Sub chkPrecioMarca_Click()
    If chkPrecioMarca.Value = 1 Then
        cmdMarcas.Enabled = True
    Else
        cmdMarcas.Enabled = False
    End If
End Sub

Private Sub dtcSucursal_Change()
txtDireccion = Retorna_Valor_General("select direccion as parametro from Glbl_Sucursal Where Id_Sucursal='" & Me.dtcSucursal.BoundText & "'", gcdynamic)
End Sub

Private Sub Form_Activate()

    If Not Atributos("Glbl", "Tllr_60_0010", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If
    
    mstrSql = "SELECT * FROM Tllr_Parametro WHERE Id_Empresa='" & gstrIdEmpresa & "' AND Id_Sucursal='" & gstrIdSucursal & "'"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPrincipal
            If Not .BOF And Not .EOF Then
                'txtPresupuestoM = IIf(IsNull(!NroPreMec), 1, !NroPreMec)
                'txtOrdenMecanica = IIf(IsNull(!NroOtMec), 1, !NroOtMec)
                txtPresupuestoDyP = IIf(IsNull(!NroPreCar), 1, !NroPreCar)
                txtOrdenCarroceria = IIf(IsNull(!NroOtCar), 1, !NroOtCar)
                dtcGarantia.BoundText = IIf(IsNull(!IdTipoOtDefecto), "NGN", !IdTipoOtDefecto)
                dtcCargo.BoundText = IIf(IsNull(!IdCargoDefecto), "", !IdCargoDefecto)
                dtcMecanico.BoundText = IIf(IsNull(!MecanicoDefectoSecMec), "", !MecanicoDefectoSecMec)
                dtcEncargado.BoundText = IIf(IsNull(!MecanicoDiasHabiles), "", !MecanicoDiasHabiles)
                If Not IsNull(!Estadoprodmecanico) Then
                    If !Estadoprodmecanico = "L" Then
                        optLiquidado.Value = True
                    ElseIf !Estadoprodmecanico = "F" Then
                        optFacturado.Value = True
                    Else
                        optLiqyFact = True
                    End If
                Else
                    optLiquidado.Value = True
                End If
                If Not IsNull(!TipoImpresion) Then
                    If !TipoImpresion = "S" Then
                        optEnBlanco.Value = True
                    Else
                        optPreImpreso.Value = True
                        If InStr(gstrEmpresa, "PIAMONTE") = 1 Then
                            optPreImpreso.Tag = "P"
                        Else
                            optPreImpreso.Tag = "C"
                        End If
                    End If
                Else
                    optEnBlanco.Value = True
                End If
                chkEnviaMail.Value = IIf(IsNull(!EnviaMailBodega), 0, IIf(!EnviaMailBodega = "S", 1, 0))
                chkImprimeImagen.Value = IIf(IsNull(!ImprimeImagen), 0, IIf(!ImprimeImagen = "S", 1, 0))
                chkValidaCostoRepuestos.Value = IIf(IsNull(!ValidaCostoRepuestos), 0, IIf(!ValidaCostoRepuestos = "S", 1, 0))
                chkTraspasaRepuestos.Value = IIf(IsNull(!TraspasaRepuestos), 0, IIf(!TraspasaRepuestos = "S", 1, 0))
                txtCostoManoObra = IIf(IsNull(!VALOR_MANO_COSTO), 0, FormatoValor(!VALOR_MANO_COSTO, gstrMonedaLocal, gintDecimalesMoneda))
                txtPrecioManoObra = IIf(IsNull(!PrecioManoObra), 0, FormatoValor(!PrecioManoObra, gstrMonedaLocal, gintDecimalesMoneda))
                txtCostoInsumos = IIf(IsNull(!Insumo), 0, FormatoValor(!Insumo, gstrMonedaLocal, gintDecimalesMoneda))
                txtSeguroTaller = IIf(IsNull(!SeguroTaller), 0, FormatoValor(!SeguroTaller, gstrMonedaLocal, gintDecimalesMoneda))
                txtMargenInsumos = IIf(IsNull(!margeninsumos), 0, !margeninsumos)
                txtMargenLubricantes = IIf(IsNull(!MargenLubricantes), 0, !MargenLubricantes)
                txtMargenRepuestos = IIf(IsNull(!MargenRepuestos), 0, !MargenRepuestos)
                txtMargenMateriales = IIf(IsNull(!MargenMateriales), 0, !MargenMateriales)
                txtRegistrosDefecto = IIf(IsNull(!NroRecDefectoQry), 25, !NroRecDefectoQry)
                txtHorasTrabajo = IIf(IsNull(!NroHorasTrabajo), 8, !NroHorasTrabajo)
                txtValorExistencia = IIf(IsNull(!Valor_Existencia), 2000000, FormatoValor(!Valor_Existencia, gstrMonedaLocal, gintDecimalesMoneda))
                txtLineasRecepcion = IIf(IsNull(!LineasRecepcion), 9, !LineasRecepcion)
                txtManoObraGtia = IIf(IsNull(!PrecioManoObraGarantia), 0, FormatoValor(ValorNulo(!PrecioManoObraGarantia), gstrMonedaLocal, gintDecimalesMoneda))
                txtHoraInicio = IIf(IsNull(!HoraInicio), 8, !HoraInicio)
                txtHoraTermino = IIf(IsNull(!HoraTermino), 20, !HoraTermino)
                txtMinutos = IIf(IsNull(!IntervaloMinutos), 30, !IntervaloMinutos)
                txtDescuentoMaximo = IIf(IsNull(!DescuentoMaximo), 15, !DescuentoMaximo)
                txtDscMaxCIA = IIf(IsNull(!DsctMaxCiaSeg), 10, !DsctMaxCiaSeg)
                txtNotaPresupuesto = IIf(IsNull(!NotaPresupuesto), "", !NotaPresupuesto)
                txtNotaRecepcion = IIf(IsNull(!NotaRecepcion), "", !NotaRecepcion)
                txtCostoInsumosPorc = IIf(IsNull(!CostoInsumosPorc), "0", !CostoInsumosPorc)
                txtCostoInsumosPesos = IIf(IsNull(!CostoInsumosPesos), "0", FormatoValor(ValorNulo(!CostoInsumosPesos), gstrMonedaLocal, gintDecimalesMoneda))
                optCostoInsumosPorc.Value = IIf(SacarFormatoValor(txtCostoInsumosPesos, gstrMonedaLocal) = "0", True, False)
                optCostoInsumosPesos.Value = IIf(txtCostoInsumosPorc = "0", True, False)
                txtInsumosMO = IIf(IsNull(!MaterialesMO), "0", !MaterialesMO)
                optInsumos.Value = IIf(txtInsumosMO = "0", True, False)
                optInsumosMO.Value = IIf(SacarFormatoValor(txtCostoInsumos, gstrMonedaLocal) = "0", True, False)
                txtMailRepuestosFallidos = IIf(!MailRepuestosFallidos = "", "", ValorNulo(!MailRepuestosFallidos))
                txtCodigoLubricantes = IIf(IsNull(!CodFamiliaLubricantes), "0", !CodFamiliaLubricantes)
                txtCodigoMateriales = IIf(IsNull(!CodFamiliaMateriales), "0", !CodFamiliaMateriales)
                txtCodigoInsumos = IIf(IsNull(!CodFamiliaInsumos), "0", !CodFamiliaInsumos)
                txtMonedaLocal = IIf(IsNull(!Id_Moneda_Local), gstrMonedaLocal, !Id_Moneda_Local)
                txtDecimalesMoneda = IIf(IsNull(!DecimalesMoneda), 0, !DecimalesMoneda)
                chkPrecioMarca = IIf(IsNull(!PreciosMarca), 0, IIf(!PreciosMarca = "S", 1, 0))
                '//LREYES
                chkServiciosGenerales = IIf(IsNull(!ServiciosMarca), vbUnchecked, IIf(!ServiciosMarca = "S", vbChecked, vbUnchecked))
                
                chkBloqueaSubtotalRep = IIf(IsNull(!BloqueaSubtotalRep), vbUnchecked, IIf(!BloqueaSubtotalRep = "S", vbChecked, vbUnchecked))
                chkValidaServiciosCero = IIf(IsNull(!ValidaServiciosCero), vbUnchecked, IIf(!ValidaServiciosCero = "S", vbChecked, vbUnchecked))
                cmdMarcas.Enabled = IIf(!PreciosMarca = "S", True, False)
                txtDeducibleMas = IIf(IsNull(!CargoDeducibleMas), "", !CargoDeducibleMas)
                txtDeducibleMenos = IIf(IsNull(!CargoDeducibleMenos), "", !CargoDeducibleMenos)
                txtValoIva = IIf(IsNull(!IVA), 18, !IVA)
                chkAsignaRecursos = IIf(IsNull(!AsignaRecursos), vbUnchecked, IIf(!AsignaRecursos = "S", vbChecked, vbUnchecked))
                txtCargoGtiaFabrica = IIf(Not IsNull(!CargoGarantiaFabrica), !CargoGarantiaFabrica, "GFB")
            End If
        End With
    End If
End Sub

Private Sub Form_Load()
mblnSW = True
If gintNroRecDefectoQry = 0 Then
    tlbBarraHerramientas.Buttons.item("Crear").Enabled = True
    tlbBarraHerramientas.Buttons.item("Grabar").Enabled = False
Else
    tlbBarraHerramientas.Buttons.item("Crear").Enabled = False
End If
FillGarantia dtcGarantia, datGarantia, False
FillTipoCargo dtcCargo, datCargo
FillMecanicos dtcMecanico, datMecanico
FillMecanicos dtcEncargado, datMecanico


'mstrSql = "SELECT Top 1 Glbl_Empresa.Id_Empresa, Glbl_Empresa.Razon_Social, Glbl_Comuna.Descripcion AS COMUNA, Glbl_Ciudad.Descripcion AS CIUDAD, Glbl_Pais.Descripcion AS PAIS, Glbl_Empresa.Direccion, Glbl_Empresa.Telefono, Glbl_Empresa.Fax, Glbl_Empresa.Web , Glbl_Empresa.Mail"
'mstrSql = mstrSql & " FROM Glbl_Ciudad LEFT OUTER JOIN Glbl_Pais ON Glbl_Ciudad.Id_Pais = Glbl_Pais.Id_Pais RIGHT OUTER JOIN Glbl_Comuna ON Glbl_Ciudad.Id_Ciudad = Glbl_Comuna.Id_Ciudad AND Glbl_Ciudad.Id_Pais = Glbl_Comuna.Id_Pais RIGHT OUTER JOIN Glbl_Empresa ON Glbl_Comuna.Id_Comuna = Glbl_Empresa.Id_Comuna AND Glbl_Comuna.Id_Ciudad = Glbl_Empresa.Id_Ciudad AND Glbl_Comuna.Id_Pais = Glbl_Empresa.Id_Pais"

mstrSql = "SELECT Glbl_Empresa.Id_Empresa, Glbl_Empresa.Razon_Social, "
mstrSql = mstrSql & "Glbl_Sucursal.Descripcion, Glbl_Sucursal.Direccion "
mstrSql = mstrSql & "FROM Glbl_Empresa INNER JOIN "
mstrSql = mstrSql & "Glbl_Sucursal ON "
mstrSql = mstrSql & "Glbl_Empresa.Id_Empresa = Glbl_Sucursal.Id_Empresa "
mstrSql = mstrSql & "Where Glbl_Empresa.Id_Empresa = '" & gstrIdEmpresa & "' And Glbl_Sucursal.Id_Sucursal = '" & gstrIdSucursal & "'"

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            txtidEmpresa = !Id_Empresa
            txtRazonSocial = !Razon_Social
            txtSucursal = !Descripcion
            txtDireccion = !Direccion
        End If
    End With
End If
Label15.Caption = "Valor " & gstrNombreIva
Me.stbParametros.TabVisible(0) = False
End Sub

Private Sub optCostoInsumosPesos_Click()
If Me.optCostoInsumosPesos.Value = True Then
    Me.txtCostoInsumosPorc = "0"
    Me.txtCostoInsumosPorc.Locked = True
    Me.txtCostoInsumosPesos.Locked = False
    Me.txtCostoInsumosPesos.SetFocus
End If
End Sub

Private Sub optCostoInsumosPorc_Click()
If Me.optCostoInsumosPorc.Value = True Then
    Me.txtCostoInsumosPesos = "0"
    Me.txtCostoInsumosPesos.Locked = True
    Me.txtCostoInsumosPorc.Locked = False
    Me.txtCostoInsumosPorc.SetFocus
End If
End Sub

Private Sub optInsumos_Click()
If Me.optInsumos.Value = True Then
    Me.txtInsumosMO = "0"
    Me.txtInsumosMO.Locked = True
    Me.txtCostoInsumos.Locked = False
    Me.txtCostoInsumos.SetFocus
End If
End Sub

Private Sub optInsumosMO_Click()
If Me.optInsumosMO.Value = True Then
    Me.txtCostoInsumos = "0"
    Me.txtCostoInsumos.Locked = True
    Me.txtInsumosMO.Locked = False
    Me.txtInsumosMO.SetFocus
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Crear"
    NuevaSucursal
Case "Grabar"
    GrabarParametrosTaller
Case "Cerrar"
    Unload Me
End Select

End Sub

Private Sub GrabarParametrosTaller()
    If Not Validacion() Then
        Exit Sub
    End If

    If Me.Tag = "Crear" Then
        mstrSql = "INSERT INTO Tllr_Parametro "
        mstrSql = mstrSql & "(Id_Empresa,Id_Sucursal,Id,"
        mstrSql = mstrSql & "NroPreMec,NroOtMec,NroPreCar,NroOtCar,IdTipoOtDefecto,"
        mstrSql = mstrSql & "IdCargoDefecto,MecanicoDefectoSecMec,EstadoProdMecanico,"
        mstrSql = mstrSql & "TipoImpresion,EnviaMailBodega,Valor_Mano_costo,PrecioManoObra,"
        mstrSql = mstrSql & "Insumo,SeguroTaller,MargenInsumos,MargenLubricantes,"
        mstrSql = mstrSql & "MargenRepuestos,MargenMateriales,"
        mstrSql = mstrSql & "NroRecDefectoQry,NroHorasTrabajo,"
        mstrSql = mstrSql & "Valor_Existencia,LineasRecepcion,PrecioManoObraGarantia,"
        mstrSql = mstrSql & "HoraInicio,HoraTermino,IntervaloMinutos,MecanicoDiasHabiles,"
        mstrSql = mstrSql & "TraspasaRepuestos,DescuentoMaximo,NotaRecepcion,NotaPresupuesto,"
'kjcv 13.03.17
'        mstrSql = mstrSql & "TraspasaRepuestos,DescuentoMaximo,DsctMaxCiaSeg,NotaRecepcion,NotaPresupuesto,"
        mstrSql = mstrSql & "CostoInsumosPorc,CostoInsumosPesos,"
        mstrSql = mstrSql & "MaterialesMO,MailRepuestosFallidos,"
        mstrSql = mstrSql & "CodFamiliaLubricantes,CodFamiliaMateriales,CodFamiliaInsumos,"
        mstrSql = mstrSql & "ImprimeImagen,ValidaCostoRepuestos,Id_Moneda_Local,DecimalesMoneda,"
        mstrSql = mstrSql & "PreciosMarca, ServiciosMarca,BloqueaSubtotalRep,"
        mstrSql = mstrSql & "Iva, CorrelativoOtrosServicios,CorrelativoTrabajoTercero,"
        mstrSql = mstrSql & "MecanicoDefectoSecCar, MecanicoDefectoSecDes,MecanicoDefectoSecPin,"
        mstrSql = mstrSql & "CargoDeducibleMas,CargoDeducibleMenos,ValidaServiciosCero,AsignaRecursos,CargoGarantiaFabrica)"
        mstrSql = mstrSql & " Values ("
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "',"
        mstrSql = mstrSql & "'" & gstrIdSucursal & "',"
        mstrSql = mstrSql & "'" & 1 & "',"
        mstrSql = mstrSql & "'" & 1 & "',"
        mstrSql = mstrSql & "'" & 1 & "',"
        mstrSql = mstrSql & txtPresupuestoDyP & ","
        mstrSql = mstrSql & txtOrdenCarroceria & ","
        mstrSql = mstrSql & "'" & dtcGarantia.BoundText & "',"
        mstrSql = mstrSql & "'" & dtcCargo.BoundText & "',"
        mstrSql = mstrSql & "'" & dtcMecanico.BoundText & "',"
        mstrSql = mstrSql & "'" & IIf(optFacturado.Value = True, "F", "L") & "',"
        mstrSql = mstrSql & "'" & IIf(optPreImpreso.Value = True, optPreImpreso.Tag, "S") & "',"
        mstrSql = mstrSql & "'" & IIf(chkEnviaMail.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & SacarFormatoValor(txtCostoManoObra, gstrMonedaLocal) & ","
        mstrSql = mstrSql & SacarFormatoValor(txtPrecioManoObra, gstrMonedaLocal) & ","
        mstrSql = mstrSql & SacarFormatoValor(txtCostoInsumos, gstrMonedaLocal) & ","
        mstrSql = mstrSql & SacarFormatoValor(txtSeguroTaller, gstrMonedaLocal) & ","
        mstrSql = mstrSql & txtMargenInsumos & ","
        mstrSql = mstrSql & txtMargenLubricantes & ","
        mstrSql = mstrSql & txtMargenRepuestos & ","
        mstrSql = mstrSql & txtMargenMateriales & ","
        mstrSql = mstrSql & txtRegistrosDefecto & ","
        mstrSql = mstrSql & txtHorasTrabajo & ","
        mstrSql = mstrSql & SacarFormatoValor(txtValorExistencia, gstrMonedaLocal) & ","
        mstrSql = mstrSql & txtLineasRecepcion & ","
        mstrSql = mstrSql & SacarFormatoValor(txtManoObraGtia, gstrMonedaLocal) & ","
        mstrSql = mstrSql & txtHoraInicio & ","
        mstrSql = mstrSql & txtHoraTermino & ","
        mstrSql = mstrSql & txtMinutos & ","
        mstrSql = mstrSql & "'" & dtcEncargado.BoundText & "',"
        mstrSql = mstrSql & "'" & IIf(chkTraspasaRepuestos.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & "'" & txtDescuentoMaximo & "',"
        'kjcv 03.03.17
'        mstrSql = mstrSql & "'" & txtDscMaxCIA & "',"
        mstrSql = mstrSql & "'" & txtNotaRecepcion & "',"
        mstrSql = mstrSql & "'" & txtNotaPresupuesto & "',"
        mstrSql = mstrSql & txtCostoInsumosPorc & "," & SacarFormatoValor(txtCostoInsumosPesos, gstrMonedaLocal) & ","
        mstrSql = mstrSql & txtInsumosMO & ",'"
        mstrSql = mstrSql & txtMailRepuestosFallidos & "','"
        mstrSql = mstrSql & txtCodigoLubricantes & "','"
        mstrSql = mstrSql & txtCodigoMateriales & "','"
        mstrSql = mstrSql & txtCodigoInsumos & "','" & IIf(Me.chkImprimeImagen.Value = 1, "S", "N") & "','"
        mstrSql = mstrSql & IIf(Me.chkValidaCostoRepuestos.Value = 1, "S", "N") & "','"
        mstrSql = mstrSql & txtMonedaLocal & "'," & txtDecimalesMoneda & ",'"
        mstrSql = mstrSql & IIf(Me.chkPrecioMarca.Value = 1, "S", "N") & "', '" & IIf(Me.chkServiciosGenerales.Value = vbChecked, "S", "N") & "','"
        mstrSql = mstrSql & IIf(Me.chkBloqueaSubtotalRep.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & Me.txtValoIva & "," & "1,1,"
        mstrSql = mstrSql & "'" & dtcMecanico.BoundText & "'," 'sec car
        mstrSql = mstrSql & "'" & dtcMecanico.BoundText & "'," 'sec des
        mstrSql = mstrSql & "'" & dtcMecanico.BoundText & "'," 'sec pin
        mstrSql = mstrSql & "'" & Me.txtDeducibleMas & "','" & Me.txtDeducibleMenos & "',"
        mstrSql = mstrSql & "'" & IIf(Me.chkValidaServiciosCero.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & "'" & IIf(Me.chkAsignaRecursos.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & "'" & txtCargoGtiaFabrica & "')"
    Else
        mstrSql = "UPDATE Tllr_Parametro SET "
        mstrSql = mstrSql & " NroPreMec = " & "1" & ","
        mstrSql = mstrSql & " NroOtMec = " & "1" & ","
        mstrSql = mstrSql & " NroPreCar = " & txtPresupuestoDyP & ","
        mstrSql = mstrSql & " NroOtCar = " & txtOrdenCarroceria & ","
        mstrSql = mstrSql & " IdTipoOtDefecto = '" & dtcGarantia.BoundText & "',"
        mstrSql = mstrSql & " IdCargoDefecto = '" & dtcCargo.BoundText & "',"
        mstrSql = mstrSql & " MecanicoDefectoSecMec = '" & dtcMecanico.BoundText & "',"
        mstrSql = mstrSql & " EstadoProdMecanico = '" & IIf(optFacturado.Value = True, "F", IIf(Me.optLiquidado.Value = True, "L", "A")) & "',"
        mstrSql = mstrSql & " TipoImpresion = '" & IIf(optPreImpreso.Value = True, optPreImpreso.Tag, "S") & "',"
        mstrSql = mstrSql & " EnviaMailBodega = '" & IIf(chkEnviaMail.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & " Valor_Mano_Costo = " & SacarFormatoValor(txtCostoManoObra, gstrMonedaLocal) & ","
        mstrSql = mstrSql & " PrecioManoObra = " & SacarFormatoValor(txtPrecioManoObra, gstrMonedaLocal) & ","
        mstrSql = mstrSql & " Insumo = " & SacarFormatoValor(txtCostoInsumos, gstrMonedaLocal) & ","
        mstrSql = mstrSql & " SeguroTaller = " & SacarFormatoValor(txtSeguroTaller, gstrMonedaLocal) & ","
        mstrSql = mstrSql & " Iva = " & Me.txtValoIva & ","
        mstrSql = mstrSql & " MargenInsumos = " & txtMargenInsumos & ","
        mstrSql = mstrSql & " MargenLubricantes = " & txtMargenLubricantes & ","
        mstrSql = mstrSql & " MargenRepuestos = " & txtMargenRepuestos & ","
        mstrSql = mstrSql & " MargenMateriales = " & txtMargenMateriales & ","
        mstrSql = mstrSql & " NroRecDefectoQry = " & txtRegistrosDefecto & ","
        mstrSql = mstrSql & " NroHorasTrabajo = " & txtHorasTrabajo & ","
        mstrSql = mstrSql & " Valor_Existencia = " & SacarFormatoValor(txtValorExistencia, gstrMonedaLocal) & ","
        mstrSql = mstrSql & " LineasRecepcion = " & txtLineasRecepcion & ","
        mstrSql = mstrSql & " PrecioManoObraGarantia = " & SacarFormatoValor(txtManoObraGtia, gstrMonedaLocal) & ","
        mstrSql = mstrSql & " HoraInicio = " & txtHoraInicio & ","
        mstrSql = mstrSql & " HoraTermino = " & txtHoraTermino & ","
        mstrSql = mstrSql & " IntervaloMinutos = " & txtMinutos & ","
        mstrSql = mstrSql & " MecanicoDiasHabiles = '" & dtcEncargado.BoundText & "',"
        mstrSql = mstrSql & " TraspasaRepuestos = '" & IIf(chkTraspasaRepuestos.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & " DescuentoMaximo = '" & txtDescuentoMaximo & "',"
        'kjcv 13.03.17
'        mstrSql = mstrSql & " DsctMaxCiaSeg = '" & txtDscMaxCIA & "',"
        mstrSql = mstrSql & " NotaRecepcion = '" & txtNotaRecepcion & "',"
        mstrSql = mstrSql & " NotaPresupuesto = '" & txtNotaPresupuesto & "',"
        mstrSql = mstrSql & " CostoInsumosPorc = " & txtCostoInsumosPorc & ","
        mstrSql = mstrSql & " CostoInsumosPesos = " & SacarFormatoValor(txtCostoInsumosPesos, gstrMonedaLocal) & ","
        mstrSql = mstrSql & " MaterialesMO = " & txtInsumosMO & ","
        mstrSql = mstrSql & " MailRepuestosFallidos = '" & txtMailRepuestosFallidos & "',"
        mstrSql = mstrSql & " CodFamiliaLubricantes = '" & txtCodigoLubricantes & "',"
        mstrSql = mstrSql & " CodFamiliaMateriales = '" & txtCodigoMateriales & "',"
        mstrSql = mstrSql & " CodFamiliaInsumos = '" & txtCodigoInsumos & "',"
        mstrSql = mstrSql & " ImprimeImagen = '" & IIf(Me.chkImprimeImagen.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & " ValidaCostoRepuestos = '" & IIf(Me.chkValidaCostoRepuestos.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & " Id_Moneda_Local = '" & txtMonedaLocal & "',"
        mstrSql = mstrSql & " DecimalesMoneda = '" & txtDecimalesMoneda & "',"
        mstrSql = mstrSql & " PreciosMarca = '" & IIf(Me.chkPrecioMarca.Value = 1, "S", "N") & "',"
        mstrSql = mstrSql & " ServiciosMarca = '" & IIf(Me.chkServiciosGenerales.Value = vbChecked, "S", "N") & "',"
        mstrSql = mstrSql & " BloqueaSubtotalRep = '" & IIf(Me.chkBloqueaSubtotalRep.Value = vbChecked, "S", "N") & "',"
        mstrSql = mstrSql & " ValidaServiciosCero = '" & IIf(Me.chkValidaServiciosCero.Value = vbChecked, "S", "N") & "',"
        mstrSql = mstrSql & " CargoDeducibleMas = '" & txtDeducibleMas & "',"
        mstrSql = mstrSql & " CargoDeducibleMenos = '" & txtDeducibleMenos & "',"
        mstrSql = mstrSql & " AsignaRecursos = '" & IIf(Me.chkAsignaRecursos.Value = vbChecked, "S", "N") & "',"
        mstrSql = mstrSql & " CargoGarantiaFabrica = '" & txtCargoGtiaFabrica & "'"
        mstrSql = mstrSql & " Where Id_Empresa = '" & gstrIdEmpresa & "' And Id_Sucursal = '" & gstrIdSucursal & "'"
    End If

    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        'mblnTablaVacia = False
        'ActivaBotones
        Me.Tag = ""
        Me.dtcSucursal.Visible = False
        Me.txtSucursal.Visible = True
        If ParametrosDefecto(gstrIdEmpresa, gstrIdSucursal) = False Then
            MsgBox LoadResString(101), vbCritical + vbOKOnly, "ElisaTaller"
        End If
    End If


End Sub

Sub LimpiaCampos()
With Me
    .txtDireccion = ""
    .txtPresupuestoDyP = "1"
    .txtOrdenCarroceria = "1"
    .dtcGarantia.BoundText = gstrIdTipoOtDefecto
    .dtcCargo.BoundText = gstrIdCargoDefecto
    .dtcMecanico = gstrMecanicoDefectoSecMec
    .optEnBlanco.Value = True
    .optLiquidado.Value = True
    .chkEnviaMail.Value = 0
    .txtCostoManoObra = 0
    .txtPrecioManoObra = 0
    .txtCostoInsumos = 0
    .txtSeguroTaller = 0
    .txtMargenInsumos = 0
    .txtMargenLubricantes = 0
    .txtMargenRepuestos = 0
    .txtMargenMateriales = 0
    .txtRegistrosDefecto = 25
    .txtHorasTrabajo = 8
    .txtValorExistencia = 2000000
    .txtLineasRecepcion = 9
    .txtManoObraGtia = 0
    .txtHoraInicio = 8
    .txtHoraTermino = 20
    .txtMinutos = 30
    .txtDescuentoMaximo = 15
    'kjcv 13.03.17
'    .txtDscMaxCIA = gintDescuentoMaximoCIA
    .txtNotaPresupuesto = ""
    .txtNotaRecepcion = ""
    .txtSucursal.Visible = False
    .dtcSucursal.Visible = True
    .txtMailRepuestosFallidos = ""
    .txtCodigoInsumos = "0"
    .txtCodigoLubricantes = "0"
    .txtCodigoMateriales = "0"
    .chkImprimeImagen.Value = 0
    .chkTraspasaRepuestos.Value = 0
    .txtDecimalesMoneda = 0
    .txtMonedaLocal = gstrMonedaLocal
    .chkBloqueaSubtotalRep.Value = 0
    .chkAsignaRecursos.Value = 0
    .txtCargoGtiaFabrica = ""
End With
End Sub

Private Sub txtCodigoInsumos_GotFocus()
MarcaTexto txtCodigoInsumos
End Sub

Private Sub txtCodigoLubricantes_GotFocus()
MarcaTexto txtCodigoLubricantes
End Sub

Private Sub txtCodigoMateriales_GotFocus()
MarcaTexto txtCodigoMateriales
End Sub

Private Sub txtCostoInsumos_GotFocus()
txtCostoInsumos = SacarFormatoValor(txtCostoInsumos, gstrMonedaLocal)
MarcaTexto txtCostoInsumos
End Sub

Private Sub txtCostoInsumos_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCostoInsumos, strDot)
End Sub

Private Sub txtCostoInsumos_LostFocus()
txtCostoInsumos = FormatoValor(txtCostoInsumos, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtCostoInsumosPesos_GotFocus()
txtCostoInsumosPesos = SacarFormatoValor(txtCostoInsumosPesos, gstrMonedaLocal)
MarcaTexto txtCostoInsumosPesos
End Sub

Private Sub txtCostoInsumosPesos_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCostoInsumosPesos, strDot)
End Sub

Private Sub txtCostoInsumosPesos_LostFocus()
txtCostoInsumosPesos = FormatoValor(txtCostoInsumosPesos, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtCostoInsumosPorc_GotFocus()
MarcaTexto txtCostoInsumosPorc
End Sub

Private Sub txtCostoInsumosPorc_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCostoInsumosPorc, strDot)
End Sub

Private Sub txtCostoManoObra_GotFocus()
txtCostoManoObra = SacarFormatoValor(txtCostoManoObra, gstrMonedaLocal)
MarcaTexto txtCostoManoObra
End Sub

Private Sub txtCostoManoObra_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCostoManoObra, strDot)
End Sub

Private Sub txtCostoManoObra_LostFocus()
txtCostoManoObra = FormatoValor(txtCostoManoObra, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtDecimalesMoneda_GotFocus()
MarcaTexto txtDecimalesMoneda
End Sub

Private Sub txtDecimalesMoneda_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtDecimalesMoneda, strDot)
End Sub

Private Sub txtDeducibleMas_GotFocus()
MarcaTexto txtDeducibleMas
End Sub

Private Sub txtDeducibleMas_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtDeducibleMenos_GotFocus()
MarcaTexto txtDeducibleMenos
End Sub

Private Sub txtDeducibleMenos_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtDescuentoMaximo_GotFocus()
MarcaTexto txtDescuentoMaximo
End Sub

Private Sub txtDescuentoMaximo_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtDescuentoMaximo, strDot)
End Sub

Private Sub txtHoraInicio_GotFocus()
MarcaTexto txtHoraInicio
End Sub

Private Sub txtHoraInicio_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHoraInicio, strDot)
End Sub

Private Sub txtHorasTrabajo_GotFocus()
MarcaTexto txtHorasTrabajo
End Sub

Private Sub txtHorasTrabajo_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasTrabajo, strDot)
End Sub

Private Sub txtHoraTermino_GotFocus()
MarcaTexto txtHoraTermino
End Sub

Private Sub txtHoraTermino_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHoraTermino, strDot)
End Sub

Private Sub txtInsumosMO_GotFocus()
txtInsumosMO = SacarFormatoValor(txtInsumosMO, gstrMonedaLocal)
MarcaTexto txtInsumosMO
End Sub

Private Sub txtInsumosMO_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtInsumosMO, strDot)
End Sub

Private Sub txtLineasRecepcion_GotFocus()
MarcaTexto txtLineasRecepcion
End Sub

Private Sub txtLineasRecepcion_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtLineasRecepcion, strDot)
End Sub

Private Sub txtManoObraGtia_GotFocus()
txtManoObraGtia = SacarFormatoValor(txtManoObraGtia, gstrMonedaLocal)
MarcaTexto txtManoObraGtia
End Sub

Private Sub txtManoObraGtia_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtManoObraGtia, strDot)
End Sub

Private Sub txtManoObraGtia_LostFocus()
txtManoObraGtia = FormatoValor(txtManoObraGtia, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtMargenInsumos_GotFocus()
MarcaTexto txtMargenInsumos
End Sub

Private Sub txtMargenInsumos_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMargenInsumos, strDot)
End Sub

Private Sub txtMargenLubricantes_GotFocus()
MarcaTexto txtMargenLubricantes
End Sub

Private Sub txtMargenLubricantes_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMargenLubricantes, strDot)
End Sub

Private Sub txtMargenMateriales_GotFocus()
MarcaTexto txtMargenMateriales
End Sub

Private Sub txtMargenMateriales_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMargenMateriales, strDot)
End Sub

Private Sub txtMargenRepuestos_GotFocus()
MarcaTexto txtMargenRepuestos
End Sub

Private Sub txtMargenRepuestos_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMargenRepuestos, strDot)
End Sub

Private Sub txtMinutos_GotFocus()
MarcaTexto txtMinutos
End Sub

Private Sub txtMinutos_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMinutos, strDot)
End Sub

Private Sub txtMonedaLocal_GotFocus()
MarcaTexto txtMonedaLocal
End Sub

Private Sub txtNotaPresupuesto_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtNotaRecepcion_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtPrecioManoObra_GotFocus()
txtPrecioManoObra = SacarFormatoValor(txtPrecioManoObra, gstrMonedaLocal)
MarcaTexto txtPrecioManoObra
End Sub

Private Sub txtPrecioManoObra_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPrecioManoObra, strDot)
End Sub

Private Sub txtPrecioManoObra_LostFocus()
txtPrecioManoObra = FormatoValor(txtPrecioManoObra, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtRegistrosDefecto_GotFocus()
MarcaTexto txtRegistrosDefecto
End Sub

Private Sub txtRegistrosDefecto_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtRegistrosDefecto, strDot)
End Sub

Private Sub txtSeguroTaller_GotFocus()
txtSeguroTaller = SacarFormatoValor(txtSeguroTaller, gstrMonedaLocal)
MarcaTexto txtSeguroTaller
End Sub

Private Sub txtSeguroTaller_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtSeguroTaller, strDot)
End Sub

Private Sub txtSeguroTaller_LostFocus()
txtSeguroTaller = FormatoValor(txtSeguroTaller, gstrMonedaLocal, gintDecimalesMoneda)
End Sub

Private Sub txtValoIva_GotFocus()
MarcaTexto txtValoIva
End Sub

Private Sub txtValoIva_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtValoIva, strDot)
End Sub

Private Sub txtValorExistencia_GotFocus()
txtValorExistencia = SacarFormatoValor(txtValorExistencia, gstrMonedaLocal)
MarcaTexto txtValorExistencia
End Sub

Private Sub txtValorExistencia_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtValorExistencia, strDot)
End Sub

Private Sub txtValorExistencia_LostFocus()
txtValorExistencia = FormatoValor(txtValorExistencia, gstrMonedaLocal, gintDecimalesMoneda)
End Sub
Function Validacion() As Boolean
Validacion = True
With Me
    If Me.txtPresupuestoDyP = "" Then
        MsgBox "El Número de Presupuesto de Carroceria debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtPresupuestoDyP.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtOrdenCarroceria = "" Then
        MsgBox "El Número de Orden de Carrocería debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtOrdenCarroceria.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.dtcGarantia.BoundText = "" Then
        MsgBox "El Tipo de OT Debe contener un Valor", vbExclamation, "Parametros Taller"
        dtcGarantia.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.dtcCargo.BoundText = "" Then
        MsgBox "El Tipo de Cargo Debe contener un Valor", vbExclamation, "Parametros Taller"
        dtcCargo.SetFocus
        Validacion = False
        Exit Function
    End If
    If dtcMecanico.BoundText = "" Then
        MsgBox "El Mecanico por Defecto Debe contener un Valor", vbExclamation, "Parametros Taller"
        dtcMecanico.SetFocus
        Validacion = False
        Exit Function
    End If
    If dtcEncargado.BoundText = "" Then
        MsgBox "El Encargado de Cambiar los Dias Habiles en Prod. de Mecanico Debe contener un Valor", vbExclamation, "Parametros Taller"
        dtcEncargado.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtCostoManoObra = "" Then
        MsgBox "El Costo de la Mano de Obra Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtCostoManoObra.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtPrecioManoObra = "" Then
        MsgBox "El Precio de la Mano de Obra Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtPrecioManoObra.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtCostoInsumos = "" Then
        MsgBox "La Venta de Insumos Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtCostoInsumos.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtSeguroTaller = "" Then
        MsgBox "El Seguro de Taller Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtSeguroTaller.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtRegistrosDefecto = "" Then
        MsgBox "El Número de Registros por defecto Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtRegistrosDefecto.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtHorasTrabajo = "" Then
        MsgBox "Las Horas de Trabajo Deben contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtHorasTrabajo.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtValorExistencia = "" Then
        MsgBox "El Valor de Existencia Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtValorExistencia.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtLineasRecepcion = "" Then
        MsgBox "Las Lineas de Recepcion Deben contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtLineasRecepcion.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtManoObraGtia = "" Then
        MsgBox "El Precio de la Mano de Obre de Garantía Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtManoObraGtia.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtHoraInicio = "" Then
        MsgBox "La Hora de Inicio Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtHoraInicio.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtHoraTermino = "" Then
        MsgBox "La Hora de Termino Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtHoraTermino.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtMinutos = "" Then
        MsgBox "El Intervalo de Minutos Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtMinutos.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtDescuentoMaximo = "" Then
        MsgBox "El Descuento Máximo en Repuestos Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtDescuentoMaximo.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtCostoInsumosPorc = "" Then
        MsgBox "El Costo de Insumos en Porcentaje Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtCostoInsumosPorc.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtCostoInsumosPesos = "" Then
        MsgBox "El Costo de Insumos en Pesos Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtCostoInsumosPesos.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtInsumosMO = "" Then
        MsgBox "El Porcentaje de Insumos Sobre Mano de Obra Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtInsumosMO.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtCodigoInsumos = "" Then
        MsgBox "El Código de Familia de Insumos Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtCodigoInsumos.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtCodigoLubricantes = "" Then
        MsgBox "El Código de Familia de Lubricantes Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtCodigoLubricantes.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtCodigoMateriales = "" Then
        MsgBox "El Código de Familia de Materiales Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtCodigoMateriales.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtMonedaLocal = "" Then
        MsgBox "La Moneda Local Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtMonedaLocal.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtDecimalesMoneda = "" Then
        MsgBox "Los Decimales de la Moneda Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtDecimalesMoneda.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtDeducibleMas = "" Then
        MsgBox "El Código del Cargo de Deducible (+) Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtDeducibleMas.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtDeducibleMenos = "" Then
        MsgBox "El Código del Cargo de Deducible (-) Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtDeducibleMenos.SetFocus
        Validacion = False
        Exit Function
    End If
    If Me.txtValoIva = "" Then
        MsgBox "El Valor del " & gstrNombreIva & " Debe contener un Valor", vbExclamation, "Parametros Taller"
        Me.txtValoIva.SetFocus
        Validacion = False
        Exit Function
    End If
    If .txtCargoGtiaFabrica = "" Then
        MsgBox "El Código Cargo Garantía Fábrica debe contener un Valor", vbExclamation, "Parametros Taller"
        txtCargoGtiaFabrica.SetFocus
        Validacion = False
        Exit Function
    End If
End With

End Function

Sub NuevaSucursal()
    Me.Tag = "Crear"
    tlbBarraHerramientas.Buttons.item(2).Enabled = True
    tlbBarraHerramientas.Buttons.item(1).Enabled = False
    LimpiaCampos
    LlenaSucursal
End Sub
Sub LlenaSucursal()
mstrSql = "Select Id_Sucursal as Codigo, Descripcion as Nombre From Glbl_Sucursal "
mstrSql = mstrSql & "Where Id_Sucursal Not In "
mstrSql = mstrSql & "(Select Id_Sucursal from Tllr_Parametro) "
mstrSql = mstrSql & "and id_empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datSucursal
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcSucursal.ListField = "Nombre"
            dtcSucursal.BoundColumn = "Codigo"
        End If
    End With
End If
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal

End Sub
