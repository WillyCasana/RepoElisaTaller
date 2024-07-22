VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOrdenTrabajoTaller 
   Caption         =   "Ordenes de Trabajo"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmOrdenTrabajoTaller.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Caption         =   "Sección"
      Height          =   555
      Left            =   30
      TabIndex        =   75
      Top             =   330
      Width           =   2940
      Begin VB.OptionButton optRecepcion 
         Caption         =   "Carrocería"
         Height          =   300
         Index           =   1
         Left            =   1620
         TabIndex        =   77
         Tag             =   "Carrocería"
         Top             =   195
         Width           =   1230
      End
      Begin VB.OptionButton optRecepcion 
         Caption         =   "Mecánica"
         Height          =   300
         Index           =   0
         Left            =   255
         TabIndex        =   76
         Tag             =   "Mecánica"
         Top             =   195
         Width           =   1155
      End
   End
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   2970
      TabIndex        =   15
      Top             =   330
      Width           =   10455
      Begin MSComCtl2.DTPicker pckFechaAtencion 
         Height          =   315
         Left            =   8865
         TabIndex        =   18
         Top             =   180
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36733
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Atención"
         Height          =   195
         Index           =   9
         Left            =   7650
         TabIndex        =   19
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblNroOT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1695
         TabIndex        =   17
         Top             =   195
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Trabajo Nº :"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   240
         Width           =   1560
      End
   End
   Begin TabDlg.SSTab stbServicios 
      Height          =   6135
      Left            =   30
      TabIndex        =   3
      Top             =   900
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOrdenTrabajoTaller.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Inventario Orden de Trabajo - Comentario"
      TabPicture(1)   =   "frmOrdenTrabajoTaller.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Servicios Mecánica"
      TabPicture(2)   =   "frmOrdenTrabajoTaller.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7(0)"
      Tab(2).Control(1)=   "stbTotalMecanica"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Servicios Carroceria"
      TabPicture(3)   =   "frmOrdenTrabajoTaller.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "StatusBar1"
      Tab(3).Control(1)=   "Frame7(1)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Servicios de Terceros"
      TabPicture(4)   =   "frmOrdenTrabajoTaller.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7(2)"
      Tab(4).Control(1)=   "StatusBar2"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Repuestos"
      TabPicture(5)   =   "frmOrdenTrabajoTaller.frx":04CE
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "StatusBar3"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame7(3)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin MSComctlLib.StatusBar stbTotalMecanica 
         Height          =   330
         Left            =   -71300
         TabIndex        =   124
         Top             =   6500
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   6174
               MinWidth        =   6174
               Text            =   "Total Servicios Mecánica :"
               TextSave        =   "Total Servicios Mecánica :"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame7 
         Caption         =   "Detalle Servicios Solicitados - Secciòn Terceros"
         Height          =   4800
         Index           =   2
         Left            =   -74900
         TabIndex        =   111
         Top             =   350
         Width           =   13000
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7185
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   113
            Text            =   "0"
            Top             =   405
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6180
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   112
            Top             =   405
            Width           =   1005
         End
         Begin MSComctlLib.ListView lvwServiciosTerceros 
            Height          =   4000
            Left            =   105
            TabIndex        =   114
            Top             =   705
            Width           =   11500
            _ExtentX        =   20294
            _ExtentY        =   7064
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "CODIGO"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "DESCRIPCION"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   "CANTIDAD"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "VALOR"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "SUBTOTAL"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   2
            Left            =   11805
            TabIndex        =   115
            Top             =   330
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
            ButtonWidth     =   1693
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            Height          =   195
            Index           =   45
            Left            =   8800
            TabIndex        =   123
            Top             =   210
            Width           =   690
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Unitario"
            Height          =   195
            Index           =   44
            Left            =   7300
            TabIndex        =   122
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   43
            Left            =   6350
            TabIndex        =   121
            Top             =   210
            Width           =   630
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   42
            Left            =   3900
            TabIndex        =   120
            Top             =   210
            Width           =   840
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   195
            Index           =   41
            Left            =   800
            TabIndex        =   119
            Top             =   210
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   118
            Top             =   405
            Width           =   2220
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2340
            TabIndex        =   117
            Top             =   405
            Width           =   3840
         End
         Begin VB.Label lblSubTotalLineaServicioTercero 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8385
            TabIndex        =   116
            Top             =   405
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Detalle Servicios Solicitados - Secciòn Repuestos"
         Height          =   4800
         Index           =   3
         Left            =   100
         TabIndex        =   98
         Top             =   350
         Width           =   13000
         Begin VB.TextBox txtCantidadRepuesto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7095
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   100
            Top             =   405
            Width           =   1000
         End
         Begin VB.TextBox txtValorUnitarioRepuesto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8100
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   99
            Text            =   "0"
            Top             =   405
            Width           =   1500
         End
         Begin MSComctlLib.ListView lvwRepuestos 
            Height          =   4000
            Left            =   105
            TabIndex        =   101
            Top             =   705
            Width           =   10000
            _ExtentX        =   17648
            _ExtentY        =   7064
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "CODIGOITEM"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "DESCRIPCION"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   "CANTIDAD"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "VALORUNITARIO"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "SUBTOTAL"
               Object.Width           =   3528
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   3
            Left            =   11805
            TabIndex        =   102
            Top             =   330
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
            ButtonWidth     =   1693
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lblSubTotalLineaRepuesto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9585
            TabIndex        =   110
            Top             =   405
            Width           =   1995
         End
         Begin VB.Label lblDescripcionRepuesto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2610
            TabIndex        =   109
            Top             =   405
            Width           =   4500
         End
         Begin VB.Label lblCodigoItem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   105
            TabIndex        =   108
            Top             =   405
            Width           =   2500
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código Item"
            Height          =   195
            Index           =   40
            Left            =   800
            TabIndex        =   107
            Top             =   210
            Width           =   840
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   39
            Left            =   3900
            TabIndex        =   106
            Top             =   210
            Width           =   840
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   38
            Left            =   6350
            TabIndex        =   105
            Top             =   210
            Width           =   630
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Unitario"
            Height          =   195
            Index           =   37
            Left            =   7300
            TabIndex        =   104
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            Height          =   195
            Index           =   36
            Left            =   8800
            TabIndex        =   103
            Top             =   210
            Width           =   690
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Detalle Servicios Solicitados - Secciòn Carrocería"
         Height          =   4800
         Index           =   1
         Left            =   -74900
         TabIndex        =   61
         Top             =   350
         Width           =   13000
         Begin VB.TextBox txtValorFin 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10065
            MaxLength       =   8
            TabIndex        =   66
            Text            =   "0"
            Top             =   405
            Width           =   1530
         End
         Begin VB.TextBox txtValorDef 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8700
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   65
            Text            =   "0"
            Top             =   405
            Width           =   1380
         End
         Begin VB.TextBox txtSeccion 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3165
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   64
            Top             =   405
            Width           =   500
         End
         Begin MSComctlLib.ListView lvwServiciosCarroceria 
            Height          =   4000
            Left            =   105
            TabIndex        =   63
            Top             =   705
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   7064
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "CONCEPTO"
               Text            =   "Concepto"
               Object.Width           =   5397
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "IDCONCEPTO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   "SECCION"
               Text            =   "Tipo"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "PARTEPIEZA"
               Text            =   "Parte / Pieza"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "IDPARTEPIEZA"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   "VALORDEF"
               Text            =   "Valor Definido"
               Object.Width           =   2434
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   "VALORFIN"
               Text            =   "Valor Final"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcPartePieza 
            Bindings        =   "frmOrdenTrabajoTaller.frx":04EA
            Height          =   315
            Left            =   3675
            TabIndex        =   67
            Top             =   405
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcConceptos 
            Bindings        =   "frmOrdenTrabajoTaller.frx":0508
            Height          =   315
            Left            =   120
            TabIndex        =   68
            Top             =   405
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datPartesPiezas 
            Height          =   330
            Left            =   3135
            Top             =   630
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
            Caption         =   "Adodc2"
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
         Begin MSAdodcLib.Adodc datConceptos 
            Height          =   330
            Left            =   150
            Top             =   600
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
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   1
            Left            =   11805
            TabIndex        =   79
            Top             =   330
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
            ButtonWidth     =   1693
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Final"
            Height          =   195
            Index           =   28
            Left            =   10365
            TabIndex        =   73
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Definido"
            Height          =   195
            Index           =   27
            Left            =   8880
            TabIndex        =   72
            Top             =   210
            Width           =   990
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parte / Pieza"
            Height          =   195
            Index           =   26
            Left            =   5700
            TabIndex        =   71
            Top             =   210
            Width           =   930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   25
            Left            =   3240
            TabIndex        =   70
            Top             =   210
            Width           =   315
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto"
            Height          =   195
            Index           =   24
            Left            =   1095
            TabIndex        =   69
            Top             =   210
            Width           =   690
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Detalle Servicios Solicitados - Secciòn Mecánica"
         Height          =   4300
         Index           =   0
         Left            =   -74900
         TabIndex        =   60
         Top             =   350
         Width           =   13000
         Begin MSComctlLib.ListView lvwServiciosMecanica 
            Height          =   4000
            Left            =   105
            TabIndex        =   62
            Top             =   210
            Width           =   10000
            _ExtentX        =   17648
            _ExtentY        =   7064
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
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "Codigo"
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "Des"
               Text            =   "Descripción"
               Object.Width           =   10019
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   "NroHoras"
               Text            =   "Nº Horas"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   "Valor"
               Text            =   "Valor"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbOpciones 
            Height          =   660
            Index           =   0
            Left            =   11805
            TabIndex        =   80
            Top             =   225
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
            ButtonWidth     =   1693
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  Key             =   "Agregar"
                  Object.ToolTipText     =   "Agrega Servicio Nuevo"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Quitar"
                  Key             =   "Quitar"
                  Object.ToolTipText     =   "Quitar Servicio"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Comentario"
         Height          =   5655
         Left            =   -70350
         TabIndex        =   58
         Top             =   330
         Width           =   6645
         Begin VB.TextBox txtComentario 
            Height          =   5300
            Left            =   150
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   240
            Width           =   6330
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Inventario Recepciòn"
         Height          =   5655
         Left            =   -74900
         TabIndex        =   56
         Top             =   330
         Width           =   4425
         Begin MSComctlLib.ListView lvwInventario 
            Height          =   5300
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   9340
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   7056
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4305
         Left            =   -74955
         TabIndex        =   21
         Top             =   315
         Width           =   11340
         Begin VB.TextBox txtSolicita 
            Height          =   315
            Left            =   6885
            MaxLength       =   50
            TabIndex        =   91
            Top             =   2070
            Width           =   4185
         End
         Begin VB.TextBox txtFolioGarantia 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8040
            MaxLength       =   30
            TabIndex        =   88
            Top             =   195
            Width           =   3000
         End
         Begin VB.TextBox txtPatente 
            Height          =   315
            Left            =   1215
            MaxLength       =   6
            TabIndex        =   26
            Top             =   225
            Width           =   1200
         End
         Begin VB.TextBox txtNroCono 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5355
            MaxLength       =   3
            TabIndex        =   25
            Top             =   3390
            Width           =   930
         End
         Begin VB.TextBox txtAño 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7260
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   24
            Text            =   "2000"
            Top             =   705
            Width           =   600
         End
         Begin VB.TextBox txtKilAct 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1230
            MaxLength       =   6
            TabIndex        =   23
            Top             =   1620
            Width           =   1380
         End
         Begin VB.ComboBox cboHora 
            Height          =   315
            Left            =   9150
            TabIndex        =   22
            Top             =   3885
            Width           =   1170
         End
         Begin MSComCtl2.DTPicker pckFechaEntrega 
            Height          =   315
            Left            =   5850
            TabIndex        =   27
            Top             =   3885
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   24772609
            CurrentDate     =   36733
         End
         Begin MSDataListLib.DataCombo dtcTipoCono 
            Bindings        =   "frmOrdenTrabajoTaller.frx":0523
            Height          =   315
            Left            =   1260
            TabIndex        =   28
            Top             =   3390
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datTipoCono 
            Height          =   330
            Left            =   2010
            Top             =   3375
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
            Caption         =   "Adodc2"
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
         Begin MSComctlLib.Toolbar tlbPatente 
            Height          =   330
            Left            =   2445
            TabIndex        =   29
            Top             =   240
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImgBarraHerramienta"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nuevo"
                  Object.ToolTipText     =   "Nuevo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Buscar"
                  Object.ToolTipText     =   "Buscar"
                  ImageIndex      =   9
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcRecepcionista 
            Bindings        =   "frmOrdenTrabajoTaller.frx":053D
            Height          =   315
            Left            =   1245
            TabIndex        =   30
            Top             =   3885
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Nombre"
            BoundColumn     =   "Codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datRecepcionista 
            Height          =   330
            Left            =   2550
            Top             =   3870
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
         Begin MSDataListLib.DataCombo dtcGarantia 
            Bindings        =   "frmOrdenTrabajoTaller.frx":055C
            Height          =   315
            Left            =   4365
            TabIndex        =   86
            Top             =   210
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "CODIGO"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc datGarantia 
            Height          =   330
            Left            =   5415
            Top             =   195
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
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   225
            X2              =   11070
            Y1              =   3300
            Y2              =   3300
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   225
            X2              =   11070
            Y1              =   3300
            Y2              =   3300
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "R.U.T."
            Height          =   195
            Index           =   35
            Left            =   5055
            TabIndex        =   97
            Top             =   3000
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comuna"
            Height          =   195
            Index           =   34
            Left            =   135
            TabIndex        =   96
            Top             =   2940
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
            Height          =   195
            Index           =   33
            Left            =   135
            TabIndex        =   95
            Top             =   2490
            Width           =   675
         End
         Begin VB.Label lblDireccionCliente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1245
            TabIndex        =   94
            Top             =   2490
            Width           =   6600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fonos"
            Height          =   195
            Index           =   32
            Left            =   8175
            TabIndex        =   93
            Top             =   2550
            Width           =   435
         End
         Begin VB.Label lblFono 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8820
            TabIndex        =   92
            Top             =   2520
            Width           =   2250
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Solicita "
            Height          =   195
            Index           =   31
            Left            =   6210
            TabIndex        =   90
            Top             =   2085
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folio Garantía "
            Height          =   195
            Index           =   30
            Left            =   6825
            TabIndex        =   89
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Garantía"
            Height          =   195
            Index           =   23
            Left            =   3645
            TabIndex        =   87
            Top             =   270
            Width           =   630
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VIN"
            Height          =   195
            Index           =   29
            Left            =   7215
            TabIndex        =   85
            Top             =   1170
            Width           =   270
         End
         Begin VB.Label lblVIN 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7545
            TabIndex        =   84
            Top             =   1140
            Width           =   3510
         End
         Begin VB.Label lblRutCliente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5655
            TabIndex        =   83
            Top             =   2940
            Width           =   1695
         End
         Begin VB.Label lblComunaCliente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1245
            TabIndex        =   82
            Top             =   2910
            Width           =   2820
         End
         Begin VB.Label lblFechaVenta 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7695
            TabIndex        =   78
            Top             =   1605
            Width           =   1485
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo "
            Height          =   195
            Index           =   22
            Left            =   8175
            TabIndex        =   55
            Top             =   735
            Width           =   360
         End
         Begin VB.Label lblTipoVeh 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8670
            TabIndex        =   54
            Top             =   690
            Width           =   2370
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recepcionista"
            Height          =   195
            Index           =   16
            Left            =   105
            TabIndex        =   51
            Top             =   3885
            Width           =   1020
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Cono"
            Height          =   195
            Index           =   4
            Left            =   4605
            TabIndex        =   50
            Top             =   3375
            Width           =   600
         End
         Begin VB.Label lblCliente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1245
            TabIndex        =   49
            Top             =   2070
            Width           =   4875
         End
         Begin VB.Label lblConcesionario 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4005
            TabIndex        =   48
            Top             =   1620
            Width           =   2490
         End
         Begin VB.Label lblColorI 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4710
            TabIndex        =   47
            Top             =   1155
            Width           =   2370
         End
         Begin VB.Label lblColorE 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1215
            TabIndex        =   46
            Top             =   1155
            Width           =   2370
         End
         Begin VB.Label lblModelo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4365
            TabIndex        =   45
            Top             =   705
            Width           =   2370
         End
         Begin VB.Label lblMarca 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1215
            TabIndex        =   44
            Top             =   705
            Width           =   2370
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patente"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   250
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   705
            Width           =   450
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo"
            Height          =   195
            Index           =   2
            Left            =   3705
            TabIndex        =   41
            Top             =   705
            Width           =   525
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   3
            Left            =   6840
            TabIndex        =   40
            Top             =   705
            Width           =   285
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color Exterior"
            Height          =   195
            Index           =   5
            Left            =   105
            TabIndex        =   39
            Top             =   1155
            Width           =   930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Venta"
            Height          =   195
            Index           =   6
            Left            =   6600
            TabIndex        =   38
            Top             =   1620
            Width           =   915
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   37
            Top             =   2070
            Width           =   480
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Concesionario"
            Height          =   195
            Index           =   10
            Left            =   2850
            TabIndex        =   36
            Top             =   1620
            Width           =   1005
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cono"
            Height          =   195
            Index           =   11
            Left            =   105
            TabIndex        =   35
            Top             =   3390
            Width           =   375
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kms. Actuales"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   34
            Top             =   1635
            Width           =   1005
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Entrega"
            Height          =   195
            Index           =   13
            Left            =   4740
            TabIndex        =   33
            Top             =   3885
            Width           =   1050
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora Entrega"
            Height          =   195
            Index           =   14
            Left            =   8145
            TabIndex        =   32
            Top             =   3885
            Width           =   945
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color Interior"
            Height          =   195
            Index           =   21
            Left            =   3690
            TabIndex        =   31
            Top             =   1155
            Width           =   885
         End
         Begin VB.Label lblIdMarca 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   53
            Top             =   705
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblIdModelo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5670
            TabIndex        =   52
            Top             =   705
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblIdCliente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4950
            TabIndex        =   81
            Top             =   2085
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1530
         Left            =   -74955
         TabIndex        =   4
         Top             =   4560
         Width           =   11340
         Begin VB.Frame Frame3 
            Caption         =   "Deducible"
            Height          =   675
            Left            =   165
            TabIndex        =   9
            Top             =   765
            Width           =   5400
            Begin VB.TextBox txtDeduciblePesos 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3330
               MaxLength       =   8
               TabIndex        =   11
               Top             =   225
               Width           =   1920
            End
            Begin VB.TextBox txtDeducibleUF 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   720
               MaxLength       =   4
               TabIndex        =   10
               Top             =   240
               Width           =   1920
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pesos"
               Height          =   195
               Index           =   19
               Left            =   2730
               TabIndex        =   13
               Top             =   270
               Width           =   435
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "U.F."
               Height          =   195
               Index           =   20
               Left            =   105
               TabIndex        =   12
               Top             =   270
               Width           =   300
            End
         End
         Begin VB.TextBox txtLiquidador 
            Height          =   315
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1125
            Width           =   4020
         End
         Begin VB.TextBox txtNroPoliza 
            Height          =   315
            Left            =   6960
            MaxLength       =   30
            TabIndex        =   1
            Top             =   750
            Width           =   2940
         End
         Begin VB.TextBox txtNroSiniestro 
            Height          =   315
            Left            =   6960
            MaxLength       =   30
            TabIndex        =   0
            Top             =   360
            Width           =   2925
         End
         Begin VB.Label lblCompañia 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   150
            TabIndex        =   14
            Top             =   420
            Width           =   5445
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Siniestro"
            Height          =   195
            Index           =   18
            Left            =   5940
            TabIndex        =   8
            Top             =   405
            Width           =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Poliza"
            Height          =   195
            Index           =   17
            Left            =   5925
            TabIndex        =   7
            Top             =   825
            Width           =   645
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liquidador"
            Height          =   195
            Index           =   15
            Left            =   5940
            TabIndex        =   6
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compañia de Seguro"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   5
            Top             =   225
            Width           =   1485
         End
         Begin VB.Label lblIdCompañia 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4530
            TabIndex        =   74
            Top             =   420
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   330
         Left            =   -71300
         TabIndex        =   125
         Top             =   6500
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   6174
               MinWidth        =   6174
               Text            =   "Total Servicios Mecánica :"
               TextSave        =   "Total Servicios Mecánica :"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar2 
         Height          =   330
         Left            =   -71300
         TabIndex        =   126
         Top             =   6500
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   6174
               MinWidth        =   6174
               Text            =   "Total Servicios Mecánica :"
               TextSave        =   "Total Servicios Mecánica :"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar3 
         Height          =   330
         Left            =   3700
         TabIndex        =   127
         Top             =   6500
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   6174
               MinWidth        =   6174
               Text            =   "Total Servicios Mecánica :"
               TextSave        =   "Total Servicios Mecánica :"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               Object.Width           =   5292
               MinWidth        =   5292
               Text            =   "0"
               TextSave        =   "0"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   10890
      Top             =   360
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
            Picture         =   "frmOrdenTrabajoTaller.frx":0576
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":0688
            Key             =   "Menos"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":0AE0
            Key             =   "Mas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":0F38
            Key             =   "Persona"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":1390
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":14A2
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":15B4
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":16C6
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":17D8
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":18EA
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":19FC
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":1B0E
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":1C20
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":1D32
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":1E44
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":1F56
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":2068
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":217A
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":228C
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":239E
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":27F0
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOrdenTrabajoTaller.frx":2C42
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Registro (Ctrl+D)"
            ImageKey        =   "Borrar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Cerrar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOrdenTrabajoTaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset
Dim mstrSql As String
Dim mstrWhere As String
Dim mstrOrderBy As String

Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Dim mblnSw As Boolean

Dim itmAux As ListItem

Dim intIndice As Integer
Dim curValor As Currency



Function DatosCliente(strIdCliente As String) As Boolean
mstrSql = "SELECT Glbl_Cliente_Proveedor.Razon_Social as NOMBRE, Glbl_Cliente_Proveedor.Direccion AS DIREC, Glbl_Comuna.Descripcion AS COMUNA, Glbl_Cliente_Proveedor.Rut AS RUT ,Glbl_Cliente_Proveedor.Telefono AS FONO FROM Glbl_Cliente_Proveedor INNER JOIN Glbl_Comuna ON Glbl_Cliente_Proveedor.Id_Comuna = Glbl_Comuna.Id_Comuna Where Glbl_Cliente_Proveedor.Id_Cliente_Proveedor='" & strIdCliente & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            lblCliente = !Nombre
            lblDireccionCliente = !DirEC
            lblComunaCliente = !Comuna
            lblRutCliente = !Rut
            lblFono = !FONO
        End If
    End With
End If
Conexion.CloseHost adoPrincipal
End Function

Function ExisteRegistro(IdCiaSeguro As String, IdConcepto As String, IdPtePza As String) As Boolean
Dim adoTemp As ADODB.Recordset
ExisteRegistro = False
mstrSql = "SELECT top 1 * From Tllr_CiaSeguro_Concepto_Parte_Pieza"
mstrSql = mstrSql & " WHERE Id_Compañia_Seguro = '" & IdCiaSeguro & "'  AND Id_Concepto = '" & IdConcepto & "' AND Id_Parte_Pieza = '" & IdPtePza & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        ExisteRegistro = True
    Else
        mstrSql = "Insert into Tllr_CiaSeguro_Concepto_Parte_Pieza (Id_Compañia_Seguro, Id_Concepto, Id_Parte_Pieza, Valor, Horas) Values ('" & IdCiaSeguro & "' ,'" & IdConcepto & "' ,'" & IdPtePza & "',0,0)"
        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
            ExisteRegistro = True
        Else
            ExisteRegistro = False
        End If
    End If
End If
End Function


Sub FillInventarioOT(strIdEmpresa As String, strIdSucursal As String, strIdOT As String, strSeccion As String)

SetCheckOff lvwInventario
mstrSql = "SELECT Id_Estado_Recepcion as Codigo From Tllr_Inventario_OT"
mstrSql = mstrSql & " WHERE Id_Empresa = '" & strIdEmpresa & "' AND Id_Sucursal = '" & strIdSucursal & "' AND Id_OT = '" & Trim(lblNroOT) & "' AND Seccion_OT = '" & strSeccion & "'"
mstrSql = mstrSql & " Order by Id_Estado_Recepcion"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
    If Not .BOF And Not .EOF Then
        While Not .EOF
            Set lvwInventario.SelectedItem = lvwInventario.FindItem(CStr(!Codigo), , , 1)
            lvwInventario.SelectedItem.Checked = True
            .MoveNext
        Wend
    End If
    End With
End If

Conexion.CloseHost adoPrincipal

End Sub
Sub FillMecanicaOT(strIdEmpresa As String, strIdSucursal As String, strIdOT As String, strSeccion As String)

lvwServiciosMecanica.ListItems.Clear
mstrSql = "SELECT Tllr_Mecanica_OT.Id_Marca, Tllr_Mecanica_OT.Id_Modelo, Tllr_Mecanica_OT.Id_Servicio AS CODIGO, Tllr_Mecanica_OT.Seccion_Servicio, Tllr_Servicio.Descripcion AS NOMBRE, Tllr_Mecanica_OT.Horas AS TIEMPO, Tllr_Mecanica_OT.Valor AS VALOR"
mstrSql = mstrSql & " FROM Tllr_Servicio RIGHT OUTER JOIN Tllr_Servicio_Modelo ON Tllr_Servicio.Id_Servicio = Tllr_Servicio_Modelo.Id_Servicio RIGHT OUTER JOIN Tllr_Mecanica_OT ON  Tllr_Servicio_Modelo.Id_Marca = Tllr_Mecanica_OT.Id_Marca AND Tllr_Servicio_Modelo.Id_Modelo = Tllr_Mecanica_OT.Id_Modelo AND Tllr_Servicio_Modelo.Id_Servicio = Tllr_Mecanica_OT.Id_Servicio"
mstrSql = mstrSql & " WHERE Tllr_Mecanica_OT.Id_Empresa = '" & strIdEmpresa & "' AND Tllr_Mecanica_OT.Id_Sucursal = '" & strIdSucursal & "' AND Tllr_Mecanica_OT.Id_OT ='" & strIdOT & "' AND Tllr_Mecanica_OT.Seccion_OT = '" & strSeccion & "' "

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwServiciosMecanica.ListItems.Add(, , !Codigo)
            itmAux.SubItems(1) = !Nombre
            itmAux.SubItems(2) = !Valor
            itmAux.SubItems(3) = !TIEMPO
            .MoveNext
        Wend
    End If
    End With
End If

Conexion.CloseHost adoPrincipal

End Sub

Sub FillCarroceriaOT(strIdEmpresa As String, strIdSucursal As String, strIdOT As String, strSeccion As String, strIdCiaSeguro As String)

lvwServiciosCarroceria.ListItems.Clear
mstrSql = "SELECT Tllr_Carroceria_OT.Id_Concepto AS IDCONCEPTO, Tllr_Concepto.Descripcion AS CONCEPTO, Tllr_Carroceria_OT.Id_Parte_Pieza AS IDPIEZA, Tllr_Parte_Pieza.Descripcion AS PIEZA, Tllr_Carroceria_OT.Valor"
mstrSql = mstrSql & " FROM Tllr_Concepto RIGHT OUTER JOIN Tllr_CiaSeguro_Concepto ON Tllr_Concepto.Id_Concepto = Tllr_CiaSeguro_Concepto.Id_Concepto RIGHT OUTER JOIN Tllr_Parte_Pieza RIGHT OUTER JOIN    Tllr_CiaSeguro_Concepto_Parte_Pieza ON Tllr_Parte_Pieza.Id_Parte_Pieza = Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Parte_Pieza ON Tllr_CiaSeguro_Concepto.Id_Compañia_Seguro = Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Compañia_Seguro AND Tllr_CiaSeguro_Concepto.Id_Concepto = Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Concepto RIGHT OUTER JOIN Tllr_Carroceria_OT ON Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Compañia_Seguro = Tllr_Carroceria_OT.Id_Compañia_Seguro AND Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Concepto = Tllr_Carroceria_OT.Id_Concepto  AND Tllr_CiaSeguro_Concepto_Parte_Pieza.Id_Parte_Pieza = Tllr_Carroceria_OT.Id_Parte_Pieza"
mstrSql = mstrSql & " WHERE Tllr_Carroceria_OT.Id_Empresa = '" & strIdEmpresa & "' AND Tllr_Carroceria_OT.Id_Sucursal = '" & strIdSucursal & "' AND Tllr_Carroceria_OT.Id_OT ='" & strIdOT & "' AND Tllr_Carroceria_OT.Seccion_OT ='" & strSeccion & "' AND Tllr_Carroceria_OT.Id_Compañia_Seguro ='" & strIdCiaSeguro & "'"

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwServiciosCarroceria.ListItems.Add(, , !CONCEPTO)
            itmAux.SubItems(1) = !IdConcepto
            itmAux.SubItems(2) = TipoConcepto(!IdConcepto)
            itmAux.SubItems(3) = !PIEZA
            itmAux.SubItems(4) = !IDPIEZA
            itmAux.SubItems(5) = TraeValorDefinido(strIdCiaSeguro, !IdConcepto, !IDPIEZA)
            itmAux.SubItems(6) = !Valor
            .MoveNext
        Wend
    End If
    End With
End If

Conexion.CloseHost adoPrincipal

End Sub
Sub ServicioCarroceria(Accion As mAccionItem)
If Accion = mAddItem Then
    Set itmAux = lvwServiciosCarroceria.ListItems.Add(, , dtcConceptos.Text)
    itmAux.SubItems(1) = dtcConceptos.BoundText
    itmAux.SubItems(2) = txtSeccion.Text
    itmAux.SubItems(3) = dtcPartePieza.Text
    itmAux.SubItems(4) = dtcPartePieza.BoundText
    itmAux.SubItems(5) = Format$(txtValorDef, "##,###,##0")
    itmAux.SubItems(6) = Format$(txtValorFin, "##,###,##0")
End If


If Accion = mDelItem Then
    If lvwServiciosCarroceria.ListItems.Count > 0 Then
        lvwServiciosCarroceria.ListItems.Remove lvwServiciosCarroceria.SelectedItem.Index
    End If
End If


If Accion = mRefItem Then

End If


End Sub


Sub FillConceptosVsCiaSeguro(strCiaSeg As String)

mstrSql = "SELECT Tllr_CiaSeguro_Concepto.Id_Concepto as codigo, Tllr_Concepto.Descripcion as nombre"
mstrSql = mstrSql & " FROM Tllr_CiaSeguro_Concepto LEFT OUTER JOIN Tllr_Concepto ON Tllr_CiaSeguro_Concepto.Id_Concepto = Tllr_Concepto.Id_Concepto"
mstrSql = mstrSql & " WHERE Tllr_CiaSeguro_Concepto.Id_Compañia_Seguro = '" & strCiaSeg & "' "

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datConceptos
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcConceptos.ListField = "Nombre"
            dtcConceptos.BoundColumn = "Codigo"
            If .Recordset.RecordCount < 2 Then
                dtcConceptos.BoundText = .Recordset!Codigo
                dtcConceptos.Enabled = False
            End If
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal
End Sub


Sub DatosVehiculo(strPatente As String)

mstrSql = "SELECT Tllr_Vehiculo_Cliente.Patente,Tllr_Vehiculo_Cliente.Id_Marca AS IdMarca,Glbl_Marca.Descripcion AS MARCA,Tllr_Vehiculo_Cliente.Id_Modelo AS IdModelo,Glbl_Modelo.Descripcion AS MODELO,Tllr_Vehiculo_Cliente.Año,Glbl_Color_Interior.Descripcion AS COLORINTERIOR, Glbl_Color_Exterior.Descripcion AS COLOREXTERIOR, Glbl_Cliente_Proveedor.Rut, Glbl_Cliente_Proveedor.Razon_Social AS CLIENTE,Tllr_Compañia_Seguro.Nombre AS COMPAÑIA,Tllr_Vehiculo_Cliente.Id_Compañia_Seguro as IDCIA,Glbl_Concesionarios.Razon_Social AS CONCESIONARIO,Tllr_Vehiculo_Cliente.Kilometros_Actuales,Tllr_Vehiculo_Cliente.Fecha_Venta,Tllr_Vehiculo_Cliente.Deducible_UF AS UF,Tllr_Vehiculo_Cliente.Deducible_Pesos AS Pesos,Glbl_Modelo.Id_TipoVehiculo,Glbl_Tipo_Vehiculo.Descripcion AS TIPOVEHICULO,Tllr_Vehiculo_Cliente.VIN AS VIN,Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor as IDCLI"
mstrSql = mstrSql & " FROM Glbl_Color_Interior RIGHT OUTER JOIN Glbl_Cliente_Proveedor RIGHT OUTER JOIN Tllr_Compañia_Seguro RIGHT OUTER JOIN Tllr_Vehiculo_Cliente ON Tllr_Compañia_Seguro.Id_Compañia_Seguro = Tllr_Vehiculo_Cliente.Id_Compañia_Seguro ON Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor ON Glbl_Color_Interior.Id_Color_Interior = Tllr_Vehiculo_Cliente.Id_Color_Interior LEFT OUTER JOIN Glbl_Color_Exterior ON Tllr_Vehiculo_Cliente.Id_Color_Exterior = Glbl_Color_Exterior.Id_Color_Exterior LEFT OUTER JOIN Glbl_Concesionarios ON Tllr_Vehiculo_Cliente.Id_Concesionario = Glbl_Concesionarios.Id_Concesionario LEFT OUTER JOIN Glbl_Marca RIGHT OUTER JOIN Glbl_Tipo_Vehiculo RIGHT OUTER JOIN Glbl_Modelo ON Glbl_Tipo_Vehiculo.Id_TipoVehiculo = Glbl_Modelo.Id_TipoVehiculo ON Glbl_Marca.Id_Marca = Glbl_Modelo.Id_Marca ON Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca And Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo"
mstrSql = mstrSql & " WHERE Tllr_Vehiculo_Cliente.Patente='" & txtPatente & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        With adoPrincipal
            lblMarca = !Marca
            lblIdMarca = !IdMarca
            lblModelo = !Modelo
            lblIdModelo = !IdModelo
            lblTipoVeh = !TipoVehiculo
            lblVIN = !VIN
            txtAño = !Año
            lblColorE = !COLOREXTERIOR
            lblColorI = !COLORINTERIOR
            lblCliente = !CLIENTE
            lblConcesionario = !Concesionario
            lblFechaVenta = !Fecha_Venta
            txtKilAct = !Kilometros_Actuales
            lblCompañia = !COMPAÑIA
            lblIdCompañia = !IDCIA
            txtDeducibleUF = !UF
            txtDeduciblePesos = !Pesos
            lblIdCliente = !IDCLI
'            DatosCliente !IDCLI '/////////////////////////DATOS DEL CLIENTE
        End With
    Else
        If MsgBox("Este Patente No esta registrada " & Chr(13) & "Desea Registrar los Datos del Vehículo", 4 + 32, "Advertencia") = 6 Then
            gstrProcedencia = "Recepcion"
            frmMantenedorVehiculoCliente.Show
        Else
            txtPatente.Text = ""
            txtPatente.SetFocus
        End If
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Sub FillConceptosInventario()

mstrSql = "SELECT Id_Estado_Recepcion AS Codigo, Descripcion AS Nombre FROM Tllr_Estado_Recepcion WHERE Vigencia = 'S' Order By Id_Estado_Recepcion"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = lvwInventario.ListItems.Add(, , !Codigo)
            itmAux.SubItems(1) = !Nombre
            .MoveNext
        Wend
    End If
    End With
End If
End Sub
Sub FillTime(intHraIni As Integer, intHraFin As Integer)
Dim intHra As Integer, intMin As Integer

For intHra = intHraIni To intHraFin
    For intMin = 0 To 59 Step 30
        cboHora.AddItem Format$(intHra, "00") & ":" & Format$(intMin, "00")
    Next
Next
End Sub

Function GuardaCarroceria(strIdOT As String, strSeccion As String, strCiaSeguro As String) As Boolean
GuardaCarroceria = True
mstrSql = "DELETE Tllr_Carroceria_OT WHERE Id_OT='" & strIdOT & "' AND Seccion_OT ='" & strSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
    If lvwServiciosCarroceria.ListItems.Count > 0 Then
        For intIndice = 1 To lvwServiciosCarroceria.ListItems.Count
            Set lvwServiciosCarroceria.SelectedItem = lvwServiciosCarroceria.ListItems(intIndice)
            '/////////////////////////////////////////////////VALIDAR SI EXISTE EN PARENT
            If ExisteRegistro(strCiaSeguro, lvwServiciosCarroceria.SelectedItem.SubItems(1), lvwServiciosCarroceria.SelectedItem.SubItems(4)) = True Then
                mstrSql = "Insert Into Tllr_Carroceria_OT"
                mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,Id_OT , Seccion_OT, Id_Compañia_Seguro, Id_Concepto, Id_Parte_Pieza, Horas, Valor)"
                mstrSql = mstrSql & " Values('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "', '" & strIdOT & "', '" & strSeccion & "','" & strCiaSeguro & "', '" & Trim(lvwServiciosCarroceria.SelectedItem.SubItems(1)) & "',  '" & Trim(lvwServiciosCarroceria.SelectedItem.SubItems(4)) & "', " & CCur(Val(Format(lvwServiciosCarroceria.SelectedItem.SubItems(5), "######"))) & "," & CCur(Val(Format(lvwServiciosCarroceria.SelectedItem.SubItems(6), "######"))) & " ) "
                If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
                    GuardaCarroceria = False
                    Exit Function
                End If
            End If
        Next
    Else
        GuardaCarroceria = True
    End If
    
Else
    GuardaCarroceria = False
    Exit Function
End If
End Function

Function GuardaInventario(strIdOT As String, strSeccion As String) As Boolean
mstrSql = "DELETE Tllr_Inventario_OT WHERE Tllr_Inventario_OT.ID_OT='" & strIdOT & "' and Tllr_Inventario_OT.Seccion_OT='" & strSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
    For intIndice = 1 To lvwInventario.ListItems.Count
        Set lvwInventario.SelectedItem = lvwInventario.ListItems(intIndice)
        If lvwInventario.SelectedItem.Checked = True Then
            mstrSql = "Insert Into Tllr_Inventario_OT"
            mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,Id_Estado_Recepcion, Id_OT, Seccion_OT) "
            mstrSql = mstrSql & " values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "','" & lvwInventario.SelectedItem & "', '" & strIdOT & "', '" & strSeccion & "' )"
            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
                GuardaInventario = False
                Exit Function
            End If
        End If
    Next
    GuardaInventario = True
Else
    GuardaInventario = False
    Exit Function
End If

End Function

Function GuardaMecanica(strIdOT As String) As Boolean
GuardaMecanica = True
mstrSql = "DELETE Tllr_Mecanica_OT WHERE Seccion_OT='" & gstrSeccion & "' And ID_OT='" & strIdOT & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
    If lvwServiciosMecanica.ListItems.Count > 0 Then
        For intIndice = 1 To lvwServiciosMecanica.ListItems.Count
        Set lvwServiciosMecanica.SelectedItem = lvwServiciosMecanica.ListItems(intIndice)
        mstrSql = "Insert Into Tllr_Mecanica_OT"
        mstrSql = mstrSql & " (Id_Empresa, Id_Sucursal,Id_OT , Seccion_OT, Id_Marca, Id_Modelo, Id_Servicio, Seccion_Servicio, Valor,Horas)"
        mstrSql = mstrSql & " Values( '" & gstrIdEmpresa & "', '" & gstrIdSucursal & "','" & strIdOT & "', '" & gstrSeccion & "','" & Trim(lblIdMarca) & "','" & Trim(lblIdModelo) & "',  '" & Trim(lvwServiciosMecanica.SelectedItem) & "', 'M', " & Format(lvwServiciosMecanica.SelectedItem.SubItems(3), "#####0") & " , " & CCur(lvwServiciosMecanica.SelectedItem.SubItems(2)) & " ) "
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
            GuardaMecanica = False
            Exit Function
        End If
        Next
    Else
        GuardaMecanica = True
    End If
Else
    GuardaMecanica = False
    Exit Function
End If
End Function

Function TipoConcepto(strIdConcepto As String) As String

mstrSql = "SELECT TOP 1 D_P AS TIPO FROM Tllr_Concepto WHERE ID_CONCEPTO='" & strIdConcepto & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            TipoConcepto = !tipo
        Else
            TipoConcepto = "N"
        End If
    End With
End If

End Function
Sub FillTipoCono()
    dtcTipoCono.Enabled = True
    mstrSql = "SELECT Id_Tipo_Cono as codigo, Color as nombre FROM Tllr_Tipo_Cono WHERE Vigencia = 'S' order by Color"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datTipoCono
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcTipoCono.ListField = "Nombre"
                dtcTipoCono.BoundColumn = "Codigo"
                If .Recordset.RecordCount < 2 Then
                    dtcTipoCono.Enabled = False
                    dtcTipoCono.BoundText = .Recordset!Codigo
                End If
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Sub FillRecepcionista()
    mstrSql = "SELECT Id_Mecanico AS CODIGO, Nombre FROM Tllr_Mecanicos WHERE Es_Recepcionista = 'S' "
    dtcRecepcionista.Enabled = True
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datRecepcionista
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcRecepcionista.ListField = "Nombre"
                dtcRecepcionista.BoundColumn = "Codigo"
                If .Recordset.RecordCount < 2 Then
                    dtcRecepcionista.BoundText = .Recordset!Codigo
                    dtcRecepcionista.Enabled = False
                End If
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub


Sub FillGarantia()
mstrSql = "SELECT Id_Garantia AS CODIGO, Descripcion AS NOMBRE FROM Tllr_Garantias ORDER BY Descripcion"
dtcGarantia.Enabled = True
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datGarantia
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcGarantia.ListField = "Nombre"
            dtcGarantia.BoundColumn = "Codigo"
            If .Recordset.RecordCount < 2 Then
                dtcGarantia.BoundText = .Recordset!Codigo
                dtcGarantia.Enabled = False
            End If
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal
End Sub



Function LetSql(strWhere As String, strOrder As String) As String

mstrSql = "SELECT Id_Empresa,"
mstrSql = mstrSql & " Id_Sucursal, "
mstrSql = mstrSql & " Id_OT, "
mstrSql = mstrSql & " Seccion_OT, "
mstrSql = mstrSql & " Id_Garantia,"
mstrSql = mstrSql & " Folio_Garantia, "
mstrSql = mstrSql & " Id_Tipo_Cono, "
mstrSql = mstrSql & " Nro_Cono, "
mstrSql = mstrSql & " Patente,"
mstrSql = mstrSql & " RealizadoPor, "
mstrSql = mstrSql & " Estado, "
mstrSql = mstrSql & " Fecha_Emision, "
mstrSql = mstrSql & " Entrega_Estimada,"
mstrSql = mstrSql & " Hora_Entrega, "
mstrSql = mstrSql & " Nro_Siniestro, "
mstrSql = mstrSql & " Nro_Poliza, "
mstrSql = mstrSql & " Liquidador,"
mstrSql = mstrSql & " Total_Mecanica, "
mstrSql = mstrSql & " Total_Carroceria, "
mstrSql = mstrSql & " Total_Desabolladura,"
mstrSql = mstrSql & " Total_Pintura, "
mstrSql = mstrSql & " Total_Terceros, "
mstrSql = mstrSql & " Total_Repuestos, "
mstrSql = mstrSql & " Total_OT,"
mstrSql = mstrSql & " Comentario , "
mstrSql = mstrSql & " Solicitado_Por"
mstrSql = mstrSql & " From Tllr_OT"
'mstrSql = mstrSql & " WHERE Id_Empresa = '01' AND Id_Sucursal = '01' AND"
'mstrSql = mstrSql & " Id_OT = '01' AND Seccion_OT = '01'"


LetSql = mstrSql & " " & strWhere & " " & strOrder

End Function

Private Sub LeerCampos()

If mblnTablaVacia Then
    LimpiaCampos
    Exit Sub
End If

With adoPrincipal
    lblNroOT.Caption = !Id_OT
    dtcGarantia.BoundText = !Id_Garantia
    dtcTipoCono.BoundText = !Id_Tipo_Cono
    
    dtcRecepcionista.BoundText = !Recepcionista
    txtNroCono = !Nro_Cono
    pckFechaAtencion.Value = !Fecha_Atencion
    pckFechaEntrega.Value = !Fecha_Entrega
    cboHora.Text = !Hora_Entrega
    txtNroSiniestro = !Nro_Siniestro
    txtNroPoliza = !Nro_Poliza
    txtLiquidador = !Liquidador
    txtComentario = !Comentario
    txtPatente = !PATENTE
    txtFolioGarantia = !Folio_Garantia
    txtSolicita = !Solicitado_Por
    DatosVehiculo !PATENTE
    '/////////////////////////////////////////////////////////////////////////////////
    FillInventarioOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    FillMecanicaOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion
    FillCarroceriaOT gstrIdEmpresa, gstrIdSucursal, !Id_OT, gstrSeccion, lblIdCompañia
    '/////////////////////////////////////////////////////////////////////////////////
End With
End Sub
Function TraeValorDefinido(strCiaSeg As String, strConcepto As String, strPartePieza As String) As Currency

mstrSql = "SELECT Valor FROM Tllr_CiaSeguro_Concepto_Parte_Pieza"
mstrSql = mstrSql & " WHERE Id_Compañia_Seguro = '" & strCiaSeg & "' AND Id_Concepto = '" & strConcepto & "' AND Id_Parte_Pieza = '" & strPartePieza & "'"

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            TraeValorDefinido = IIf(Not IsNull(!Valor), !Valor, 0)
        Else
            TraeValorDefinido = 0
        End If
    End With
End If
End Function

Function VerificaServicioCarroceria(strIdConcepto As String, strIdParte As String) As Boolean

VerificaServicioCarroceria = True
For intIndice = 1 To lvwServiciosCarroceria.ListItems.Count
    Set lvwServiciosCarroceria.SelectedItem = lvwServiciosCarroceria.ListItems(intIndice)
    If lvwServiciosCarroceria.SelectedItem.SubItems(1) = strIdConcepto Then
        If lvwServiciosCarroceria.SelectedItem.SubItems(4) = strIdParte Then
            VerificaServicioCarroceria = False
        End If
    End If
Next intIndice


End Function

Private Sub dtcConceptos_Change()
dtcPartePieza.BoundText = ""
txtSeccion.Text = TipoConcepto(dtcConceptos.BoundText)
End Sub

Private Sub dtcPartePieza_Change()
curValor = TraeValorDefinido(lblIdCompañia, dtcConceptos.BoundText, dtcPartePieza.BoundText)
txtValorDef = Format$(curValor, "##,###,##0")
txtValorFin = Format$(curValor, "##,###,##0")
txtValorFin.SetFocus

End Sub

Private Sub Form_Load()
    mblnSw = True
    gstrSeccion = "M"
    stbServicios.Tab = 0
End Sub



Private Sub lblIdCliente_Change()
If DatosCliente(lblIdCliente) Then DoEvents
End Sub

Private Sub lblIdCompañia_Change()
FillConceptosVsCiaSeguro Trim(lblIdCompañia) '////LLENAR LOS CONCEPTOS QUE MANEJA LA CIA SEG
End Sub

Private Sub optRecepcion_Click(Index As Integer)
Select Case Index
Case 0
    gstrSeccion = "M"
    Renovar
    stbServicios.TabEnabled(3) = False
    Frame1.Enabled = False
Case 1
    gstrSeccion = "C"
    Renovar
    stbServicios.TabEnabled(3) = True
    Frame1.Enabled = True
End Select

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
            ImprimirInforme
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
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSw Then
        RevizaAtributos
        FillConceptosInventario
        FillGarantia
        FillRecepcionista
        FillTipoCono
        FillTime 9, 20
        FillPartePieza
        If gapAccion = apcrear Then
           AgregarRegistro
           lblNroOT = gstrBusca
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
            'WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT='" & gstrBusca & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"""
                mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
                gstrSql = LetSql(mstrWhere, mstrOrderBy)
                If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost adoPrincipal
            End If
            Me.SetFocus
        End If
        If gapAccion = apninguno Then
           Renovar
        End If
        optRecepcion(0).Value = True
    End If
    gapAccion = apninguno
    mblnSw = False
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
            ImprimirInforme
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
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    Me.lblNroOT = TraeCorrelativo(IIf(gstrSeccion = "M", gcRecepcionMecanica, gcRecepcionCarroceria), gstrIdEmpresa, gstrIdSucursal)
    SetCheckOff lvwInventario
    lvwServiciosMecanica.ListItems.Clear
    lvwServiciosCarroceria.ListItems.Clear
    stbServicios.Tab = 0
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT >'" & Trim(lblNroOT) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
    gstrSql = LetSql(mstrWhere, mstrOrderBy)
    
'    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
    
    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.ID_OT < '" & Trim(lblNroOT) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
            gstrSql = LetSql(mstrWhere, mstrOrderBy)
            
'            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo
            
            If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost adoPrincipal
    'txtNombre.SetFocus
End Sub
Private Sub GrabarRegistro()
    If Not Validacion() Then
        Exit Sub
    End If

    If Me.Tag = "Crear" Then
        mstrSql = "INSERT INTO Tllr_OT (Id_Empresa, Id_Sucursal, Id_OT , Id_Garantia, Seccion_OT, Id_Tipo_Cono, Nro_Cono, Patente, Recepcionista, Fecha_Atencion, Fecha_Entrega, Hora_Entrega, Nro_Siniestro, Nro_Poliza, Liquidador, Comentario, Folio_Garantia, Solicitado_Por ) "                                                                                                             'Nro_Siniestro, Nro_Poliza, Liquidador, Comentario                                                                                                    Folio_Garantia, Solicitado_Por
        mstrSql = mstrSql & " values ('" & gstrIdEmpresa & "', '" & gstrIdSucursal & "','" & Trim(lblNroOT) & "', '" & Trim(dtcGarantia.BoundText) & "','" & gstrSeccion & "','" & dtcTipoCono.BoundText & "', " & CLng(txtNroCono.Text) & ",'" & txtPatente.Text & "','" & dtcRecepcionista.BoundText & "','" & pckFechaAtencion.Value & "','" & pckFechaEntrega.Value & "','" & cboHora.Text & "','" & UCase(Trim(txtNroSiniestro.Text)) & "','" & UCase(Trim(txtNroPoliza.Text)) & "','" & UCase(Trim(txtLiquidador.Text)) & "','" & UCase(Trim(txtComentario.Text)) & "','" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), ".") & "','" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), ".") & "' )"
    Else
        mstrSql = "UPDATE Tllr_OT SET Id_Garantia='" & Trim(dtcGarantia.BoundText) & "', Id_Tipo_Cono='" & dtcTipoCono.BoundText & "', Nro_Cono=" & CLng(txtNroCono.Text) & ", Patente='" & txtPatente.Text & "', Recepcionista='" & dtcRecepcionista.BoundText & "', Fecha_Atencion='" & pckFechaAtencion.Value & "', Fecha_Entrega='" & pckFechaEntrega.Value & "', Hora_Entrega='" & cboHora.Text & "', Nro_Siniestro='" & UCase(Trim(txtNroSiniestro.Text)) & "', Nro_Poliza='" & UCase(Trim(txtNroPoliza.Text)) & "', Liquidador='" & UCase(Trim(txtLiquidador.Text)) & "', Comentario='" & UCase(Trim(txtComentario.Text)) & "', Folio_Garantia='" & IIf(Trim(txtFolioGarantia) <> "", UCase(Trim(txtFolioGarantia)), ".") & "',Solicitado_Por='" & IIf(Trim(txtSolicita) <> "", UCase(Trim(txtSolicita)), ".") & "' "
        mstrSql = mstrSql & " where Id_OT ='" & Trim(Trim(lblNroOT)) & "' AND Seccion_OT='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    End If
    
    
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        '////////////////GRABAR INVENTARIO
        If GuardaInventario(lblNroOT, gstrSeccion) = False Then MsgBox "Guardar Inventario Fallo, Verifique"
        '////////////////GRABAR MECANICA
        If GuardaMecanica(lblNroOT) = False Then MsgBox "Guardar Mecanica Fallo, Verifique"
        '////////////////GRABAR CARROCERIA
        If GuardaCarroceria(lblNroOT, gstrSeccion, lblIdCompañia) = False Then MsgBox "Guardar Inventario Fallo, Verifique"
        '////////////////INCREMENTO CORRELATIVO
        If Me.Tag = "Crear" Then
            If gstrSeccion = "M" Then
                Call IncrementaCorrelativo(gcRecepcionMecanica, gstrIdEmpresa, gstrIdSucursal)
            Else
                Call IncrementaCorrelativo(gcRecepcionCarroceria, gstrIdEmpresa, gstrIdSucursal)
            End If
        End If
        '///////////////////////ACTUALIZA DATOS DEL VEHICULO
        If gstrSeccion = "M" Then
            mstrSql = " Update Tllr_Vehiculo_Cliente Set Kilometros_Actuales =" & IIf(Trim(txtKilAct) <> "", CLng(txtKilAct), 0) & ""
            mstrSql = mstrSql & " Where Patente='" & txtPatente & "'"
            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then MsgBox "Actualizar Datos Foraneos Fallo, Verifique"
        Else
            mstrSql = " Update Tllr_Vehiculo_Cliente Set Kilometros_Actuales =" & IIf(Trim(txtKilAct) <> "", CLng(txtKilAct), 0) & ", Deducible_UF=" & IIf(Trim(txtDeducibleUF) <> "", CCur(txtDeducibleUF), 0) & ", Deducible_Pesos=" & IIf(Trim(txtDeduciblePesos) <> "", CCur(txtDeduciblePesos), 0) & ""
            mstrSql = mstrSql & " Where Patente='" & txtPatente & "'"
            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then MsgBox "Actualizar Datos Foraneos Fallo, Verifique"
        End If
        
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
    
    
    
End Sub
Private Sub BorrarRegistro()
    
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        '////////////////////////////////ELIMINAR SERVICIOS DE MECANICA///////////////////////////////////
        mstrSql = "DELETE FROM Tllr_Mecanica_OT  WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT=" & Trim(lblNroOT) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        '////////////////////////////////ELIMINAR SERVICIOS DE CARRPCERIA///////////////////////////////////
        mstrSql = "DELETE FROM Tllr_Carroceria_OT WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT=" & Trim(lblNroOT) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        '////////////////////////////////////ELIMINAR INENTARIO///////////////////////////////
        mstrSql = "DELETE FROM Tllr_Inventario WHERE Seccion_OT = '" & gstrSeccion & "' AND Id_OT=" & Trim(lblNroOT) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
        '//////////////////////////////////////ENCABEZADO/////////////////////////////
        mstrSql = "DELETE FROM Tllr_OT WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT=" & Trim(lblNroOT) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
'            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
            mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > " & Trim(lblNroOT) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
            mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
            gstrSql = LetSql(mstrWhere, mstrOrderBy)
            If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    'mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo
                    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < " & Trim(lblNroOT) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
                    gstrSql = LetSql(mstrWhere, mstrOrderBy)
                    
                    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                            LeerCampos
                        Else
                            mblnTablaVacia = True
                            LimpiaCampos
                        End If
                    End If
                End If
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
End Sub
Private Sub BuscarRegistro()
'    Set FormVol1 = New APFORM1.APFORM
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption)
    If gstrBusca <> "" Then
'        mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
        
        mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT=  " & CLng(gstrBusca) & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
        gstrSql = LetSql(mstrWhere, mstrOrderBy)
        
        If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
    Me.SetFocus
End Sub
Private Sub ImprimirInforme()
   ' FormVol1.ImprimirRegistros Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption, gstrPathReporte, "APCARROC.RPT", gstrUSUARIO, gstrCodigoEmpresa
   ImprimirDocumento gRecepcion
End Sub
Private Sub PrimerRegistro()
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT"
    gstrSql = LetSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroAnterior()
    
    'mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo & " DESC"
    'AND Tllr_Recepcion=  " & CLng(gstrBusca) &
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT < '" & Trim(lblNroOT) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
    gstrSql = LetSql(mstrWhere, mstrOrderBy)
    
    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroSiguiente()

    'mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
    
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' AND Tllr_OT.Id_OT > '" & Trim(lblNroOT) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT "
    gstrSql = LetSql(mstrWhere, mstrOrderBy)
    
    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub UltimoRegistro()

    'mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo & " DESC"
    
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT DESC"
    gstrSql = LetSql(mstrWhere, mstrOrderBy)
    
    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub Renovar()
    mstrWhere = " WHERE Tllr_OT.Seccion_OT = '" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    mstrOrderBy = " ORDER BY Tllr_OT.Id_OT "
    gstrSql = LetSql(mstrWhere, mstrOrderBy)
    If Conexion.SendHost(gstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        VerificaTablaVacia
        ActivaBotones
        If Not mblnTablaVacia Then
            PrimerRegistro
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
    'txtCodigo.Enabled = False
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
    End With
End Sub
Private Sub DesactivaBotones()
    'txtCodigo.Enabled = True
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = False
        .Item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
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

Private Sub LimpiaCampos()
With Me
    SetCheckOff .lvwInventario
    .lvwServiciosCarroceria.ListItems.Clear
    .lvwServiciosMecanica.ListItems.Clear
    .lblNroOT.Caption = ""
    .dtcGarantia.BoundText = ""
    .pckFechaAtencion.Value = Now
    .txtPatente.Text = ""
    .lblMarca.Caption = "": .lblIdMarca = ""
    .lblModelo.Caption = "": .lblIdModelo = ""
    .txtAño.Text = ""
    .lblColorE.Caption = ""
    .lblColorI.Caption = ""
    .lblTipoVeh.Caption = ""
    .lblCliente.Caption = ""
    .txtKilAct.Text = ""
    .lblConcesionario.Caption = ""
    .lblFechaVenta.Caption = ""
    .dtcTipoCono.BoundText = ""
    .txtNroCono.Text = ""
    .dtcRecepcionista.BoundText = ""
    .pckFechaEntrega.Value = Now
    .cboHora.Text = ""
    .lblCompañia.Caption = ""
    .lblIdCompañia.Caption = ""
    .txtDeducibleUF.Text = "0"
    .txtDeduciblePesos.Text = "0"
    .txtNroSiniestro.Text = ""
    .txtNroPoliza.Text = ""
    .txtLiquidador.Text = ""
    .lblFono.Caption = ""
    .lblVIN.Caption = ""
    .txtSolicita.Text = ""
    .txtFolioGarantia.Text = ""
    .lblRutCliente.Caption = ""
    .lblComunaCliente.Caption = ""
    .lblDireccionCliente.Caption = ""
    .lblIdCliente.Caption = ""
End With
End Sub
Private Sub ValoresporDefecto()
    txtAño.Text = Year(Now)
    txtDeducibleUF.Text = "0"
    txtNroCono.Text = "0"
    txtDeduciblePesos.Text = "0"
    txtNroSiniestro.Text = "."
    txtNroPoliza.Text = "."
    txtLiquidador.Text = "."
    txtKilAct.Text = "0"
End Sub
Private Function Validacion() As Boolean
    Validacion = True
With Me
    If .dtcGarantia.BoundText = "" Then
        MsgBox "La Garantía  debe Especificarse...", vbInformation, "Advertencia"
        dtcGarantia.Enabled = True
        dtcGarantia.SetFocus
        Validacion = False
        Exit Function
    End If
    If .txtPatente = "" Then
        MsgBox "La Patente debe Especificarse...", vbInformation, "Advertencia"
        txtPatente.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If .txtKilAct = "" Then
        MsgBox "Los Kilometros deben Especificarse...", vbInformation, "Advertencia"
        txtKilAct.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If .dtcTipoCono.BoundText = "" Then
        MsgBox "El Tipo de Cono debe Especificarse...", vbInformation, "Advertencia"
        dtcTipoCono.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If .txtNroCono = "" Then
        MsgBox "El Numero de Cono debe Especificarse...", vbInformation, "Advertencia"
        txtNroCono.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If .dtcRecepcionista.BoundText = "" Then
        MsgBox "El Recepcionista debe Especificarse...", vbInformation, "Advertencia"
        dtcRecepcionista.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If .cboHora.Text = "" Then
        MsgBox "La Hora de Entrega debe Especificarse...", vbInformation, "Advertencia"
        cboHora.SetFocus
        Validacion = False
        Exit Function
    End If
    
    '//////////////////////////////////CARROCERIA
    If .optRecepcion(1).Value = True Then
    
        If .txtDeducibleUF.Text = "" Then
            MsgBox "El Deducible en UF debe Especificarse...", vbInformation, "Advertencia"
            txtDeducibleUF.SetFocus
            Validacion = False
            Exit Function
        End If
        
        If .txtDeduciblePesos.Text = "" Then
            MsgBox "El Deducible en Pesos debe Especificarse...", vbInformation, "Advertencia"
            txtDeduciblePesos.SetFocus
            Validacion = False
            Exit Function
        End If
        
        If .txtNroSiniestro.Text = "" Then
            MsgBox "El Deducible en Pesos debe Especificarse...", vbInformation, "Advertencia"
            txtNroSiniestro.SetFocus
            Validacion = False
            Exit Function
        End If
        
        If .txtNroPoliza.Text = "" Then
            MsgBox "El Deducible en Pesos debe Especificarse...", vbInformation, "Advertencia"
            txtNroPoliza.SetFocus
            Validacion = False
            Exit Function
        End If
        
        If .txtLiquidador.Text = "" Then
            MsgBox "El Deducible en Pesos debe Especificarse...", vbInformation, "Advertencia"
            txtLiquidador.SetFocus
            Validacion = False
            Exit Function
        End If
        
    End If
    '//////////////////////////////////CARROCERIA
    
End With
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As ADODB.Recordset
        mstrSql = "select ID_OT from TLLR_OT where SECCION_OT = '" & gstrSeccion & "' AND ID_OT ='" & Trim(lblNroOT) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción "
                Validacion = False
'                txtCodigo.SetFocus
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
    
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmRecepcionTaller = Nothing
    gstrBusca = lblNroOT.Caption
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

Sub FillPartePieza()
mstrSql = "select Id_Parte_Pieza as CODIGO, Descripcion AS NOMBRE From Tllr_Parte_Pieza Where Vigencia ='S' ORDER BY Descripcion"
dtcPartePieza.Enabled = True
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datPartesPiezas
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcPartePieza.ListField = "Nombre"
            dtcPartePieza.BoundColumn = "Codigo"
            If .Recordset.RecordCount < 2 Then
                dtcPartePieza.BoundText = .Recordset!Codigo
                dtcPartePieza.Enabled = False
            End If
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal
End Sub

Private Sub tlbOpciones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Select Case Index
Case 0 '///////////////////////MECANICA
    Select Case Button.Key
    Case "Agregar" ' ////////////////AGREGAR
        If Trim(txtPatente.Text) <> "" And Len(txtPatente) = 6 Then
            gstrProcedencia = "Mecánica"
            frmAddServiciosMarMod.Show 1
        End If
    Case "Quitar" ' ////////////////QUITAR
        If Not lvwServiciosMecanica.SelectedItem Is Nothing Then
            lvwServiciosMecanica.ListItems.Remove (lvwServiciosMecanica.SelectedItem.Index)
        End If
    End Select
Case 1 '///////////////////////CARROCERIA
    Select Case Button.Key
    Case "Agregar" ' ////////////////AGREGAR
        If VerificaServicioCarroceria(dtcConceptos.BoundText, dtcPartePieza.BoundText) = True Then
            Call ServicioCarroceria(mAddItem)
        Else
            '//////////////MENSAJE
        End If
    Case "Quitar" ' ////////////////QUITAR
        Call ServicioCarroceria(mDelItem)
    End Select

End Select
End Sub

Private Sub tlbPatente_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Nuevo"
    gstrProcedencia = "Recepcion"
    frmMantenedorVehiculoCliente.Show 1
Case "Buscar"
    gstrProcedencia = "Recepcion"
    frmBuscaVehiculo.Show 1
End Select
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
KEYASCY = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
End Sub

Private Sub txtValorFin_GotFocus()
txtValorFin.SelStart = 0
txtValorFin.SelLength = Len(txtValorFin)
End Sub


