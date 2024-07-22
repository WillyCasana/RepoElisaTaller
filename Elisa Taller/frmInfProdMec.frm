VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfProdMec 
   Caption         =   "Informe de Productividad por Mecánico"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15540
   Icon            =   "frmInfProdMec.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   15540
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   495
      Left            =   11040
      TabIndex        =   48
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerGestion 
      Appearance      =   0  'Flat
      Caption         =   "Ver Gestión"
      Height          =   495
      Left            =   13200
      TabIndex        =   39
      Top             =   90
      Width           =   1095
   End
   Begin Crystal.CrystalReport rptProdMec 
      Left            =   6975
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame5 
      Caption         =   "Información Adicional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   9960
      TabIndex        =   33
      Top             =   5880
      Visible         =   0   'False
      Width           =   4785
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "% de Productividad :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   35
         Top             =   2100
         Width           =   1770
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas Realizadas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   1650
         Width           =   2115
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Resumen Estadístico del Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4980
      Left            =   9960
      TabIndex        =   22
      Top             =   720
      Width           =   5505
      Begin VB.CommandButton cmdCambiaDiasHabiles 
         Caption         =   "N°"
         Height          =   375
         Left            =   4320
         TabIndex        =   36
         ToolTipText     =   "Cambia el Número de Dias Habiles"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblProdAsig 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   54
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "% de Produc. Hrs Asig:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   3480
         Width           =   1965
      End
      Begin VB.Label Label8 
         Caption         =   "Total Horas Asignadas :"
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
         Left            =   195
         TabIndex        =   52
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblHorasAsignadas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   51
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label lblValorPesos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   47
         Top             =   5445
         Width           =   1500
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total en (S/.):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   5520
         Width           =   1230
      End
      Begin VB.Label lblValorHora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   45
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Valor Hora :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   5160
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas Reales :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   2640
         Width           =   1770
      End
      Begin VB.Label lblHorasReales1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   40
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Total O/t Trabajadas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   3960
         Width           =   1920
      End
      Begin VB.Label lblOtTrabajadas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   37
         Top             =   3840
         Width           =   1500
      End
      Begin VB.Label lblTotHorEst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   32
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas Estimadas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   31
         Top             =   1230
         Width           =   2040
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Días Habiles Estimados :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   420
         Width           =   2145
      End
      Begin VB.Label lblDiasHabEst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   29
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblHorasReales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   28
         Top             =   1575
         Width           =   1500
      End
      Begin VB.Label lblPorProd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   27
         Top             =   2940
         Width           =   1500
      End
      Begin VB.Label lblHorEst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2670
         TabIndex        =   26
         Top             =   765
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas Realizadas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   25
         Top             =   1650
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Horas/Día Estimadas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   24
         Top             =   810
         Width           =   1950
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "% de Produc.Hrs Realiz. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   23
         Top             =   3060
         Width           =   2190
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "D y P"
      Height          =   2250
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Visible         =   0   'False
      Width           =   9525
      Begin MSComctlLib.ListView lvwDyP 
         Height          =   1590
         Left            =   75
         TabIndex        =   15
         Top             =   180
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OT  Nº"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sección"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha OT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Horas"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Horas Asignadas"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label lblTotalCar3 
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
         Height          =   360
         Left            =   7800
         TabIndex        =   55
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         TabIndex        =   21
         Top             =   1830
         Width           =   1125
      End
      Begin VB.Label lblTotalCar 
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
         Height          =   360
         Left            =   6105
         TabIndex        =   20
         Top             =   1800
         Width           =   1185
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mecánica"
      Height          =   3210
      Left            =   45
      TabIndex        =   11
      Top             =   810
      Width           =   9645
      Begin MSComctlLib.ListView lvwMecanica 
         Height          =   2415
         Left            =   75
         TabIndex        =   12
         Top             =   180
         Width           =   9125
         _ExtentX        =   16087
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OT  Nº"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mecánico"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sección/N°Factura"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Fact"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Horas"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Horas Reales"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Horas Asignadas"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label lblTotalMec3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   360
         Left            =   8400
         TabIndex        =   49
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label lblTotalMec2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   360
         Left            =   7440
         TabIndex        =   42
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5265
         TabIndex        =   17
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label lblTotalMec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   360
         Left            =   6450
         TabIndex        =   16
         Top             =   2775
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Otros Servicios"
      Height          =   3210
      Left            =   45
      TabIndex        =   10
      Top             =   4200
      Width           =   9645
      Begin MSComctlLib.ListView lvwOtro 
         Height          =   2415
         Left            =   75
         TabIndex        =   14
         Top             =   180
         Width           =   9125
         _ExtentX        =   16087
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OT  Nº"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mecánico"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sección/N°Factura"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Fact."
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Horas"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Horas Reales"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Horas Asignadas"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label lblTotalOtro3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   360
         Left            =   8400
         TabIndex        =   50
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label lblTotalOtro2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   360
         Left            =   7440
         TabIndex        =   43
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Total Horas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   19
         Top             =   2745
         Width           =   1125
      End
      Begin VB.Label lblTotalOtr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
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
         Height          =   360
         Left            =   6345
         TabIndex        =   18
         Top             =   2760
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   495
      Left            =   9945
      TabIndex        =   8
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   12120
      TabIndex        =   7
      Top             =   90
      Width           =   1080
   End
   Begin VB.CommandButton cmdCerrar 
      Appearance      =   0  'Flat
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   14280
      TabIndex        =   6
      Top             =   90
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo dtcSupervisor 
      Bindings        =   "frmInfProdMec.frx":179A
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc datSupervisor 
      Height          =   330
      Left            =   2175
      Top             =   240
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
   Begin MSComCtl2.DTPicker pckFechaDesde 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   210
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   184025089
      CurrentDate     =   36776
   End
   Begin MSComCtl2.DTPicker pckFechaHasta 
      Height          =   315
      Left            =   5565
      TabIndex        =   3
      Top             =   210
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   184025089
      CurrentDate     =   36776
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   12000
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Resumen Mano de Obra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   13
      Top             =   615
      Width           =   2055
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   2340
      X2              =   5865
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   2340
      X2              =   5865
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Término"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5550
      TabIndex        =   5
      Top             =   0
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   3855
      TabIndex        =   4
      Top             =   15
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mecánico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   15
      Width           =   840
   End
End
Attribute VB_Name = "frmInfProdMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW  As Boolean
Dim mitmAux As ListItem
Dim mstrNumeroDocumento As String

Sub FillMecanicos()
gstrSql = "SELECT Id_Mecanico AS Codigo, Nombre FROM Tllr_Mecanicos where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and vigencia='S'  and Es_Recepcionista='N' and Es_Liquidador='N' order by Nombre"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
With datSupervisor
    Set .Recordset = gadoPrincipal
    If Not .Recordset.BOF And Not .Recordset.EOF Then
        .Recordset.MoveFirst
        dtcSupervisor.ListField = "Nombre"
        dtcSupervisor.BoundColumn = "Codigo"
    End If
End With
End If
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub

Sub FillMecanica(pstrIdMecanico As String, pdteFechaIni As Date, pdteFechaFin As Date)
Dim adoTemp As New ADODB.Recordset
    
    lvwMecanica.ListItems.Clear
    Screen.MousePointer = 11
    
    '/// llama al procedimiento almacenado
'    gstrSql = "Exec Tllr_ProdMecanico_Mecanica " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & pstrIdMecanico & "','" & pdteFechaIni & "','" & pdteFechaFin & "','" & gstrEstadoProdMecanico & "'"
    
    gstrSql = "SELECT SUM(Tllr_Mecanica_OT.Horas) AS TOTALHORAS, Sum(isnull(Tllr_Mecanica_Ot.HorasReales,0)) as TotalHorasReales, Tllr_Facturacion.Id_Empresa, Tllr_Facturacion.Id_Sucursal, Tllr_Facturacion.Id_OT,"
    gstrSql = gstrSql & " Tllr_Facturacion.Seccion_OT, Tllr_Mecanica_OT.Mecanico_Designado, Tllr_Mecanicos.Nombre as Mecanico ,Tllr_Facturacion.Estado,Tllr_Facturacion.Id_Cargo,"
    gstrSql = gstrSql & " Tllr_Facturacion.Nro_Factura_Emitida As NUMDOCUMENTO, Tllr_Facturacion.Fecha_Facturacion As Fecha_Emision,Tllr_Facturacion.Fecha_Liquidacion"
    gstrSql = gstrSql & " FROM Tllr_Facturacion LEFT OUTER JOIN"
    gstrSql = gstrSql & " Tllr_Mecanica_OT ON Tllr_Facturacion.Id_Empresa = Tllr_Mecanica_OT.Id_Empresa AND"
    gstrSql = gstrSql & " Tllr_Facturacion.Id_Sucursal = Tllr_Mecanica_OT.Id_Sucursal AND Tllr_Facturacion.Id_OT = Tllr_Mecanica_OT.Id_OT AND"
    gstrSql = gstrSql & " Tllr_Facturacion.Seccion_OT = Tllr_Mecanica_OT.Seccion_OT AND"
    gstrSql = gstrSql & " Tllr_Facturacion.Id_Cargo = Tllr_Mecanica_OT.Id_Tipo_Cargo"
    'kjcv 12.08.16
    gstrSql = gstrSql & " INNER JOIN Tllr_Mecanicos on Tllr_Mecanicos.Id_Mecanico=Tllr_Mecanica_OT.Mecanico_Designado"
    gstrSql = gstrSql & " GROUP BY Tllr_Facturacion.Id_Empresa, Tllr_Facturacion.Id_Sucursal, Tllr_Facturacion.Id_OT,"
    gstrSql = gstrSql & " Tllr_Facturacion.Seccion_OT, Tllr_Facturacion.Estado, Tllr_Mecanica_OT.Mecanico_Designado, Tllr_Mecanicos.Nombre, "
    gstrSql = gstrSql & " Tllr_Facturacion.Nro_Factura_Emitida, Tllr_Facturacion.Fecha_Facturacion,"
    gstrSql = gstrSql & " Tllr_Facturacion.Fecha_Liquidacion,Tllr_Mecanica_Ot.Id_OT, Tllr_MEcanica_Ot.Id_Tipo_Cargo,"
    gstrSql = gstrSql & " Tllr_Facturacion.Id_Cargo"
    gstrSql = gstrSql & " HAVING Tllr_Mecanica_OT.Id_OT NOT IN (SELECT Id_Ot FROM Tllr_Actividades_Mecanico)"
    gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Empresa = '" & gstrIdEmpresa & "'"
    gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Sucursal = '" & gstrIdSucursal & "'"
    'kjcv 15.08.16
    If dtcSupervisor.BoundText <> "" Then
    gstrSql = gstrSql & " AND Tllr_Mecanica_OT.Mecanico_Designado = '" & pstrIdMecanico & "'"
    End If
    gstrSql = gstrSql & " And Tllr_Facturacion.Id_Cargo = Tllr_Mecanica_Ot.Id_Tipo_Cargo"
    gstrSql = gstrSql & " AND Tllr_Facturacion.Fecha_Facturacion BETWEEN '" & pdteFechaIni & "' AND '" & pdteFechaFin & "'"
    If gstrEstadoProdMecanico = "F" Then
        gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B')"
    ElseIf gstrEstadoProdMecanico = "A" Then
        gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B','V')"
    Else
        gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('V')"
    End If
    
    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With gadoPrincipal
       If Not .BOF And Not .EOF Then
          While Not .EOF
          
                If !estado = "B" Or !estado = "F" Then
                    mstrNumeroDocumento = ValorNulo(!NUMDOCUMENTO)
                Else
                    mstrNumeroDocumento = "S/N"
                End If
                
                Set mitmAux = lvwMecanica.ListItems.Add(, , !Id_OT)
                
                mitmAux.SubItems(1) = !Mecanico
                mitmAux.SubItems(2) = IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA") & "(" & mstrNumeroDocumento & ")"
                mitmAux.SubItems(3) = Format(!Fecha_Emision, "dd/mm/yyyy")
                mitmAux.SubItems(4) = FormatoValor(!TotalHoras, "", 2)
                mitmAux.SubItems(5) = FormatoValor(!TotalHorasReales, "", 2)
                
                
                'busca horas reales de las actividades de los servicios.  //// vale hongo
'                gstrSql = "SELECT SUM(isnull(HorasReales,0)) AS SumaHoraActividades From TLLR_ACTIVIDADES_MECANICO "
'                gstrSql = gstrSql & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Mecanico='" & Me.dtcSupervisor.BoundText & "' "
'                gstrSql = gstrSql & "And Id_OT='" & !id_ot & "' And Seccion_OT='" & !Seccion_ot & "'"
'                If Conexion.SendHost(gstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'                    If Not adoTemp.BOF And Not adoTemp.EOF Then
'                        mitmAux.SubItems(4) = IIf(IsNull(adoTemp!SumaHoraActividades), FormatoValor(0, "", 2), FormatoValor(ValorNulo(adoTemp!SumaHoraActividades), "", 2))
'                    End If
'                End If

                .MoveNext
            Wend
        End If
        .Close
    End With
    
    'busca horas de actividades de servicios de otra ot
    gstrSql = "SELECT TLLR_ACTIVIDADES_MECANICO.Id_OT, TLLR_ACTIVIDADES_MECANICO.Seccion_Ot, TLLR_ACTIVIDADES_MECANICO.FechaEmision, "
    gstrSql = gstrSql & "Tllr_OT.Nro_Factura_Emitida, Tllr_Mecanicos.Nombre as Mecanico, SUM(TLLR_ACTIVIDADES_MECANICO.HorasReales) AS SumaHorasReales, "
    gstrSql = gstrSql & "SUM(TLLR_ACTIVIDADES_MECANICO.HorasActividad) As SumaHorasActividad "
    gstrSql = gstrSql & "FROM TLLR_ACTIVIDADES_MECANICO INNER JOIN "
    gstrSql = gstrSql & "Tllr_OT ON TLLR_ACTIVIDADES_MECANICO.Id_OT = Tllr_OT.Id_OT AND "
    gstrSql = gstrSql & "TLLR_ACTIVIDADES_MECANICO.Id_Empresa = Tllr_OT.Id_Empresa AND "
    gstrSql = gstrSql & "TLLR_ACTIVIDADES_MECANICO.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
    gstrSql = gstrSql & "TLLR_ACTIVIDADES_MECANICO.Seccion_OT = Tllr_OT.Seccion_OT "
    gstrSql = gstrSql & " INNER JOIN Tllr_Mecanicos on Tllr_Mecanicos.Id_Mecanico=Tllr_Actividades_Mecanico.Id_Mecanico "
    gstrSql = gstrSql & "Where Tllr_Actividades_Mecanico.Id_Empresa='" & gstrIdEmpresa & "' And Tllr_Actividades_Mecanico.Id_Sucursal='" & gstrIdSucursal & "' "
    'kjcv 15.08.16
    If dtcSupervisor.BoundText <> "" Then
    gstrSql = gstrSql & "And Tllr_Actividades_Mecanico.Id_Mecanico='" & dtcSupervisor.BoundText & "' "
    End If
    gstrSql = gstrSql & "GROUP BY TLLR_ACTIVIDADES_MECANICO.Id_OT, TLLR_ACTIVIDADES_MECANICO.Seccion_Ot, TLLR_ACTIVIDADES_MECANICO.FechaEmision, "
    gstrSql = gstrSql & "Tllr_OT.Nro_Factura_Emitida, Tllr_Mecanicos.Nombre"

    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With gadoPrincipal
            If Not .BOF And Not .EOF Then
                While Not .EOF
                    Set mitmAux = lvwMecanica.ListItems.Add(, , !Id_OT)
'                    Set mitmAux = lvwMecanica.ListItems.Add(, , !Mecanico)
'                    mitmAux.SubItems(1) = !Id_OT
                    mitmAux.SubItems(1) = !Mecanico
                    mitmAux.SubItems(2) = IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA") & "(" & ValorNulo(!Nro_Factura_Emitida) & ")"
                    mitmAux.SubItems(3) = Format(!FechaEmision, "dd/mm/yyyy")
                    mitmAux.SubItems(4) = IIf(IsNull(!SumaHorasactividad), FormatoValor(0, "", 2), FormatoValor(ValorNulo(!SumaHorasactividad), "", 2))
                    mitmAux.SubItems(5) = IIf(IsNull(!SumaHorasReales), FormatoValor(0, "", 2), FormatoValor(ValorNulo(!SumaHorasReales), "", 2))
                    .MoveNext
                Wend
            End If
        .Close
        End With
    End If

    
    
End If
lblTotalMec = TotalSeccion(lvwMecanica, 4)
lblTotalMec2 = TotalSeccion(lvwMecanica, 5)
lblTotalMec3 = TotalSeccion(lvwMecanica, 6)
End Sub

Sub Resumen()
Dim dblTotalHoras As Double
Dim dblTotalHorasAsignadas As Double
Dim lblPorProd2 As Double


dblTotalHoras = Val(lblTotalMec) + Val(lblTotalOtr) + Val(lblTotalCar)
dblTotalHorasAsignadas = Val(lblTotalMec3) + Val(lblTotalOtro3)

If lblDiasHabEst = "0" Then
    lblDiasHabEst = FormatoValor(NroDiasHabiles(CDate(pckFechaDesde.Value), CDate(pckFechaHasta.Value) & " 23:59:00"), "", 0)
End If
'lblHorEst = gdblNroHorOblg
lblTotHorEst = NroDiasHabiles(CDate(pckFechaDesde.Value), CDate(pckFechaHasta.Value) & " 23:59:59") * gdblNroHorOblg
'lblHorasReales = FormatoValor(dblTotalHoras, "", 1)
'lblPorProd = PorcentajeMonto((NroDiasHabiles(CDate(pckFechaDesde.Value), CDate(pckFechaHasta.Value) & " 23:59:59") * gdblNroHorOblg), CSng(dblTotalHoras))
'lblOtTrabajadas = lvwMecanica.ListItems.Count + lvwOtro.ListItems.Count + lvwDyP.ListItems.Count

lblHorEst = gdblNroHorOblg
'lblTotHorEst = CDbl(lblDiasHabEst) * gdblNroHorOblg
lblHorasReales = FormatoValor(dblTotalHoras, "", 2)
lblHorasAsignadas = FormatoValor(dblTotalHorasAsignadas, "", 2)
'lblPorProd = PorcentajeMonto((CDbl(Me.lblDiasHabEst) * gdblNroHorOblg), CSng(dblTotalHoras))


lblProdAsig = PorcentajeMonto((CDbl(Me.lblDiasHabEst) * gdblNroHorOblg), CSng(dblTotalHorasAsignadas))

lblPorProd = PorcentajeMonto((CDbl(Me.lblDiasHabEst) * gdblNroHorOblg), CSng(dblTotalHoras))



lblOtTrabajadas = lvwMecanica.ListItems.Count + lvwOtro.ListItems.Count + lvwDyP.ListItems.Count
lblHorasReales1 = CDbl(lblTotalMec2.Caption) + CDbl(lblTotalOtro2.Caption)

gstrSql = "SELECT Valor_Hora FROM Tllr_Mecanicos where Id_Mecanico='" & Me.dtcSupervisor.BoundText & "' And Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and vigencia='S'"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not gadoPrincipal.BOF And Not gadoPrincipal.EOF Then
        lblValorHora = FormatoValor(gadoPrincipal!Valor_Hora, "", gintDecimalesMoneda)
    End If
End If
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal

lblValorPesos = FormatoValor(CDbl(Me.lblValorHora) * CDbl(Me.lblHorasReales), "", gintDecimalesMoneda)

End Sub

Sub ResumenPorEstado(pstrIdMecanico As String, pdteFechaIni As Date, pdteFechaFin As Date, pstrEstado As String)


'gstrIdEmpresa & "') AND (Tllr_Mecanica_OT.Id_Sucursal = '" & gstrIdSucursal & "') AND (Tllr_Mecanica_OT.Mecanico_Designado = '" & pstrIdMecanico & "') And (Tllr_OT.Fecha_Emision Between '" & pdteFechaIni & "'  And  '" & pdteFechaFin & "')

gstrSql = "SELECT Id_OT, Seccion_OT, Mecanico_Designado, Estado, Fecha_Emision, SUM(TOTALHORAS) FROM"
gstrSql = gstrSql & " (   SELECT Tllr_Mecanica_OT.Id_OT, Tllr_Mecanica_OT.Seccion_OT, Tllr_Mecanica_OT.Mecanico_Designado,  Tllr_OT.Estado, Tllr_OT.Fecha_Emision ,  Sum(Tllr_Mecanica_OT.Horas) AS TOTALHORAS"
    gstrSql = gstrSql & " FROM Tllr_Mecanica_OT LEFT OUTER JOIN Tllr_OT ON Tllr_Mecanica_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Mecanica_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND Tllr_Mecanica_OT.Id_OT = Tllr_OT.Id_OT AND Tllr_Mecanica_OT.Seccion_OT = Tllr_OT.Seccion_OT"
    gstrSql = gstrSql & " GROUP BY Tllr_Mecanica_OT.Id_Empresa, Tllr_Mecanica_OT.Id_Sucursal, Tllr_Mecanica_OT.Id_OT, Tllr_Mecanica_OT.Seccion_OT, Tllr_OT.Estado, Tllr_Mecanica_OT.Mecanico_Designado, Tllr_OT.Fecha_Emision"
    gstrSql = gstrSql & " HAVING (Tllr_Mecanica_OT.Id_Empresa = '" & gstrIdEmpresa & "') AND (Tllr_Mecanica_OT.Id_Sucursal = '" & gstrIdSucursal & "') AND (Tllr_Mecanica_OT.Mecanico_Designado = '" & pstrIdMecanico & "') And (Tllr_OT.Fecha_Emision Between '" & pdteFechaIni & "'  And  '" & pdteFechaFin & "')"
gstrSql = gstrSql & " Union All"
    gstrSql = gstrSql & " SELECT Tllr_Otro_OT.Id_OT, Tllr_Otro_OT.Seccion_OT, Tllr_Otro_OT.Mecanico_Asignado,  Tllr_OT.Estado, Tllr_OT.Fecha_Emision ,  Sum(Tllr_Otro_OT.Horas) AS TOTALHORAS"
    gstrSql = gstrSql & " FROM Tllr_Otro_OT LEFT OUTER JOIN Tllr_OT ON Tllr_Otro_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Otro_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND Tllr_Otro_OT.Id_OT = Tllr_OT.Id_OT AND Tllr_Otro_OT.Seccion_OT = Tllr_OT.Seccion_OT"
    gstrSql = gstrSql & " GROUP BY Tllr_Otro_OT.Id_Empresa, Tllr_Otro_OT.Id_Sucursal, Tllr_Otro_OT.Id_OT, Tllr_Otro_OT.Seccion_OT, Tllr_OT.Estado, Tllr_Otro_OT.Mecanico_Asignado, Tllr_OT.Fecha_Emision"
    gstrSql = gstrSql & " HAVING (Tllr_Otro_OT.Id_Empresa = '" & gstrIdEmpresa & "') AND (Tllr_Otro_OT.Id_Sucursal = '" & gstrIdSucursal & "') AND (Tllr_Otro_OT.Mecanico_Asignado = '" & pstrIdMecanico & "') And (Tllr_OT.Fecha_Emision Between '" & pdteFechaIni & "'  And  '" & pdteFechaFin & "')"
gstrSql = gstrSql & " Union All"
    gstrSql = gstrSql & " SELECT Tllr_Carroceria_OT.Id_OT, Tllr_Carroceria_OT.Seccion_OT, Tllr_Carroceria_OT.Mecanico_Designado,  Tllr_OT.Estado, Tllr_OT.Fecha_Emision ,  Sum(Tllr_Carroceria_OT.Horas) AS TOTALHORAS"
    gstrSql = gstrSql & " FROM Tllr_Carroceria_OT LEFT OUTER JOIN Tllr_OT ON Tllr_Carroceria_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Carroceria_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND Tllr_Carroceria_OT.Id_OT = Tllr_OT.Id_OT AND Tllr_Carroceria_OT.Seccion_OT = Tllr_OT.Seccion_OT"
    gstrSql = gstrSql & " GROUP BY Tllr_Carroceria_OT.Id_Empresa, Tllr_Carroceria_OT.Id_Sucursal, Tllr_Carroceria_OT.Id_OT, Tllr_Carroceria_OT.Seccion_OT, Tllr_OT.Estado, Tllr_Carroceria_OT.Mecanico_Designado, Tllr_OT.Fecha_Emision"
    gstrSql = gstrSql & " HAVING (Tllr_Carroceria_OT.Id_Empresa = '" & gstrIdEmpresa & "') AND (Tllr_Carroceria_OT.Id_Sucursal = '" & gstrIdSucursal & "') AND (Tllr_Carroceria_OT.Mecanico_Designado = '" & pstrIdMecanico & "') And (Tllr_OT.Fecha_Emision Between '" & pdteFechaIni & "'  And  '" & pdteFechaFin & "'))"
gstrSql = gstrSql & " AS RESUMEN"
gstrSql = gstrSql & " GROUP BY Id_OT, Seccion_OT, Mecanico_Designado, Estado, Fecha_Emision, TOTALHORAS"
gstrSql = gstrSql & " HAVING (Estado = '" & pstrEstado & "')"

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
    
End If
End Sub

Function TotalSeccion(lvwObjeto As ListView, IndiceSubItem As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwObjeto
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
    Next
End With
TotalSeccion = dblPreSuma
End Function
Sub FillOtro(pstrIdMecanico As String, pdteFechaIni As Date, pdteFechaFin As Date)

    lvwOtro.ListItems.Clear
    Screen.MousePointer = 11
    
    '/// llama al procedimiento almacenado
'    gstrSql = "Exec Tllr_ProdMecanico_Otro " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & pstrIdMecanico & "','" & pdteFechaIni & "','" & pdteFechaFin & "','" & gstrEstadoProdMecanico & "'"
    
    gstrSql = "SELECT SUM(Tllr_Otro_OT.Horas) AS TOTALHORAS, Sum(isnull(Tllr_Otro_Ot.HorasReales,0)) as TotalHorasReales,SUM(isnull(Tllr_Otro_OT.HorasAsignadas,0))as TotalHorasAsignadas,"
    gstrSql = gstrSql & " Tllr_Facturacion.Id_Empresa, Tllr_Facturacion.Id_Sucursal,Tllr_Facturacion.Id_OT,"
    gstrSql = gstrSql & " Tllr_Facturacion.Seccion_OT, Tllr_Otro_OT.Mecanico_Asignado,Tllr_Mecanicos.Nombre as Mecanico, Tllr_Facturacion.Estado,Tllr_Facturacion.Id_Cargo,"
    gstrSql = gstrSql & " Tllr_Facturacion.Nro_Factura_Emitida As NUMDOCUMENTO, Tllr_Facturacion.Fecha_Facturacion As Fecha_Emision,"
    gstrSql = gstrSql & " Tllr_Facturacion.Fecha_Liquidacion FROM Tllr_Facturacion LEFT OUTER JOIN"
    gstrSql = gstrSql & " Tllr_Otro_OT ON Tllr_Facturacion.Id_Empresa = Tllr_Otro_OT.Id_Empresa AND"
    gstrSql = gstrSql & " Tllr_Facturacion.Id_Sucursal = Tllr_Otro_OT.Id_Sucursal AND Tllr_Facturacion.Id_OT = Tllr_Otro_OT.Id_OT AND"
    gstrSql = gstrSql & " Tllr_Facturacion.Seccion_OT = Tllr_Otro_OT.Seccion_OT AND"
    gstrSql = gstrSql & " Tllr_Facturacion.Id_Cargo = Tllr_Otro_OT.Id_Tipo_Cargo"
    'kjcv 12.08.16
    gstrSql = gstrSql & " INNER JOIN Tllr_Mecanicos on Tllr_Mecanicos.Id_Mecanico= Tllr_Otro_OT.Mecanico_Asignado"
    gstrSql = gstrSql & " GROUP BY Tllr_Facturacion.Id_Empresa, Tllr_Facturacion.Id_Sucursal, Tllr_Facturacion.Id_OT,"
    gstrSql = gstrSql & " Tllr_Facturacion.Seccion_OT, Tllr_Facturacion.Estado, Tllr_Otro_OT.Mecanico_Asignado,Tllr_Mecanicos.Nombre,"
    gstrSql = gstrSql & " Tllr_Facturacion.Nro_Factura_Emitida, Tllr_Facturacion.Fecha_Facturacion,"
    gstrSql = gstrSql & " Tllr_Facturacion.Fecha_Liquidacion,Tllr_Otro_Ot.Id_OT, Tllr_Otro_Ot.Id_Tipo_Cargo,"
    gstrSql = gstrSql & " Tllr_Facturacion.Id_Cargo"
    gstrSql = gstrSql & " HAVING Tllr_Facturacion.Id_Empresa = '" & gstrIdEmpresa & "'"
    gstrSql = gstrSql & " AND Tllr_Facturacion.Id_Sucursal = '" & gstrIdSucursal & "'"
    'kjcv 12.08.16
    If dtcSupervisor.BoundText <> "" Then
    gstrSql = gstrSql & " AND Tllr_Otro_OT.Mecanico_Asignado = '" & pstrIdMecanico & "'"
    End If
    gstrSql = gstrSql & " And Tllr_Facturacion.Id_Cargo = Tllr_Otro_Ot.Id_Tipo_Cargo"
    gstrSql = gstrSql & " AND Tllr_Facturacion.Fecha_Facturacion BETWEEN '" & pdteFechaIni & "' AND '" & pdteFechaFin & "'"
    If gstrEstadoProdMecanico = "F" Then
        gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B')"
    ElseIf gstrEstadoProdMecanico = "A" Then
        gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('F','B','V')"
    Else
        gstrSql = gstrSql & " AND Tllr_Facturacion.Estado IN ('V')"
    End If
    
    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With gadoPrincipal
       If Not .BOF And Not .EOF Then
          While Not .EOF

                If !estado = "B" Or !estado = "F" Then
                    mstrNumeroDocumento = ValorNulo(!NUMDOCUMENTO)
                Else
                    mstrNumeroDocumento = "S/N"
                End If
                
                Set mitmAux = lvwOtro.ListItems.Add(, , !Id_OT)
'                Set mitmAux = lvwOtro.ListItems.Add(, , !Mecanico)
                mitmAux.SubItems(1) = !Mecanico
                mitmAux.SubItems(2) = IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA") & "(" & mstrNumeroDocumento & ")"
                mitmAux.SubItems(3) = Format(!Fecha_Emision, "dd/mm/yyyy")
                mitmAux.SubItems(4) = FormatoValor(!TotalHoras, "", 2)
                mitmAux.SubItems(5) = FormatoValor(!TotalHorasReales, "", 2)
                mitmAux.SubItems(6) = FormatoValor(!TotalHorasAsignadas, "", 2)
                .MoveNext
            Wend
        End If
        .Close
    End With
End If
'lblTotalOtr = TotalSeccion(lvwOtro, 3)
'lblTotalOtro2 = TotalSeccion(lvwOtro, 4)
lblTotalOtr = TotalSeccion(lvwOtro, 4)
lblTotalOtro2 = TotalSeccion(lvwOtro, 5)
lblTotalOtro3 = TotalSeccion(lvwOtro, 6)

End Sub
Sub FillCarroceria(pstrIdMecanico As String, pdteFechaIni As Date, pdteFechaFin As Date)
lvwDyP.ListItems.Clear
'gstrSql = "SELECT Tllr_Carroceria_OT.Id_OT AS ID, Tllr_Carroceria_OT.Seccion_OT AS SEC, Tllr_Carroceria_OT.Mecanico_Designado AS MEC, SUM(Tllr_Carroceria_OT.Horas) AS SHORAS,Tllr_OT.Fecha_Emision AS FEC"
'gstrSql = gstrSql & " FROM Tllr_Carroceria_OT LEFT OUTER JOIN Tllr_OT ON Tllr_Carroceria_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Carroceria_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND Tllr_Carroceria_OT.Id_OT = Tllr_OT.Id_OT AND Tllr_Carroceria_OT.Seccion_OT = Tllr_OT.Seccion_OT "
'gstrSql = gstrSql & " WHERE (Tllr_Carroceria_OT.Mecanico_Designado ='" & pstrIdMecanico & "' And Tllr_OT.Fecha_Emision Between '" & pdteFechaIni & "' And '" & pdteFechaFin & "' )"
'gstrSql = gstrSql & " GROUP BY Tllr_Carroceria_OT.Mecanico_Designado, Tllr_Carroceria_OT.Id_OT, Tllr_Carroceria_OT.Seccion_OT, Tllr_OT.Fecha_Emision"

gstrSql = "SELECT Tllr_Carroceria_OT.Id_Empresa,"
gstrSql = gstrSql & " Tllr_Carroceria_OT.Id_Sucursal, "
gstrSql = gstrSql & " Tllr_Carroceria_OT.Id_OT,"
gstrSql = gstrSql & " Tllr_Carroceria_OT.Seccion_OT,"
gstrSql = gstrSql & " Tllr_Carroceria_OT.Mecanico_Designado, "
gstrSql = gstrSql & " Tllr_OT.Estado,"
gstrSql = gstrSql & " Tllr_OT.Fecha_Emision , Tllr_OT.Fecha_Liquidacion,"
gstrSql = gstrSql & " Sum(Tllr_Carroceria_OT.Horas) AS TOTALHORAS, Sum(Tllr_Carroceria_OT.HorasAsignadas) AS TOTALHORASASIGNADAS"
gstrSql = gstrSql & " FROM Tllr_Carroceria_OT LEFT OUTER JOIN Tllr_OT ON Tllr_Carroceria_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Carroceria_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND Tllr_Carroceria_OT.Id_OT = Tllr_OT.Id_OT AND Tllr_Carroceria_OT.Seccion_OT = Tllr_OT.Seccion_OT"
gstrSql = gstrSql & " GROUP BY Tllr_Carroceria_OT.Id_Empresa, Tllr_Carroceria_OT.Id_Sucursal, Tllr_Carroceria_OT.Id_OT, Tllr_Carroceria_OT.Seccion_OT, Tllr_OT.Estado, Tllr_Carroceria_OT.Mecanico_Designado, Tllr_OT.Fecha_Emision, Tllr_OT.Fecha_Liquidacion"
gstrSql = gstrSql & " HAVING (Tllr_Carroceria_OT.Id_Empresa = '" & gstrIdEmpresa & "') AND (Tllr_Carroceria_OT.Id_Sucursal = '" & gstrIdSucursal & "') AND (Tllr_Carroceria_OT.Mecanico_Designado = '" & pstrIdMecanico & "') And ((Tllr_OT.Fecha_Emision Between '" & pdteFechaIni & "'  And  '" & pdteFechaFin & "') or (Tllr_OT.Fecha_Liquidacion Between '" & pdteFechaIni & "'  And  '" & pdteFechaFin & "')) AND Tllr_OT.Estado " & gstrEstadoProdMecanico 'IN('L','B','F','C')"

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set mitmAux = lvwDyP.ListItems.Add(, , !Id_OT)
                mitmAux.SubItems(1) = IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA")
                mitmAux.SubItems(2) = Format(!Fecha_Emision, "dd/mm/yyyy")
                mitmAux.SubItems(3) = FormatoValor(!TotalHoras, "", 2)
                mitmAux.SubItems(4) = FormatoValor(!TotalHorasAsignadas, "", 2)
                .MoveNext
            Wend
        End If
        .Close
    End With
End If
lblTotalCar = TotalSeccion(lvwDyP, 3)
lblTotalCar3 = TotalSeccion(lvwDyP, 4)
End Sub
Private Sub cmdBuscar_Click()

Screen.MousePointer = 11
'If dtcSupervisor.BoundText <> "" Then
    FillMecanica dtcSupervisor.BoundText, CDate(pckFechaDesde.Value) & " 00:00:00", CDate(pckFechaHasta.Value) & " 23:59:00"
    FillOtro dtcSupervisor.BoundText, CDate(pckFechaDesde.Value) & " 00:00:00", CDate(pckFechaHasta.Value) & " 23:59:00"
    
    'FillCarroceria dtcSupervisor.BoundText, CDate(pckFechaDesde.Value) & " 00:00:00", CDate(pckFechaHasta.Value) & " 23:59:59"
    'ResumenPorEstado dtcSupervisor.BoundText, CDate(pckFechaDesde.Value) & " 00:00:00", CDate(pckFechaHasta.Value) & " 23:59:00", "V"
    
    Resumen
'End If
Screen.MousePointer = 1
End Sub

Private Sub cmdCambiaDiasHabiles_Click()
Screen.MousePointer = 1
frmPermisoDiasHabiles.Show 1

If NoEsLaPassword(gstrVerificacion, gstrMecanicoDiasHabiles) Then
    lblDiasHabEst = FormatoValor(gintDiasHabiles, "", 0)
    lblTotHorEst = gintDiasHabiles * gdblNroHorOblg
    lblPorProd = PorcentajeMonto((gintDiasHabiles * gdblNroHorOblg), CSng(Me.lblHorasReales))
    'lblPorProd2 = PorcentajeMonto((gintDiasHabiles * gdblNroHorOblg), CSng(Me.lblHorasReales))
    
Else
    MsgBox "Lo Siento, La passWord ingresada no es Correcta", vbExclamation, "Password"
End If
    
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdExcel_Click()
'Exporta Otro
If lvwOtro.ListItems.Count > 0 Then
    ExportarDatos Me.lvwOtro, Me.cdExportar, Me.hwnd
Else
    MsgBox "No existen datos en la lista Trabajos Adicionales"
End If
'Exporta Mecanica
If lvwMecanica.ListItems.Count > 0 Then
    ExportarDatos Me.lvwMecanica, Me.cdExportar, Me.hwnd
Else
    MsgBox "No existen datos en la lista Trabajos Mecánica"
End If

End Sub

Private Sub cmdImprimir_Click()
    Dim Dbsnueva As Database
    Dim Tabla As DAO.Recordset
    Dim i As Integer
    Dim GcamBaseTem As String
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Dim mstrNumeroDocumento As String
    
    'Devuelve la ruta del directorio Windows
    Dim rc As Long
    Dim WinPath As String
    WinPath = Space$(300)
    rc = GetWindowsDirectory(WinPath, 300)
    GcamBaseTem = Trim$(WinPath)
    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) + "\Temp"
    '---------------------------------------
        
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (OT TEXT,SECCION TEXT,FECHAOT TEXT,HORAS TEXT,HORASREALES TEXT,HORASASIGNADAS TEXT)"
    
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    
    Tabla.AddNew
    
    Tabla!OT = "MECANICA"
    
    Tabla.Update
    
    For i = 1 To Me.lvwMecanica.ListItems.Count
        Tabla.AddNew
        Set Me.lvwMecanica.SelectedItem = Me.lvwMecanica.ListItems(i)
        Tabla!OT = IIf(Me.lvwMecanica.ListItems(i) = "", " ", Me.lvwMecanica.ListItems(i))
        Tabla!Seccion = IIf(Me.lvwMecanica.SelectedItem.SubItems(2) = "", " ", Me.lvwMecanica.SelectedItem.SubItems(2))
        Tabla!FECHAOT = IIf(Me.lvwMecanica.SelectedItem.SubItems(3) = "", " ", Me.lvwMecanica.SelectedItem.SubItems(3))
        Tabla!Horas = IIf(Me.lvwMecanica.SelectedItem.SubItems(4) = "", " ", Me.lvwMecanica.SelectedItem.SubItems(4))
        Tabla!HorasReales = IIf(Me.lvwMecanica.SelectedItem.SubItems(5) = "", " ", Me.lvwMecanica.SelectedItem.SubItems(5))
        Tabla!HorasAsignadas = IIf(Me.lvwMecanica.SelectedItem.SubItems(6) = "", " ", Me.lvwMecanica.SelectedItem.SubItems(6))
        Tabla.Update
    Next i
   'Tabla.Close
   
    Tabla.AddNew
    
    Tabla!OT = "OTROS SERVICIOS"
    
    Tabla.Update
   
    For i = 1 To Me.lvwOtro.ListItems.Count
        Tabla.AddNew
        Set Me.lvwOtro.SelectedItem = Me.lvwOtro.ListItems(i)
        Tabla!OT = IIf(Me.lvwOtro.ListItems(i) = "", " ", Me.lvwOtro.ListItems(i))
        Tabla!Seccion = IIf(Me.lvwOtro.SelectedItem.SubItems(2) = "", " ", Me.lvwOtro.SelectedItem.SubItems(2))
        Tabla!FECHAOT = IIf(Me.lvwOtro.SelectedItem.SubItems(3) = "", " ", Me.lvwOtro.SelectedItem.SubItems(3))
        Tabla!Horas = IIf(Me.lvwOtro.SelectedItem.SubItems(4) = "", " ", Me.lvwOtro.SelectedItem.SubItems(4))
        Tabla!HorasReales = IIf(Me.lvwOtro.SelectedItem.SubItems(5) = "", " ", Me.lvwOtro.SelectedItem.SubItems(5))
        Tabla!HorasAsignadas = IIf(Me.lvwOtro.SelectedItem.SubItems(6) = "", " ", Me.lvwOtro.SelectedItem.SubItems(6))
        Tabla.Update
    Next i
   
    Tabla.AddNew
    
    Tabla!OT = "DYP"
    
    Tabla.Update
   
    For i = 1 To Me.lvwDyP.ListItems.Count
         Tabla.AddNew
         Set Me.lvwDyP.SelectedItem = Me.lvwDyP.ListItems(i)
         Tabla!OT = IIf(Me.lvwDyP.ListItems(i) = "", " ", Me.lvwDyP.ListItems(i))
         Tabla!Seccion = IIf(Me.lvwDyP.SelectedItem.SubItems(1) = "", " ", Me.lvwDyP.SelectedItem.SubItems(1))
         Tabla!FECHAOT = IIf(Me.lvwDyP.SelectedItem.SubItems(2) = "", " ", Me.lvwDyP.SelectedItem.SubItems(2))
         Tabla!Horas = IIf(Me.lvwDyP.SelectedItem.SubItems(3) = "", " ", Me.lvwDyP.SelectedItem.SubItems(3))
         Tabla.Update
     Next i
     
     Tabla.Close
   '=========================
Dbsnueva.Close
With rptProdMec
     '"//MODIFICADO POR FDO DIAZ EL 29/11/2000
    .ReportFileName = gstrPathReporte & "\PRODMEC2.RPT"
    .WindowState = crptMaximized
    .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
    .Destination = crptToWindow
    .Formulas(0) = "DESDE='" & Format(pckFechaDesde.Value, "dd/mm/yyyy") & "'"
    .Formulas(1) = "HASTA='" & Format(pckFechaHasta.Value, "dd/mm/yyyy") & "'"
    .Formulas(2) = "NBEMEC='" & dtcSupervisor.Text & "'"
    .Formulas(3) = "USUARIO='" & gstrIdUsuario & "'"
    .Formulas(4) = "HMEC=" & Val(lblTotalMec) & ""
    .Formulas(5) = "HOTR=" & Val(lblTotalOtr) & ""
    .Formulas(6) = "HDYP=" & Val(lblTotalCar) & ""
    .Formulas(7) = "DHABEST=" & Val(lblDiasHabEst) & ""
    .Formulas(8) = "HRADIAEST=" & Val(lblHorEst) & ""
    .Formulas(9) = "THRAEST=" & Val(lblTotHorEst) & ""
    .Formulas(10) = "THRAREAL=" & Val(lblHorasReales) & ""
    .Formulas(19) = "THRAASIG=" & Val(lblHorasAsignadas) & ""
    .Formulas(11) = "PORCPROD=" & Val(lblPorProd) & ""
    .Formulas(12) = "TITULO='REPORTE DE PRODUCTIVIDAD POR MECANICO'"
    .Formulas(13) = "RAZONSOCIAL='" & NombreEmpresa(gstrIdEmpresa) & "'"
    .Formulas(14) = "SUCURSAL='" & NombreSucursal(gstrIdEmpresa, gstrIdSucursal) & "'"
    .Formulas(15) = "DIRECCION='" & DireccionSucursal(gstrIdEmpresa, gstrIdSucursal) & "'"
    .Formulas(16) = "HorasReales='" & Me.lblHorasReales1 & "'"
    .Formulas(17) = "ValorHora='" & Me.lblValorHora & "'"
    .Formulas(18) = "MontoCancelar='" & Me.lblValorPesos & "'"
        
    .Action = True
End With
''Dbsnueva.Close
Screen.MousePointer = 1
End Sub

Private Sub cmdVerGestion_Click()
    frmGestionTaller.Show vbModal
End Sub



Private Sub Form_Activate()
If mblnSW Then

    If Not Atributos("Glbl", "Tllr_30_00120", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If

    FillMecanicos
    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    mblnSW = False
End If
End Sub
Private Sub Form_Load()
mblnSW = True
End Sub


