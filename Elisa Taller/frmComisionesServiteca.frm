VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmComisionesServiteca 
   Caption         =   "Informe Comisiones de Serviteca"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "frmComisionesServiteca.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbTotales 
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   7320
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1323
            MinWidth        =   353
            Text            =   "Registros"
            TextSave        =   "Registros"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
            MinWidth        =   1411
            Key             =   "Registros"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1138
            MinWidth        =   1147
            Text            =   "OTs"
            TextSave        =   "OTs"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1411
            MinWidth        =   1411
            Key             =   "Ots"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Text            =   "Total Servicios"
            TextSave        =   "Total Servicios"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2205
            MinWidth        =   2205
            Key             =   "Servicios"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Com. Mecánico"
            TextSave        =   "Com. Mecánico"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2205
            MinWidth        =   2205
            Key             =   "Mecanico"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Com. Recepcionista"
            TextSave        =   "Com. Recepcionista"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2205
            MinWidth        =   2205
            Key             =   "Recepcionista"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox OtrosCriterios 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   4080
      ScaleHeight     =   2025
      ScaleWidth      =   3345
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CheckBox chkBoleteadas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Boleteadas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkNulas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Nulas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkFacturadas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Facturadas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkLiquidadas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Liquidadas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkVigentes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Vigentes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.OptionButton opcAgrupaOT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agrupar por O.T."
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   -300
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton opcAgrupaServicio 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agrupar por Servicio"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   -480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dbcboSucursal 
         Bindings        =   "frmComisionesServiteca.frx":0442
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Sucursal"
         Text            =   "dbcboSucursal"
      End
      Begin MSAdodcLib.Adodc datSucursal 
         Height          =   270
         Left            =   120
         Top             =   1680
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
         Caption         =   "datSucursal"
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
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   390
         Index           =   3
         Left            =   3000
         TabIndex        =   33
         ToolTipText     =   "Limpiar Sucursal"
         Top             =   1635
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   120
         X2              =   2040
         Y1              =   0
         Y2              =   0
      End
   End
   Begin Crystal.CrystalReport rptOT 
      Left            =   4860
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      Begin VB.TextBox txtOtHasta 
         Height          =   315
         Left            =   9960
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtOtDesde 
         Height          =   315
         Left            =   9960
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   9720
         TabIndex        =   16
         Top             =   120
         Width           =   30
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   7680
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36880
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   7680
         TabIndex        =   13
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36880
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Index           =   0
         Left            =   5160
         TabIndex        =   9
         ToolTipText     =   "Limopiar Mecánico"
         Top             =   480
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
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   7440
         TabIndex        =   8
         Top             =   120
         Width           =   30
      End
      Begin MSDataListLib.DataCombo dbcboMecanico 
         Bindings        =   "frmComisionesServiteca.frx":045C
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Mecanico"
         Text            =   "dbcboMecanico"
      End
      Begin MSAdodcLib.Adodc datMecanico 
         Height          =   270
         Left            =   1920
         Top             =   720
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
         Caption         =   "datMecanico"
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
      Begin MSAdodcLib.Adodc datConceptoServicio 
         Height          =   270
         Left            =   240
         Top             =   1440
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
         Caption         =   "datConceptoServicio"
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
      Begin MSDataListLib.DataCombo dbCboConceptoServicio 
         Bindings        =   "frmComisionesServiteca.frx":0476
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Concepto_Servicio"
         Text            =   "dbCboConceptoServicio"
      End
      Begin MSAdodcLib.Adodc datServicio 
         Height          =   270
         Left            =   3960
         Top             =   1440
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
         Caption         =   "datServicio"
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
      Begin MSDataListLib.DataCombo dbCboServicio 
         Bindings        =   "frmComisionesServiteca.frx":0498
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Servicio"
         Text            =   "dbCboServicio"
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Index           =   1
         Left            =   3000
         TabIndex        =   10
         ToolTipText     =   "Limpiar Concepto Servicio"
         Top             =   1200
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
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Index           =   2
         Left            =   6840
         TabIndex        =   11
         Top             =   1200
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
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbOtrosCriterios 
         Height          =   330
         Left            =   5880
         TabIndex        =   22
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonWidth     =   2434
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Otros Criterios"
               Key             =   "Otros"
               Object.ToolTipText     =   "Selección de otros criterios"
               ImageKey        =   "Editar"
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "O.T. Hasta"
         Height          =   255
         Left            =   9960
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "O.T. Desde"
         Height          =   255
         Left            =   9960
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   7680
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Left            =   7680
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Servicio"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Servicio"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Mecánico"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   0
      Top             =   720
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
            Picture         =   "frmComisionesServiteca.frx":04B2
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":05C4
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":06D6
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":07E8
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":08FA
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":0A0C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":0B1E
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":0C30
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":0D42
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":0E54
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":0F66
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":1078
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":118A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":129C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":13AE
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":14C0
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":15D2
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":1A24
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":1E76
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":1F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":24CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComisionesServiteca.frx":27EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
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
            Object.Visible         =   0   'False
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro (Ctrl+G)"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
            ImageKey        =   "Cancelar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
            ImageKey        =   "Primero"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
            ImageKey        =   "Siguiente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Cerrar"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Liquidar"
            Object.ToolTipText     =   "Estados Orden de Trabajo"
            ImageIndex      =   17
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
   Begin MSComctlLib.ListView lsvComisiones 
      Height          =   4815
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   32
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "linea"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero O.T."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Monto Total O.T."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Mecanico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Concepto Servicio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Servicio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Descuento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Factor (%) Mecánico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Factor ($) Mecánico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Factor (%) Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Factor ($) Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Comisión Mecánico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Comisión Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "FechaO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "MontoOTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "ValorO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "DescuentoO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "TotalO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "FactorMec%O"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "FactorMec$O"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "FactorRec%O"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "FactorRec$O"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "ComMecO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "ComRecO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Estado OT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Text            =   "Nro. Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Sucursal"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmComisionesServiteca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Item As ListItem

Public gstrPrefijoSistema As String
Public gstrCodigoAcceso As String

Private Sub ImprimirConsulta()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
Dim lstrCriterios As String

'Devuelve la ruta del directorio Windows
Dim rc As Long
Dim WinPath As String
WinPath = Space$(300)
rc = GetWindowsDirectory(WinPath, 300)
GcamBaseTem = Trim$(WinPath)
GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
'---------------------------------------

If lsvComisiones.ListItems.Count = 0 Then
    MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
    Exit Sub
End If

Screen.MousePointer = 11
Dim wrkPredeterminado As Workspace
Dim prpBucle As Property
Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NUMERO_OT DOUBLE, FECHA DATE, MONTO_OT DOUBLE, MECANICO TEXT, RECEPCIONISTA TEXT, CONCEPTO_SERV TEXT, SERVICIO TEXT, VALOR DOUBLE, CANTIDAD DOUBLE, DESCUENTO DOUBLE, TOTAL DOUBLE, FACTORPOR_MEC DOUBLE, FACTORPES_MEC DOUBLE, FACTORPOR_REC DOUBLE, FACTORPES_REC DOUBLE, COMISION_MEC DOUBLE, COMISION_REC DOUBLE, ESTADO TEXT, NUM_DOC TEXT)"
Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
For i = 1 To Me.lsvComisiones.ListItems.Count
    Tabla.AddNew
    Set lsvComisiones.SelectedItem = lsvComisiones.ListItems(i)
    Tabla!Numero_ot = IIf(lsvComisiones.SelectedItem.SubItems(1) = "", " ", lsvComisiones.SelectedItem.SubItems(1))
    Tabla!Fecha = IIf(lsvComisiones.SelectedItem.SubItems(2) = "", " ", lsvComisiones.SelectedItem.SubItems(2))
    Tabla!Monto_ot = IIf(lsvComisiones.SelectedItem.SubItems(3) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(3), gstrMonedaLocal)))
    Tabla!Mecanico = IIf(lsvComisiones.SelectedItem.SubItems(4) = "", " ", lsvComisiones.SelectedItem.SubItems(4))
    Tabla!Recepcionista = IIf(lsvComisiones.SelectedItem.SubItems(5) = "", " ", lsvComisiones.SelectedItem.SubItems(5))
    Tabla!CONCEPTO_SERV = IIf(lsvComisiones.SelectedItem.SubItems(6) = "", " ", lsvComisiones.SelectedItem.SubItems(6))
    Tabla!servicio = IIf(lsvComisiones.SelectedItem.SubItems(7) = "", " ", lsvComisiones.SelectedItem.SubItems(7))
    Tabla!Valor = IIf(lsvComisiones.SelectedItem.SubItems(8) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(8), gstrMonedaLocal)))
    Tabla!cantidad = IIf(lsvComisiones.SelectedItem.SubItems(9) = "", 0, CDbl(lsvComisiones.SelectedItem.SubItems(9)))
    Tabla!Descuento = IIf(lsvComisiones.SelectedItem.SubItems(10) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(10), "%")))
    Tabla!Total = IIf(lsvComisiones.SelectedItem.SubItems(11) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(11), gstrMonedaLocal)))
    Tabla!Factorpor_mec = IIf(lsvComisiones.SelectedItem.SubItems(12) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(12), "%")))
    Tabla!Factorpes_mec = IIf(lsvComisiones.SelectedItem.SubItems(13) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(13), gstrMonedaLocal)))
    Tabla!Factorpor_rec = IIf(lsvComisiones.SelectedItem.SubItems(14) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(14), "%")))
    Tabla!Factorpes_rec = IIf(lsvComisiones.SelectedItem.SubItems(15) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(15), gstrMonedaLocal)))
    Tabla!Comision_mec = IIf(lsvComisiones.SelectedItem.SubItems(16) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(16), gstrMonedaLocal)))
    Tabla!Comision_rec = IIf(lsvComisiones.SelectedItem.SubItems(17) = "", 0, CDbl(SacarFormatoValor(lsvComisiones.SelectedItem.SubItems(17), gstrMonedaLocal)))
    Tabla!estado = IIf(lsvComisiones.SelectedItem.SubItems(29) = "", " ", Mid$(lsvComisiones.SelectedItem.SubItems(29), 1, 1))
    Tabla!Num_Doc = IIf(lsvComisiones.SelectedItem.SubItems(30) = "", " ", lsvComisiones.SelectedItem.SubItems(30))
    Tabla.Update
Next i
Tabla.Close
   
With rptOT
    .ReportFileName = gstrPathReporte & "\COMISIONES.RPT"
    .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
    .Formulas(1) = "TITULO='COMISIONES SERVITECA'"
    .Formulas(2) = "RazonSocial='" & gstrEmpresa & "'"
    .Formulas(3) = "SUCURSAL='" & gstrSucursal & "'"
    .Formulas(4) = "DIRECCION='" & gstrDirSuc & "'"
    .Formulas(5) = "Registros='" & Me.stbTotales.Panels("Registros").Text & "'"
    .Formulas(6) = "OTs='" & Me.stbTotales.Panels(4).Text & "'"
    .Formulas(7) = "Maestro='" & IIf(Me.dbcboMecanico.Text <> "", Me.dbcboMecanico.Text, "TODOS") & "'"
    .Formulas(8) = "ConceptoServicio='" & IIf(Me.dbCboConceptoServicio.Text <> "", Me.dbCboConceptoServicio.Text, "TODOS") & "'"
    .Formulas(9) = "Servicio='" & IIf(Me.dbCboServicio.Text <> "", Me.dbCboServicio.Text, "TODOS") & "'"
    .Formulas(10) = "Rango='" & "Desde el " & Me.dtpDesde.Value & " Hasta el " & Me.dtpHasta.Value & "'"
    .Formulas(11) = "Desde='" & Me.txtOtDesde.Text & "'"
    .Formulas(12) = "Hasta='" & Me.txtOtHasta.Text & "'"
    lstrCriterios = "Criterios: Ordenes "
    If Me.chkVigentes.Value = vbChecked Then
        lstrCriterios = lstrCriterios & IIf(lstrCriterios <> "Criterios: Ordenes ", ", ", " ") & "VIGENTES"
    End If
    If Me.chkLiquidadas.Value = vbChecked Then
        lstrCriterios = lstrCriterios & IIf(lstrCriterios <> "Criterios: Ordenes ", ", ", " ") & "LIQUIDADAS"
    End If
    If Me.chkFacturadas.Value = vbChecked Then
        lstrCriterios = lstrCriterios & IIf(lstrCriterios <> "Criterios: Ordenes ", ", ", " ") & "FACTURADAS"
    End If
    If Me.chkBoleteadas.Value = vbChecked Then
        lstrCriterios = lstrCriterios & IIf(lstrCriterios <> "Criterios: Ordenes ", ", ", " ") & "BOLETEADAS"
    End If
    If Me.chkNulas.Value = vbChecked Then
        lstrCriterios = lstrCriterios & IIf(lstrCriterios <> "Criterios: Ordenes ", ", ", " ") & "NULAS"
    End If
    .Formulas(13) = "Criterios='" & lstrCriterios & "'"
    .Destination = crptToWindow
    .Action = True
End With
End Sub

Private Sub dbCboConceptoServicio_Click(Area As Integer)
If Area = 2 Then
    LLena_Servicio
    Me.lsvComisiones.ListItems.Clear
    LimpiaTotales
End If
End Sub


Private Function TraeMecanico(CodMecanico As String) As String
Dim tablaMecanico As New ADODB.Recordset
Dim lsql As String

Set tablaMecanico = New ADODB.Recordset
lsql = ""
lsql = "SELECT Nombre FROM Tllr_Mecanicos WHERE Id_Mecanico = '" & CodMecanico & "'"
If Conexion.SendHost(lsql, tablaMecanico, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaMecanico.EOF = False And tablaMecanico.BOF = False Then
        TraeMecanico = tablaMecanico!Nombre
    Else
        TraeMecanico = "."
    End If
End If
Conexion.CloseHost tablaMecanico

End Function

Private Sub Buscar()
Dim Tabla As New ADODB.Recordset
Dim tabla2 As New ADODB.Recordset
Dim sql As String
Dim ldblTotalNetoOT As Double
Dim ldblComisionMecanico As Double
Dim ldblComisionRecepcionista As Double
Dim ldblTotItem As Double
Dim ldblTotTotal As Double
Dim lstrConexion As String
Dim ldblTotalServicios As Double
Dim ldblCont As Double
Dim ldblOTs As Double

If Me.dtpDesde.Value > Me.dtpHasta.Value Then
    MsgBox "Los valores de fecha son incongruentes." & Chr(13) & "Asegurese de que la fecha DESDE sea inferior o igual a la fecha HASTA.", vbExclamation + vbOKOnly, "Rango de Fechas Inválido"
    Me.dtpDesde.SetFocus
    Exit Sub
End If

If Trim$(Me.txtOtDesde.Text) <> "" And Trim$(Me.txtOtHasta.Text) <> "" Then
    If CDbl(Me.txtOtDesde.Text) > CDbl(Me.txtOtHasta.Text) Then
        MsgBox "Los valores para búsqueda de OTs específicas es incongruente." & Chr(13) & "Asegurese de que la OT DESDE sea inferior o igual a la OT HASTA.", vbExclamation + vbOKOnly, "Rango de Fechas Inválido"
        Me.txtOtDesde.SetFocus
        Exit Sub
    End If
End If

If Trim$(Me.txtOtDesde.Text) <> "" And Trim$(Me.txtOtHasta.Text) = "" Then
    Me.txtOtHasta.Text = "9999999"
End If

If Trim$(Me.txtOtDesde.Text) = "" And Trim$(Me.txtOtHasta.Text) <> "" Then
    Me.txtOtDesde.Text = "0"
End If

ldblTotItem = 0
ldblTotTotal = 0
ldblTotalNetoOT = 0
ldblComisionMecanico = 0
ldblComisionRecepcionista = 0
ldblTotalServicios = 0
ldblOTs = 0

Me.lsvComisiones.ListItems.Clear
LimpiaTotales

lstrConexion = LetConnectionString("TLLR", "DSN", "AUTOPRO", 256)

Set apConexion = New APCONADO.ConnectionAdo
Set adoConexion = New ADODB.Connection
If apConexion.ConnectHost(adoConexion, adUseClient, lstrConexion, gcTiempoEspera, "c:\windows") <> apOk Then
    MsgBox "La conexión al origen de datos fue cancelada...", vbCritical, "Error"
    End
End If

ProcesoRegistros gcInicioProceso
Me.Refresh
ProcesoRegistros gcAvanceProceso, 20
sql = ""
If Me.opcAgrupaServicio.Value = True Then
    sql = sql & "SELECT Srvt_Servicios_OT.Id_OT AS OT, "
    sql = sql & "Srvt_OT.Fecha_Apertura AS Fecha, "
    sql = sql & "Srvt_OT.Valor_OT AS Total_Ot, "
    sql = sql & "Srvt_OT.Id_Sucursal, "
    sql = sql & "Tllr_Mecanicos.Nombre AS Mecanico, "
    sql = sql & "Srvt_OT.Id_Mecanico AS Cod_Recepcionista, "
    sql = sql & "Srvt_Servicios_OT.Id_Concepto_Servicio AS Cod_Concepto, "
    sql = sql & "Srvt_Concepto_Servicio.Descripcion AS Concepto, "
    sql = sql & "Srvt_Servicios_OT.Id_Servicio AS Cod_Servicio, "
    sql = sql & "Srvt_Servicios.Descripcion AS Servicio, "
    sql = sql & "Srvt_Servicios_OT.Valor AS Valor_Servicio_En_OT, "
    sql = sql & "Srvt_Servicios_OT.Cantidad AS Cant_Servicio_En_OT, "
    sql = sql & "Srvt_Servicios_OT.Descuento AS Desc_Servicio_En_OT, "
    sql = sql & "Srvt_Servicios_OT.Total AS Total_Servicio_En_OT, "
    sql = sql & "Srvt_Mecanico_Factor.Factor_Monto AS F_Monto_Mec, "
    sql = sql & "Srvt_Mecanico_Factor.Factor_Porcentaje AS F_Pje_Mec, "
    sql = sql & "Factor_Recepcionista.Factor_Monto AS F_Monto_Recep, "
    sql = sql & "Factor_Recepcionista.Factor_Porcentaje AS F_Pje_Recep, "
    sql = sql & "Copia_Mecanicos.Nombre AS Recepcionista, "
    sql = sql & "Srvt_Servicios.Factor_Porcentaje AS F_Pje_Srv, "
    sql = sql & "Srvt_Servicios.Factor_Monto AS F_Monto_Srv, "
    sql = sql & "Srvt_OT.Estado "
    sql = sql & "FROM Srvt_Mecanico_Factor RIGHT OUTER JOIN "
    sql = sql & "Factor_Recepcionista RIGHT OUTER JOIN "
    sql = sql & "Srvt_Concepto_Servicio RIGHT OUTER JOIN "
    sql = sql & "Srvt_Servicios ON "
    sql = sql & "Srvt_Concepto_Servicio.Id_Concepto_Servicio = Srvt_Servicios.Id_Concepto_Servicio "
    sql = sql & "RIGHT OUTER JOIN "
    sql = sql & "Srvt_Servicios_OT ON "
    sql = sql & "Srvt_Servicios.Id_Concepto_Servicio = Srvt_Servicios_OT.Id_Concepto_Servicio AND "
    sql = sql & "Srvt_Servicios.Id_Servicio = Srvt_Servicios_OT.Id_Servicio LEFT Outer Join "
    sql = sql & "Srvt_OT LEFT OUTER JOIN "
    sql = sql & "Copia_Mecanicos ON "
    sql = sql & "Srvt_OT.Id_Mecanico = Copia_Mecanicos.Id_Mecanico ON "
    sql = sql & "Srvt_Servicios_OT.Id_OT = Srvt_OT.Id_OT AND "
    sql = sql & "Srvt_Servicios_OT.Id_Sucursal = Srvt_OT.Id_Sucursal AND "
    sql = sql & "Srvt_Servicios_OT.Id_Empresa = Srvt_OT.Id_Empresa ON "
    sql = sql & "Factor_Recepcionista.Id_Mecanico = Srvt_OT.Id_Mecanico ON "
    sql = sql & "Srvt_Mecanico_Factor.Id_Mecanico = Srvt_Servicios_OT.Id_Mecanico "
    sql = sql & "LEFT OUTER JOIN Tllr_Mecanicos ON "
    sql = sql & "Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
Else
    sql = "SELECT Srvt_OT.Fecha_Apertura AS Fecha, "
    sql = sql & "Tllr_Mecanicos.Nombre AS Recepcionista, "
    sql = sql & "Srvt_OT.Id_Mecanico AS Cod_Recepcionista, Srvt_OT.Estado, "
    sql = sql & "Srvt_Mecanico_Factor.Factor_Porcentaje AS F_Pje_Recep, "
    sql = sql & "Srvt_OT.Id_Sucursal, "
    sql = sql & "Srvt_Mecanico_Factor.Factor_Monto AS F_Monto_Recep, "
    sql = sql & "Tllr_Mecanicos.Vigencia, "
    sql = sql & "Srvt_OT.Valor_OT AS Monto_Neto_Ot, "
    sql = sql & "Vpro_Facturacion.Estado_Facturacion, "
    sql = sql & "Vpro_Facturacion.Tipo_Docto, "
    sql = sql & "Vpro_Facturacion.Id_Tipo_Rescate , Srvt_OT.Id_OT AS OT "
    sql = sql & "FROM Srvt_Mecanico_Factor RIGHT OUTER JOIN "
    sql = sql & "Tllr_Mecanicos RIGHT OUTER JOIN "
    sql = sql & "Srvt_OT LEFT OUTER JOIN "
    sql = sql & "Vpro_Facturacion ON CONVERT(nvarchar, Srvt_OT.Id_OT) "
    sql = sql & "= Vpro_Facturacion.Numero_Rescate ON "
    sql = sql & "Tllr_Mecanicos.Id_Mecanico = Srvt_OT.Id_Mecanico ON "
    sql = sql & "Srvt_Mecanico_Factor.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
End If

sql = sql & "WHERE Tllr_Mecanicos.Vigencia = 'S' "

If Me.dbcboSucursal.Text <> "" Then
    sql = sql & "AND (Srvt_OT.Id_Empresa = '" & gstrIdEmpresa & "' AND Srvt_OT.Id_Sucursal = '" & Me.dbcboSucursal.BoundText & "') "
End If

sql = sql & "AND (Srvt_OT.Estado = '.' "

If Me.chkLiquidadas.Value = 1 Then
    sql = sql & "OR Srvt_OT.Estado = 'L' "
End If

If Me.chkVigentes.Value = 1 Then
    sql = sql & "OR Srvt_OT.Estado = 'V' "
End If

If Me.chkFacturadas.Value = 1 Then
    sql = sql & "OR Srvt_OT.Estado = 'F' "
End If

If Me.chkBoleteadas.Value = 1 Then
    sql = sql & "OR Srvt_OT.Estado = 'B' "
End If

If Me.chkNulas.Value = 1 Then
    sql = sql & "OR Srvt_OT.Estado = 'N' "
End If

sql = sql & ") AND (Srvt_OT.Fecha_Apertura BETWEEN '" & Me.dtpDesde.Value & " 00:00:01' AND '" & Me.dtpHasta.Value & " 23:59:00') "

If Me.dbcboMecanico.Text <> "" Then
    If Me.opcAgrupaServicio.Value = True Then
        sql = sql & "AND (Srvt_Servicios_OT.Id_Mecanico = '" & Me.dbcboMecanico.BoundText & "' OR Srvt_OT.Id_Mecanico = '" & Me.dbcboMecanico.BoundText & "') "
    Else
        sql = sql & "AND Srvt_OT.Id_Mecanico = '" & Me.dbcboMecanico.BoundText & "' "
    End If
End If
If Me.dbCboConceptoServicio.Text <> "" Then
    sql = sql & "AND Srvt_Servicios_OT.Id_Concepto_Servicio = '" & Me.dbCboConceptoServicio.BoundText & "' "
End If
If Me.dbCboServicio.Text <> "" Then
    sql = sql & "AND Srvt_Servicios_OT.Id_Servicio = '" & Me.dbCboServicio.BoundText & "' "
End If
If Me.txtOtDesde.Text <> "" And Me.txtOtHasta.Text <> "" Then
    sql = sql & "AND (Srvt_Servicios_OT.Id_OT BETWEEN '" & Me.txtOtDesde.Text & "' AND '" & Me.txtOtHasta.Text & "') "
End If

sql = sql & " ORDER BY Srvt_Servicios_OT.Id_OT"

If Conexion.SendHost(sql, Tabla, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then
        Tabla.MoveFirst
        While Tabla.EOF = False
            Set Item = Me.lsvComisiones.ListItems.Add(, , Me.lsvComisiones.ListItems.Count + 1)
            Item.SubItems(1) = ValorNulo(Tabla!OT)
            If Me.lsvComisiones.ListItems.Count = 1 Then
                ldblOTs = ldblOTs + 1
            Else
                If Item.SubItems(1) <> Me.lsvComisiones.ListItems(Me.lsvComisiones.ListItems.Count - 1).SubItems(1) Then
                    ldblOTs = ldblOTs + 1
                End If
            End If
            Item.SubItems(2) = Format$(ValorNulo(Tabla!Fecha), "dd/MM/yyyy")
            sql = ""
            sql = sql & "SELECT Srvt_Servicios_OT.Id_OT, "
            sql = sql & "Fact_Con_Detalle.Despachado AS Cant_Fac, "
            sql = sql & "Fact_Con_Detalle.Precio_Venta AS Valor_Fac, "
            sql = sql & "Fact_Con_Detalle.Pje_Descto AS Desc_Fac, "
            sql = sql & "Fact_Con_Detalle.Total AS Total_Fac, "
            sql = sql & "Srvt_Servicios_OT.Id_Servicio, "
            sql = sql & "Fact_Con_Detalle.GRAN_TOTAL, "
            sql = sql & "Fact_Con_Detalle.Tipo_Docto, "
            sql = sql & "Fact_Con_Detalle.Numero_Documento "
            sql = sql & "FROM Srvt_Servicios_OT LEFT OUTER JOIN "
            sql = sql & "Fact_Con_Detalle ON CONVERT(nvarchar (30), "
            sql = sql & "Srvt_Servicios_OT.Id_OT) = CONVERT(nvarchar (30), "
            sql = sql & "Fact_Con_Detalle.Numero_Rescate) AND "
            sql = sql & "CONVERT(nvarchar (30), Srvt_Servicios_OT.Id_Servicio) "
            sql = sql & "= SUBSTRING (CONVERT(nvarchar (30), "
            sql = sql & "Fact_Con_Detalle.Id_Item), 5, len (Fact_Con_Detalle.Id_Item)) AND "
            sql = sql & "Srvt_Servicios_OT.Id_Empresa = Fact_Con_Detalle.Id_Empresa AND "
            sql = sql & "Srvt_Servicios_OT.Id_Sucursal = Fact_Con_Detalle.Id_Sucursal "
            sql = sql & "WHERE Srvt_Servicios_OT.Id_OT = " & Item.SubItems(1) & " "
            sql = sql & "AND Srvt_Servicios_OT.Id_Servicio = '" & ValorNulo(Tabla!Cod_Servicio) & "'"
            Set tabla2 = New ADODB.Recordset
            If apConexion.SendHost(sql, tabla2, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
                If (tabla2.EOF = False And tabla2.BOF = False) And Not IsNull(tabla2!Gran_Total) Then
                    tabla2.MoveFirst
                    While ValorNuloNum(tabla2!Cant_Fac) <> ValorNuloNum(Tabla!Cant_Servicio_En_OT) And tabla2.EOF = False And tabla2.RecordCount > 1
                        tabla2.MoveNext
                    Wend
                    If ValorNulo(tabla2!Tipo_Docto) = "BV" Then
                        Item.SubItems(3) = FormatoValor(ValorNuloNum(tabla2!Gran_Total) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), gstrMonedaLocal, gintDecimalesMoneda)
                        Item.SubItems(8) = FormatoValor(IIf(ValorNuloNum(tabla2!Valor_Fac) > 0, ValorNuloNum(tabla2!Valor_Fac) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), 0), gstrMonedaLocal, gintDecimalesMoneda)
                        Item.SubItems(9) = ValorNuloNum(tabla2!Cant_Fac)
                        Item.SubItems(10) = FormatoValor(ValorNuloNum(tabla2!Desc_Fac), "%", 2)
                        Item.SubItems(11) = FormatoValor(IIf(ValorNuloNum(tabla2!Total_Fac) > 0, ValorNuloNum(tabla2!Total_Fac) / IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), 0), gstrMonedaLocal, gintDecimalesMoneda)
                    Else
                        Item.SubItems(3) = FormatoValor(ValorNuloNum(tabla2!Gran_Total), gstrMonedaLocal, gintDecimalesMoneda)
                        Item.SubItems(8) = FormatoValor(ValorNuloNum(tabla2!Valor_Fac), gstrMonedaLocal, gintDecimalesMoneda)
                        Item.SubItems(9) = ValorNuloNum(tabla2!Cant_Fac)
                        Item.SubItems(10) = FormatoValor(ValorNuloNum(tabla2!Desc_Fac), "%", 2)
                        Item.SubItems(11) = FormatoValor(ValorNuloNum(tabla2!Total_Fac), gstrMonedaLocal, gintDecimalesMoneda)
                    End If
                    Item.SubItems(30) = IIf(ValorNulo(tabla2!Numero_Documento) <> "", ValorNulo(tabla2!Numero_Documento), "s/d")
                Else
                    Item.SubItems(3) = FormatoValor(ValorNuloNum(Tabla!Total_Ot), gstrMonedaLocal, gintDecimalesMoneda)
                    Item.SubItems(8) = FormatoValor(ValorNuloNum(Tabla!Valor_Servicio_En_OT), gstrMonedaLocal, gintDecimalesMoneda)
                    Item.SubItems(9) = ValorNuloNum(Tabla!Cant_Servicio_En_OT)
                    Item.SubItems(10) = FormatoValor(ValorNuloNum(Tabla!Desc_Servicio_En_OT), "%", 2)
                    Item.SubItems(11) = FormatoValor(ValorNuloNum(Tabla!Total_Servicio_En_OT), gstrMonedaLocal, gintDecimalesMoneda)
                    Item.SubItems(30) = "?"
                End If
            End If
            apConexion.CloseHost tabla2
            ldblTotItem = SacarFormatoValor(Item.SubItems(11), gstrMonedaLocal)
            ldblTotTotal = SacarFormatoValor(Item.SubItems(3), gstrMonedaLocal)
            ldblTotalNetoOT = ldblTotalNetoOT + SacarFormatoValor(Item.SubItems(3), gstrMonedaLocal)
            If Me.opcAgrupaServicio.Value = True Then
                Item.SubItems(4) = ValorNulo(Tabla!Mecanico)
            Else
                Item.SubItems(4) = ""
            End If
            Item.SubItems(5) = ValorNulo(Tabla!Recepcionista)
            If Me.opcAgrupaServicio.Value = True Then
                Item.SubItems(6) = ValorNulo(Tabla!Concepto)
                Item.SubItems(7) = ValorNulo(Tabla!servicio)
                If ValorNuloNum(Tabla!F_Pje_Mec) > 0 Or ValorNuloNum(Tabla!F_Monto_Mec) > 0 Then
                    Item.SubItems(12) = FormatoValor(ValorNuloNum(Tabla!F_Pje_Mec), "%", 2)
                    Item.SubItems(13) = FormatoValor(ValorNuloNum(Tabla!F_Monto_Mec), gstrMonedaLocal, gintDecimalesMoneda)
                Else
                    Item.SubItems(12) = FormatoValor(ValorNuloNum(Tabla!F_Pje_Srv), "%", 2)
                    Item.SubItems(13) = FormatoValor(ValorNuloNum(Tabla!F_Monto_Srv), gstrMonedaLocal, gintDecimalesMoneda)
                End If
            Else
                Item.SubItems(6) = ""
                Item.SubItems(7) = ""
                Item.SubItems(8) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                Item.SubItems(9) = "0"
                Item.SubItems(10) = FormatoValor(0, "%", 2)
                Item.SubItems(11) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                Item.SubItems(12) = FormatoValor(0, "%", 2)
                Item.SubItems(13) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
            End If
            If ValorNuloNum(Tabla!F_Pje_Recep) > 0 Or ValorNuloNum(Tabla!F_Monto_Recep) > 0 Then
                Item.SubItems(14) = FormatoValor(ValorNuloNum(Tabla!F_Pje_Recep), "%", 2)
                Item.SubItems(15) = FormatoValor(ValorNuloNum(Tabla!F_Monto_Recep), gstrMonedaLocal, gintDecimalesMoneda)
            Else
                Item.SubItems(14) = FormatoValor(ValorNuloNum(Tabla!F_Pje_Srv), "%", 2)
                Item.SubItems(15) = FormatoValor(ValorNuloNum(Tabla!F_Monto_Srv), gstrMonedaLocal, gintDecimalesMoneda)
            End If
            If Me.opcAgrupaServicio.Value = True Then
                If CDbl(SacarFormatoValor(Item.SubItems(12), "%")) <> 0 Then
                    Item.SubItems(16) = FormatoValor(ldblTotItem * (SacarFormatoValor(Item.SubItems(12), "%") / 100), gstrMonedaLocal, gintDecimalesMoneda)
                Else
                    If CDbl(SacarFormatoValor(Item.SubItems(13), "%")) <> 0 Then
                        Item.SubItems(16) = FormatoValor(SacarFormatoValor(Item.SubItems(13), "%") * Item.SubItems(9), gstrMonedaLocal, gintDecimalesMoneda)
                    Else
                        Item.SubItems(16) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                    End If
                End If
                If CDbl(SacarFormatoValor(Item.SubItems(14), "%")) <> 0 Then
                    Item.SubItems(17) = FormatoValor(ldblTotItem * (SacarFormatoValor(Item.SubItems(14), "%") / 100), gstrMonedaLocal, gintDecimalesMoneda)
                Else
                    If CDbl(SacarFormatoValor(Item.SubItems(15), "%")) <> 0 Then
                        Item.SubItems(17) = FormatoValor(SacarFormatoValor(Item.SubItems(15), "%"), gstrMonedaLocal, gintDecimalesMoneda)
                    Else
                        Item.SubItems(17) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                    End If
                End If
            End If
            Item.ListSubItems(16).Bold = True
            Item.ListSubItems(17).Bold = True
            Item.SubItems(18) = Format$(ValorNulo(Tabla!Fecha), "yyyy/MM/dd")
            Item.SubItems(19) = CDbl(SacarFormatoValor(Item.SubItems(3), gstrMonedaLocal))
            Item.SubItems(20) = SacarFormatoValor(Item.SubItems(8), gstrMonedaLocal)
            Item.SubItems(21) = SacarFormatoValor(Item.SubItems(10), "%")
            Item.SubItems(22) = SacarFormatoValor(Item.SubItems(11), gstrMonedaLocal)
            Item.SubItems(23) = SacarFormatoValor(Item.SubItems(12), "%")
            Item.SubItems(24) = SacarFormatoValor(Item.SubItems(13), gstrMonedaLocal)
            Item.SubItems(25) = SacarFormatoValor(Item.SubItems(14), "%")
            Item.SubItems(26) = SacarFormatoValor(Item.SubItems(15), gstrMonedaLocal)
            Item.SubItems(27) = SacarFormatoValor(Item.SubItems(16), gstrMonedaLocal)
            Item.SubItems(28) = SacarFormatoValor(Item.SubItems(17), gstrMonedaLocal)
            Item.SubItems(29) = IIf(ValorNulo(Tabla!estado) = "F", "FACTURADA", IIf(ValorNulo(Tabla!estado) = "B", "BOLETEADA", IIf(ValorNulo(Tabla!estado) = "N", "NULA", IIf(ValorNulo(Tabla!estado) = "L", "LIQUIDADA", ""))))
            ldblTotalServicios = ldblTotalServicios + SacarFormatoValor(Item.SubItems(11), gstrMonedaLocal)
            ldblComisionMecanico = ldblComisionMecanico + SacarFormatoValor(Item.SubItems(16), gstrMonedaLocal)
            ldblComisionRecepcionista = ldblComisionRecepcionista + SacarFormatoValor(Item.SubItems(17), gstrMonedaLocal)
            Tabla.MoveNext
            Me.lsvComisiones.Refresh
            ProcesoRegistros gcAvanceProceso, 50
        Wend
    End If
End If
Conexion.CloseHost Tabla

With Me.stbTotales
    .Panels("Registros").Text = Me.lsvComisiones.ListItems.Count
    .Panels("Ots").Text = ldblOTs
    .Panels("Servicios").Text = FormatoValor(ldblTotalServicios, gstrMonedaLocal, gintDecimalesMoneda)
    .Panels("Mecanico").Text = FormatoValor(ldblComisionMecanico, gstrMonedaLocal, gintDecimalesMoneda)
    .Panels("Recepcionista").Text = FormatoValor(ldblComisionRecepcionista, gstrMonedaLocal, gintDecimalesMoneda)
End With

ProcesoRegistros gcFinProceso
End Sub

Private Sub dbcboMecanico_Click(Area As Integer)
If Area = 2 Then
    Me.lsvComisiones.ListItems.Clear
    LimpiaTotales
End If
End Sub

Private Sub dbCboServicio_Click(Area As Integer)
If Area = 2 Then
    Me.lsvComisiones.ListItems.Clear
    LimpiaTotales
End If
End Sub

Private Sub dtpDesde_Change()
Me.lsvComisiones.ListItems.Clear
LimpiaTotales
End Sub

Private Sub dtpHasta_Change()
Me.lsvComisiones.ListItems.Clear
LimpiaTotales
End Sub

Private Sub Form_Load()
LLena_Mecanico
LLena_TipoServicio
LLena_Sucursales
LimpiaTotales
Me.dtpDesde.Value = "01/" & Format$(Date, "mm/yyyy")
Me.dtpHasta.Value = Date
Me.dbcboSucursal.BoundText = gstrIdSucursal
End Sub

Private Sub LimpiaTotales()
With Me.stbTotales
    .Panels("Registros").Text = "0"
    .Panels("Ots").Text = "0"
    .Panels("Servicios").Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    .Panels("Mecanico").Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
    .Panels("Recepcionista").Text = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
End With
End Sub

Public Sub LLena_Sucursales()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Sucursal, Descripcion FROM Glbl_Sucursal WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Vigencia = 'S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datSucursal.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Mecanico()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Mecanico, Nombre FROM Tllr_Mecanicos WHERE Vigencia='S' ORDER BY Nombre"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datMecanico.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_TipoServicio()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Concepto_Servicio, Descripcion FROM Srvt_Concepto_Servicio WHERE Vigencia='S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datConceptoServicio.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Servicio()
Dim Tabla As New ADODB.Recordset
Dim sql As String

Me.dbCboServicio.Text = ""
sql = ""
sql = "SELECT Id_Servicio, Descripcion FROM Srvt_Servicios WHERE Id_Concepto_Servicio='" & Me.dbCboConceptoServicio.BoundText & "' AND Vigencia='S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datServicio.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Private Sub lsvComisiones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
OrdenaLista Me.lsvComisiones, ColumnHeader
End Sub

Private Sub OrdenaLista(ByRef Lista As ListView, ByVal Cabecera As MSComctlLib.ColumnHeader)
If Lista.SortKey = Cabecera.Index - 1 Then
'    If Me.optOrden(0).Value = True Then
        Lista.SortOrder = lvwAscending
'    Else
'        Lista.SortOrder = lvwDescending
'    End If
Else
    Lista.SortOrder = lvwAscending
    Lista.Sorted = False
    If Cabecera.Index - 1 = 2 Then
        Lista.SortKey = 18
    Else
        If Cabecera.Index - 1 = 3 Then
            Lista.SortKey = 19
        Else
            If Cabecera.Index - 1 = 8 Then
                Lista.SortKey = 20
            Else
                If Cabecera.Index - 1 = 9 Then
                    Lista.SortKey = 21
                Else
                    If Cabecera.Index - 1 = 10 Then
                        Lista.SortKey = 22
                    Else
                        If Cabecera.Index - 1 = 11 Then
                            Lista.SortKey = 23
                        Else
                            If Cabecera.Index - 1 = 12 Then
                                Lista.SortKey = 24
                            Else
                                If Cabecera.Index - 1 = 13 Then
                                    Lista.SortKey = 25
                                Else
                                    If Cabecera.Index - 1 = 10 Then
                                        Lista.SortKey = 22
                                    Else
                                        If Cabecera.Index - 1 = 10 Then
                                            Lista.SortKey = 22
                                        Else
                                            If Cabecera.Index - 1 = 10 Then
                                                Lista.SortKey = 22
                                            Else
                                                If Cabecera.Index - 1 = 10 Then
                                                    Lista.SortKey = 22
                                                Else
                                                    Lista.SortKey = Cabecera.Index - 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
'    If Me.optOrden(0).Value = True Then
        Lista.SortOrder = lvwAscending
'    Else
'        Lista.SortOrder = lvwDescending
'    End If
    Lista.Sorted = True
End If
End Sub

Private Sub opcAgrupaOT_Click()
If Me.opcAgrupaOT.Value = True Then
    Me.dbCboConceptoServicio.Text = ""
    Me.dbCboServicio.Text = ""
    Me.dbCboConceptoServicio.Enabled = False
    Me.dbCboServicio.Enabled = False
    Me.tlbBotones(1).Buttons(1).Enabled = False
    Me.tlbBotones(2).Buttons(1).Enabled = False
Else
    Me.dbCboConceptoServicio.Enabled = True
    Me.dbCboServicio.Enabled = True
    Me.tlbBotones(1).Buttons(1).Enabled = True
    Me.tlbBotones(2).Buttons(1).Enabled = True
End If
End Sub

Private Sub opcAgrupaServicio_Click()
If Me.opcAgrupaOT.Value = True Then
    Me.dbCboConceptoServicio.Text = ""
    Me.dbCboServicio.Text = ""
    Me.dbCboConceptoServicio.Enabled = False
    Me.dbCboServicio.Enabled = False
    Me.tlbBotones(1).Buttons(1).Enabled = False
    Me.tlbBotones(2).Buttons(1).Enabled = False
Else
    Me.dbCboConceptoServicio.Enabled = True
    Me.dbCboServicio.Enabled = True
    Me.tlbBotones(1).Buttons(1).Enabled = True
    Me.tlbBotones(2).Buttons(1).Enabled = True
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Screen.MousePointer = vbHourglass
Select Case Button.Key
    Case "Buscar"
        Buscar
    Case "Imprimir"
        ImprimirConsulta
    Case "Cerrar"
        Unload Me
End Select
Screen.MousePointer = vbDefault

End Sub

Private Sub tlbBotones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

Screen.MousePointer = vbHourglass
Select Case Button.Key
    Case "Cancelar"
        Select Case Index
            Case 0
                Me.dbcboMecanico.Text = ""
            Case 1
                Me.dbCboConceptoServicio.Text = ""
            Case 2
                Me.dbCboServicio.Text = ""
            Case 3
                Me.dbcboSucursal.Text = ""
        End Select
End Select
Screen.MousePointer = vbDefault
Me.lsvComisiones.ListItems.Clear
LimpiaTotales
End Sub

Private Sub tlbOtrosCriterios_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Value = tbrPressed Then
    Me.OtrosCriterios.Visible = True
Else
    Me.OtrosCriterios.Visible = False
End If
End Sub

Private Sub txtOtDesde_Change()
Me.lsvComisiones.ListItems.Clear
LimpiaTotales
End Sub

Private Sub txtOtDesde_GotFocus()
Me.txtOtDesde.SelStart = 0
Me.txtOtDesde.SelLength = Len(Me.txtOtDesde.Text)
End Sub

Private Sub txtOtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 75 And KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End If
If KeyAscii = 13 Then
    Me.txtOtHasta.Text = Me.txtOtDesde.Text
    Me.txtOtHasta.SetFocus
End If
End Sub

Private Sub txtOtHasta_Change()
Me.lsvComisiones.ListItems.Clear
LimpiaTotales
End Sub

Private Sub txtOtHasta_GotFocus()
Me.txtOtHasta.SelStart = 0
Me.txtOtHasta.SelLength = Len(Me.txtOtHasta.Text)
End Sub

Private Sub txtOtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 75 And KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End If
End Sub
