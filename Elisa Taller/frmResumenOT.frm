VERSION 5.00
Begin VB.Form frmResumenOT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen OT"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmResumenOT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTotalTer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   17
      Top             =   3810
      Width           =   2970
   End
   Begin VB.Label Label1 
      Caption         =   "Estado :"
      Height          =   300
      Index           =   16
      Left            =   60
      TabIndex        =   33
      Top             =   345
      Width           =   2580
   End
   Begin VB.Label lblEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2655
      TabIndex        =   32
      Top             =   330
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Total Orden de Trabajo :"
      Height          =   300
      Index           =   15
      Left            =   60
      TabIndex        =   31
      Top             =   5760
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "I.V.A. :"
      Height          =   300
      Index           =   14
      Left            =   60
      TabIndex        =   30
      Top             =   5445
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "SubTotal Orden de Trabajo:"
      Height          =   300
      Index           =   13
      Left            =   60
      TabIndex        =   29
      Top             =   5130
      Width           =   2580
   End
   Begin VB.Label lblPatente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   28
      Top             =   2100
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Patente :"
      Height          =   300
      Index           =   12
      Left            =   60
      TabIndex        =   27
      Top             =   2100
      Width           =   2580
   End
   Begin VB.Label lblTotalOT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   26
      Top             =   5745
      Width           =   2970
   End
   Begin VB.Label lblIva 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   25
      Top             =   5430
      Width           =   2970
   End
   Begin VB.Label lblSeccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   24
      Top             =   825
      Width           =   3000
   End
   Begin VB.Label lblCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   23
      Top             =   1140
      Width           =   3000
   End
   Begin VB.Label lblMarca 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   22
      Top             =   1455
      Width           =   3000
   End
   Begin VB.Label lblModelo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   21
      Top             =   1770
      Width           =   3000
   End
   Begin VB.Label lblTotalMec 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   20
      Top             =   2595
      Width           =   2970
   End
   Begin VB.Label lblTotalCar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   19
      Top             =   3525
      Width           =   2970
   End
   Begin VB.Label lblTotalOtr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   18
      Top             =   2895
      Width           =   2970
   End
   Begin VB.Label lblTotalRep 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   16
      Top             =   3210
      Width           =   2970
   End
   Begin VB.Label lblTotalMat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   15
      Top             =   4125
      Width           =   2970
   End
   Begin VB.Label lblTotalIns 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   14
      Top             =   4440
      Width           =   2970
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   13
      Top             =   5115
      Width           =   2970
   End
   Begin VB.Label lblIdOT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2670
      TabIndex        =   12
      Top             =   15
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Total Insumos :"
      Height          =   300
      Index           =   11
      Left            =   60
      TabIndex        =   11
      Top             =   4425
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Materiales :"
      Height          =   300
      Index           =   4
      Left            =   60
      TabIndex        =   10
      Top             =   4125
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Sección :"
      Height          =   300
      Index           =   10
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono :"
      Height          =   300
      Index           =   9
      Left            =   60
      TabIndex        =   8
      Top             =   1455
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo :"
      Height          =   300
      Index           =   8
      Left            =   60
      TabIndex        =   7
      Top             =   1770
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Servicios Mecánica :"
      Height          =   300
      Index           =   7
      Left            =   60
      TabIndex        =   6
      Top             =   2610
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Servicios Carrocería :"
      Height          =   300
      Index           =   6
      Left            =   60
      TabIndex        =   5
      Top             =   3525
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Otros Servicios :"
      Height          =   300
      Index           =   5
      Left            =   60
      TabIndex        =   4
      Top             =   2910
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Servicios de Terceros :"
      Height          =   300
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   3825
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Repuestos :"
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   3210
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   300
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   1155
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Trabajo Nº:"
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   2580
   End
End
Attribute VB_Name = "frmResumenOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrValorInsumo As String
Dim mstrValorMateriales As String
Sub ReCalculo()
Dim Neto As Currency
With Me
    Neto = Val(SacarFormatoValor(lblTotalMec, ""))
    Neto = Neto + Val(SacarFormatoValor(lblTotalCar, ""))
    Neto = Neto + Val(SacarFormatoValor(lblTotalOtr, ""))
    Neto = Neto + Val(SacarFormatoValor(lblTotalTer, ""))
    Neto = Neto + Val(SacarFormatoValor(lblTotalRep, ""))
    Neto = Neto + Val(SacarFormatoValor(lblTotalMat, ""))
    Neto = Neto + Val(SacarFormatoValor(lblTotalIns, ""))
    
    .lblsubtotal = FormatoValor(Neto, "", gintDecimalesMoneda)
    .lblIva = FormatoValor(CStr(Neto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto)), "", gintDecimalesMoneda)
    .lblTotalOT = FormatoValor(CStr(Neto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto)), "", gintDecimalesMoneda)
    gcurTotalNeto = Neto
    gcurTotalIVA = SacarFormatoValor(.lblIva, "") 'Neto * 0.18
    gcurTotalNetoMasIVA = SacarFormatoValor(.lblTotalOT, "") 'Neto * 1.18
End With

End Sub

Private Sub Form_Load()
Me.Label1(12).Caption = gstrNombrePatente
Me.Label1(14).Caption = gstrNombreIva
End Sub

