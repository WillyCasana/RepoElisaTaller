VERSION 5.00
Begin VB.Form frmLiquidacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ControlBox      =   0   'False
   Icon            =   "frmLiquidacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3120
      TabIndex        =   38
      Top             =   5790
      Width           =   1215
   End
   Begin VB.CommandButton cmdSeguroTaller 
      Height          =   240
      Left            =   2325
      TabIndex        =   36
      Top             =   4365
      Width           =   285
   End
   Begin VB.CommandButton cmdMateriales 
      Height          =   240
      Left            =   2325
      TabIndex        =   34
      Top             =   3750
      Width           =   285
   End
   Begin VB.CommandButton cmdInsumos 
      Height          =   240
      Left            =   2325
      TabIndex        =   33
      Top             =   4050
      Width           =   285
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   5790
      Width           =   1215
   End
   Begin VB.Label lblTotalTer 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   18
      Top             =   3405
      Width           =   2970
   End
   Begin VB.Label lblSeguroTaller 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   2670
      TabIndex        =   37
      Top             =   4365
      Width           =   2970
   End
   Begin VB.Label Label1 
      Caption         =   "Seguro Taller :"
      Height          =   300
      Index           =   16
      Left            =   60
      TabIndex        =   35
      Top             =   4320
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Orden de Trabajo :"
      Height          =   300
      Index           =   15
      Left            =   60
      TabIndex        =   32
      Top             =   5430
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "I.V.A. :"
      Height          =   300
      Index           =   14
      Left            =   60
      TabIndex        =   31
      Top             =   5115
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "SubTotal Orden de Trabajo:"
      Height          =   300
      Index           =   13
      Left            =   60
      TabIndex        =   30
      Top             =   4800
      Width           =   2580
   End
   Begin VB.Label lblPatente 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   29
      Top             =   1605
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Patente :"
      Height          =   300
      Index           =   12
      Left            =   60
      TabIndex        =   28
      Top             =   1605
      Width           =   2580
   End
   Begin VB.Label lblTotalOT 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2670
      TabIndex        =   27
      Top             =   5415
      Width           =   2970
   End
   Begin VB.Label lblIva 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2670
      TabIndex        =   26
      Top             =   5100
      Width           =   2970
   End
   Begin VB.Label lblSeccion 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   25
      Top             =   330
      Width           =   3000
   End
   Begin VB.Label lblCliente 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   24
      Top             =   645
      Width           =   3000
   End
   Begin VB.Label lblMarca 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   23
      Top             =   960
      Width           =   3000
   End
   Begin VB.Label lblModelo 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   22
      Top             =   1275
      Width           =   3000
   End
   Begin VB.Label lblTotalMec 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   21
      Top             =   2190
      Width           =   2970
   End
   Begin VB.Label lblTotalCar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   20
      Top             =   3120
      Width           =   2970
   End
   Begin VB.Label lblTotalOtr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   19
      Top             =   2490
      Width           =   2970
   End
   Begin VB.Label lblTotalRep 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   17
      Top             =   2805
      Width           =   2970
   End
   Begin VB.Label lblTotalMat 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   2670
      TabIndex        =   16
      Top             =   3720
      Width           =   2970
   End
   Begin VB.Label lblTotalIns 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   15
      Top             =   4035
      Width           =   2970
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2670
      TabIndex        =   14
      Top             =   4785
      Width           =   2970
   End
   Begin VB.Label lblIdOT 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2670
      TabIndex        =   13
      Top             =   15
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Total Insumos :"
      Height          =   300
      Index           =   11
      Left            =   60
      TabIndex        =   12
      Top             =   4020
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Materiales :"
      Height          =   300
      Index           =   4
      Left            =   60
      TabIndex        =   11
      Top             =   3720
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Sección :"
      Height          =   300
      Index           =   10
      Left            =   60
      TabIndex        =   10
      Top             =   345
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Marca :"
      Height          =   300
      Index           =   9
      Left            =   60
      TabIndex        =   9
      Top             =   960
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo :"
      Height          =   300
      Index           =   8
      Left            =   60
      TabIndex        =   8
      Top             =   1275
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Servicios Mecánica :"
      Height          =   300
      Index           =   7
      Left            =   60
      TabIndex        =   7
      Top             =   2205
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Servicios Carrocería :"
      Height          =   300
      Index           =   6
      Left            =   60
      TabIndex        =   6
      Top             =   3120
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Otros Servicios :"
      Height          =   300
      Index           =   5
      Left            =   60
      TabIndex        =   5
      Top             =   2505
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Servicios de Terceros :"
      Height          =   300
      Index           =   3
      Left            =   60
      TabIndex        =   4
      Top             =   3420
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Total Repuestos :"
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   3
      Top             =   2805
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   300
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   2580
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Trabajo Nº:"
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   15
      Width           =   2580
   End
End
Attribute VB_Name = "frmLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrValorInsumo As String
Dim mstrValorMateriales As String
Dim mstrValorSeguroTaller As String
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
    Neto = Neto + Val(SacarFormatoValor(lblSeguroTaller, ""))
    .lblSubTotal = FormatoValor(Neto, "", gintDecimalesMoneda)
    
    .lblIva = FormatoValor(CStr(Neto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto)), "", gintDecimalesMoneda)
    .lblTotalOT = FormatoValor(CStr(Neto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto)), "", gintDecimalesMoneda)
    gcurTotalNeto = Neto
    gcurTotalIVA = SacarFormatoValor(.lblIva, "")             'Neto * 0.18
    gcurTotalNetoMasIVA = SacarFormatoValor(.lblTotalOT, "")  'Neto * 1.18
End With


End Sub



Private Sub cmdAceptar_Click()
ReCalculo
Unload Me
End Sub

Private Sub cmdCancelar_Click()
gblnCierraLiq = False
Unload Me
End Sub

Private Sub cmdInsumos_Click()
mstrValorInsumo = InputBox("", "Insumos", CStr(gcurInsumo))
If mstrValorInsumo <> "" Then
        gcurInsumo = Val(mstrValorInsumo)
        lblTotalIns = FormatoValor(mstrValorInsumo, "", gintDecimalesMoneda)
        ReCalculo
End If
End Sub

Private Sub cmdMateriales_Click()
mstrValorMateriales = InputBox("", "Materiales", CStr(gcurMateriales))
If mstrValorMateriales <> "" Then
    gcurMateriales = Val(mstrValorMateriales)
    lblTotalMat = FormatoValor(mstrValorMateriales, "", gintDecimalesMoneda)
    ReCalculo
End If
End Sub

Private Sub cmdSeguroTaller_Click()
mstrValorSeguroTaller = InputBox("", "Seguro Taller", CStr(gcurSeguroTaller))
If mstrValorSeguroTaller <> "" Then
    gcurSeguroTaller = Val(mstrValorSeguroTaller)
    lblSeguroTaller = FormatoValor(mstrValorSeguroTaller, "", gintDecimalesMoneda)
    ReCalculo
End If

End Sub

Private Sub Form_Load()
''kjcv 08.02.16
'Inhabilitado para que Recepcionista No Modifique Insumos
'If gstrIdPerfil = "Tllr_0010" Then
'    Me.cmdInsumos.Enabled = False
'Else
    Me.cmdInsumos.Enabled = True
'End If
With frmRecepcion
    Me.Label1(12).Caption = gstrNombrePatente
    Me.Label1(14).Caption = gstrNombreIva
    gblnCierraLiq = True
    gcurMateriales = CDbl(.stbTotalMateriales.Panels(2).Text)
    lblIdOT = .lblNroRecepcion
    lblSeccion = IIf(gstrSeccion = "C", "Carrocería", "Mecánica")
    lblCliente = .lblCliente
    lblMarca = .lblMarca
    lblModelo = .lblModelo
    lblPatente = .txtPatente
    
    lblTotalMec = .stbTotalMec.Panels(2).Text
    lblTotalCar = .stbTotalCarroceria.Panels(2).Text
    lblTotalOtr = .stbTotalOtros.Panels(2).Text
    lblTotalTer = .stbTotalTerceros.Panels(2).Text
    lblTotalRep = FormatoValor(CDbl(.stbTotalRepuestos.Panels(2).Text) + CDbl(.StbLubricantes.Panels(2).Text) + IIf(CDbl(.stbInsumos.Panels(2).Text) <> 0, (CDbl(.stbInsumos.Panels(2).Text) - gcurInsumo), CDbl(.stbInsumos.Panels(2).Text)), "", gintDecimalesMoneda)
    'pregunto si tiene insumos en porcentaje o pesos
    If gcurMaterialesMO <> 0 Then
        gcurInsumo = Round(((CDbl(.stbTotalMec.Panels(2)) + CDbl(.stbTotalOtros.Panels(2))) * gcurMaterialesMO) / 100, gintDecimalesMoneda)
    End If
    lblTotalMat = FormatoValor(gcurMateriales, "", gintDecimalesMoneda)
    lblTotalIns = FormatoValor(gcurInsumo, "", gintDecimalesMoneda)
    lblSeguroTaller = FormatoValor(gcurSeguroTaller, "", gintDecimalesMoneda)
    ReCalculo
    'lblsubtotal = .stbTotalOT.Panels(2).Text
    'lblIva = FormatoValor(CCur(Val(SacarFormatoValor(.stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto), "", 0)
    'lblTotalOT = FormatoValor(CCur(Val(SacarFormatoValor(.stbTotalOT.Panels(2).Text, ""))) * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto), "", 0)
    Screen.MousePointer = vbDefault
End With

End Sub
