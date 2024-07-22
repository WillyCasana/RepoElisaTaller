Attribute VB_Name = "basEnums"
Option Explicit

Public Enum gcIva
    gcIvaUnoPto
    gcIvaCeroPto
End Enum

Public Enum gAccionEstadoOT
    gOTAnular
    gOTActivar
    gOTLiquidar
End Enum


Public Enum mcFicha
    mcFichaMecanica
    mcFichaCarroceria
    mcFichaTerceros
    mcFichaRepuestos
    mcFichaOtros
End Enum
Public Enum tpoFields
    tpoCodigo
    tpoNombre
End Enum
Public Enum apColor
     apinterno
     apExterno
End Enum
Public Enum mAccionItem
    mAddItem
    mDelItem
    mRefItem
End Enum
Public Enum gcProceso
    gcInicioProceso = 0
    gcFinProceso = 1
    gcAvanceProceso = 2
End Enum
Public Enum gcParametro
    gcPresupuesto = 1
    gcOrdenTrabajo = 2
End Enum
Public Enum gopOpcionItem
    gcSelectAll
    gcUnSelectAll
End Enum
Public Enum gInforme
    gRecepcion
    gPresupuesto
    gOT
End Enum

Public Enum SumSec
    ssMec
    ssOtr
    ssCar
    ssTer
    ssRep
End Enum

Enum gcApertura
    gcdynamic = 0
    gcstatic = 1
    gckeyset = 2
    gcForOnly = 3
End Enum

Public Enum gcFamilia
    gcRepuesto = 1
    gcMateriales = 2
    gcLubricantes = 3
    gcTodos = 4
End Enum

