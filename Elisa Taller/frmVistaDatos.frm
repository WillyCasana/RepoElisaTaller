VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVistaDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista de Datos"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVistaDatos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   90
      TabIndex        =   2
      Top             =   3030
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Campos"
      Height          =   2970
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   4170
      Begin MSComctlLib.ListView lvwVistas 
         Height          =   2475
         Left            =   120
         TabIndex        =   1
         Top             =   315
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   4366
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Columna"
            Object.Width           =   6174
         EndProperty
      End
   End
End
Attribute VB_Name = "frmVistaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub lvwVistas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
       
    If Item.Checked Then
        gObjListView.ColumnHeaders(Item.Index).Width = "2500"
        'frmFullMailing.lvwListaFiltro.ColumnHeaders(Item.Index).Width = "2500"
    Else
        gObjListView.ColumnHeaders(Item.Index).Width = "0"
        'frmFullMailing.lvwListaFiltro.ColumnHeaders(Item.Index).Width = "0"
    End If

End Sub
