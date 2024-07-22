VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistorico 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "VEH"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2790
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OT"
      Height          =   495
      Left            =   165
      TabIndex        =   2
      Top             =   2775
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1860
      Left            =   210
      TabIndex        =   0
      Top             =   570
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   90
      TabIndex        =   1
      Top             =   75
      Width           =   7680
   End
End
Attribute VB_Name = "frmHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

