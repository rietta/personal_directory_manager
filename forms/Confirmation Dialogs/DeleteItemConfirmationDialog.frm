VERSION 5.00
Begin VB.Form DeleteItemConfirmationDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Item"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "DeleteItemConfirmationDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Don't Delete the Item"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdWasteBin 
      Caption         =   "&Move Item to Waste Bin"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Delete the Item, but move it to the waste bin so it can be recovered later."
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Item"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      ToolTipText     =   "Delete the item"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblQuestion 
      Caption         =   "Are you sure you want to delete the *item* from *binder*?"
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "DeleteItemConfirmationDialog.frx":0442
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "DeleteItemConfirmationDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    nDialogButton = CANCELBUTTON
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    nDialogButton = DELETEBUTTON
    Unload Me
End Sub

Private Sub cmdWasteBin_Click()
    nDialogButton = WASTEBUTTON
    Unload Me
End Sub

Private Sub Form_Load()
    nDialogButton = CANCELBUTTON    ' Cancel operation is default
End Sub
