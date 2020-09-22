VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "You MUST have..."
   ClientHeight    =   1245
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6075
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "YOU MUST HAVE WINDOWS 98 FOR THIS TO WORK!  I AM NOT SURE IF IT WORKS ON OTHER PLATFORMS!!!!!"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit 'not needed

Private Sub OKButton_Click()
Form1.Show 'show Form1
Unload Me 'unload this form
End Sub
