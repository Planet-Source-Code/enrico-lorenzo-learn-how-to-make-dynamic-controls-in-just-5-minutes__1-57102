VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cmdDynamic As CommandButton
Attribute cmdDynamic.VB_VarHelpID = -1

Private Sub Form_Load()
Set cmdDynamic = Controls.Add("VB.CommandButton", "Command1")

cmdDynamic.Caption = "Click Me"
cmdDynamic.Visible = True
End Sub

Private Sub cmdDynamic_Click()
MsgBox "This is a dynamic control."
End Sub

 

