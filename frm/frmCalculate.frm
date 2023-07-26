VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalculate 
   Caption         =   "Calculate"
   ClientHeight    =   2664
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   4740
   OleObjectBlob   =   "frmCalculate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCompound_Click()
frmCompound.Show 'Open form for calculting compounds
End Sub

Private Sub btnMolarity_Click()
frmMolarity.Show 'Open form for calculting molarity
End Sub

Private Sub btnPercent_Click()
frmPercent.Show 'Open form for calculting percent composition
End Sub

Private Sub btnProNeu_Click()
frmProNeu.Show 'Open form for calculting protons and neutrons
End Sub
