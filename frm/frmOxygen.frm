VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOxygen 
   Caption         =   "Oxygen"
   ClientHeight    =   3972
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3888
   OleObjectBlob   =   "frmOxygen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOxygen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSoundO_Click()
Application.Speech.Speak "Oxygen" 'play a sound file of a man saying "Oxygen"
End Sub
