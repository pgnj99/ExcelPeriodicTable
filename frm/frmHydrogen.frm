VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHydrogen 
   Caption         =   "Hydrogen"
   ClientHeight    =   3972
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   3888
   OleObjectBlob   =   "frmHydrogen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHydrogen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSoundH_Click()
Application.Speech.Speak "Hydrogen" 'play a sound file of a man saying "Hydrogen"
End Sub
