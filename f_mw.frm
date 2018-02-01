VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_mw 
   Caption         =   "Magic Word"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6540
   OleObjectBlob   =   "f_mw.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "f_mw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub img_bg_Click()

    Unload Me

End Sub

Private Sub UserForm_Activate()

    HideTitleBar Me
    With Me
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
End Sub
