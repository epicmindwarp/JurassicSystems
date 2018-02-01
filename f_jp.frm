VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_jp 
   Caption         =   "Jurassic Park"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10665
   OleObjectBlob   =   "f_jp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_jp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub b_close_Click()
    
    Unload Me
    If Application.EnableEvents Then ThisWorkbook.Close False

End Sub

Private Sub img_header_Click()

End Sub

Private Sub tb_1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    Me.lbl_error1.Visible = True
    Me.tb_a2.Visible = True
    Me.tb_2.Visible = True
    Me.tb_2.SetFocus

End Sub

Private Sub tb_2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Me.lbl_error2.Visible = True
    Me.tb_a3.Visible = True
    Me.tb_3.Visible = True
    Me.tb_3.SetFocus
    
End Sub

Private Sub tb_3_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim ctl As Control

    If Me.Visible = True Then
    
        Me.lbl_error3.Visible = True
        
        Application.Wait Now + #12:00:01 AM#
        
        Do Until Me.tb_magic.Top = 0
        
            For Each ctl In Me.Controls
                
                Select Case TypeName(ctl)
                
                    Case "Label", "TextBox"
                    
                        ctl.Top = ctl.Top - 3
                        
                End Select
            
                Me.Repaint
            
            Next
        
        Loop
        
        f_mw.Show
        Unload Me
        If Application.EnableEvents Then ThisWorkbook.Close False
    
    End If
    
End Sub

Private Sub UserForm_Activate()

    HideTitleBar Me
    
    With Me
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    Me.tb_1.SetFocus

End Sub
