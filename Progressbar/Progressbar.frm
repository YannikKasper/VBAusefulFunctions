VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progressbar 
   Caption         =   "Progressbar"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "Progressbar.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    Me.barProgress.Width = 0
End Sub

Public Sub progress(ByVal percentProgress As Integer)
    If percentProgress >= 0 And percentProgress <= 100 Then
        Me.barProgress.Width = percentProgress * 2
        Me.labelPercent.caption = percentProgress & "%"
    End If
End Sub

Public Sub addProgress(ByVal percentProgress As Integer)
    If (percentProgress + Me.barProgress.Width / 2) >= 0 And (percentProgress + Me.barProgress.Width / 2) <= 100 Then
        Me.barProgress.Width = Me.barProgress.Width + percentProgress * 2
        Me.labelPercent.caption = Me.barProgress.Width / 2 & "%"
    End If
End Sub
Public Sub changeLabel(ByVal caption As String)
    Me.labelAction.caption = caption
End Sub

