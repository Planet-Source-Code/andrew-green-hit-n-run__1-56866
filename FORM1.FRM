VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hit'n Run!"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then GHelp = Not (GHelp)
If KeyCode = 113 Then GOps = Not (GOps)
If GOps Then
    Select Case KeyCode
        Case 49:
            If AIT = "None" Then
                AIT = "Scared"
            ElseIf AIT = "Scared" Then
                AIT = "Suicidal"
            Else
                AIT = "None"
            End If
        Case 50:
            AIRnd = Not (AIRnd)
    End Select
End If
End Sub

Private Sub Form_Load()
InitiateDX (Me.hWnd)
BeginGLoop
End Sub

