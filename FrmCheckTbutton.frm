VERSION 5.00
Begin VB.Form FrmCheckTbutton 
   Caption         =   "Tbutton"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmCheckTbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mInDev                  As Boolean
Dim WithEvents TButton1     As cTButton
Attribute TButton1.VB_VarHelpID = -1

Private Sub Form_Load()
    Debug.Assert (InDev())
    Set TButton1 = New cTButton
    With TButton1
        If mInDev Then
            .IconFilename = App.Path & "\Tbutton.ico"
            .IconFilenameBG = App.Path & "\TbuttonBG.ico"
        Else
            .ResourceID = 1000
        End If
        
        .Edge = 100
        .hwnd = Me.hwnd
    End With
    
    
End Sub

Private Sub TButton1_Click()

    MsgBox "  Vote For TButton !!!!"
End Sub

Private Function InDev() As Boolean
    mInDev = True
    InDev = True
End Function
