VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInoVBEError 
   Caption         =   "UserForm1"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frmInoVBEError.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmInoVBEError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ErrButton As Integer

Private strCaption As String
Private strBtnOK As String
Private strBtnDebug As String

Private Sub cmdDebug_Click()
    ErrButton = 2
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ErrButton = 1
    Unload Me
End Sub

Public Function ShowForm(strLabel As String, Optional strLang As String = "de-DE") As Integer
    SetLanguage strLang
    
    Me.Caption = strCaption
    Me.cmdDebug.Caption = strBtnDebug
    Me.cmdOK.Caption = strBtnOK
    
    Me.lblErrorMsg.Caption = strLabel
    
    Me.Show
    
    ShowForm = ErrButton
End Function

Private Sub ExampleCall()
    Select Case frmInoVBEError.ShowForm(ErrorMessage, Language)
        Case 1
            
        Case 2
            Debug.Assert False
    End Select
End Sub

Private Sub SetLanguage(strLang As String)
    Select Case strLang
        Case "de-DE"
            strCaption = "Fehler"
            strBtnOK = "OK"
            strBtnDebug = "Debug"
        Case Else ' default English
            strCaption = "Error"
            strBtnOK = "OK"
            strBtnDebug = "Debug"
    End Select
End Sub


