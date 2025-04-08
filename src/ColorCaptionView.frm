VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorCaptionView 
   ClientHeight    =   9345.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4875
   OleObjectBlob   =   "ColorCaptionView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorCaptionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'===============================================================================
' # State

Private Const MIN_SIZE As Double = 0

Public IsOk As Boolean
Public IsCancel As Boolean

Public BulkaHandler As TextBoxHandler
Public RulkaHandler As TextBoxHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = APP_DISPLAYNAME & " (v" & APP_VERSION & ")"
    Logo.ControlTipText = APP_URL
    btnOk.Default = True
    
    'Set BulkaHandler = _
        TextBoxHandler.New_(Bulka, TextBoxTypeDouble, MIN_SIZE)
    'Set RulkaHandler = _
        TextBoxHandler.New_(Rulka, TextBoxTypeDouble, MIN_SIZE)
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
'
End Sub

Private Sub btnOk_Click()
    FormОК
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

Private Sub Logo_Click()
    VBA.CreateObject("WScript.Shell").Run APP_URL '"https://vk.com/elvin_macro"
End Sub

'===============================================================================
' # Logic

Private Sub FormОК()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers


'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Сancel = True
        FormCancel
    End If
End Sub
