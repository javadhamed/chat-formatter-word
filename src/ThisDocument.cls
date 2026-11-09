VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_Base = "0{00020906-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

'=======================================================================
'  Chat Formatter - document & application level event handlers
'
'  These run automatically because the template lives in Word's STARTUP
'  folder and is loaded as a global add-in. No button press needed.
'=======================================================================

Private WithEvents mApp As Word.Application


Public Sub HookApplication()
    If mApp Is Nothing Then
        Set mApp = Application
    End If
End Sub


Private Sub Document_Open()
    On Error Resume Next
    ChatFormatter.AutoExec
    On Error GoTo 0
End Sub


Private Sub Document_New()
    On Error Resume Next
    ChatFormatter.AutoExec
    On Error GoTo 0
End Sub


' Right-click anywhere in Word: if the current selection looks like
' text copied from an AI chat, format it in place.
Private Sub mApp_WindowBeforeRightClick(ByVal Sel As Selection, Cancel As Boolean)
    On Error Resume Next
    If Not Sel Is Nothing Then
        ChatFormatter.CF_RightClickHandler Sel.Range
    End If
    On Error GoTo 0
End Sub
