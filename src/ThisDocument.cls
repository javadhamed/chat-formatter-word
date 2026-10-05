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
'  Chat Formatter - document level event handlers
'
'  These run automatically because the template lives in Word's STARTUP
'  folder and is loaded as a global add-in. No button press needed.
'=======================================================================

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


' Right-click anywhere in a document: if the current selection looks like
' text copied from an AI chat, format it in place.
Private Sub WindowBeforeRightClick(ByVal Target As Range, ByVal Cancel As Boolean)
    On Error Resume Next
    ChatFormatter.CF_RightClickHandler Target
    On Error GoTo 0
End Sub