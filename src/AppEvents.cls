VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application

Private Sub App_WindowBeforeRightClick(ByVal Target As Range, ByRef Cancel As Boolean)
    On Error Resume Next
    ChatFormatter.CF_RightClickHandler Target
    On Error GoTo 0
End Sub