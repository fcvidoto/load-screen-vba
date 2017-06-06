VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Constructor da classe
Private Sub Class_Initialize()
    frmLoad.Show False
End Sub

Private Sub Class_Terminate()
    Unload frmLoad
End Sub

