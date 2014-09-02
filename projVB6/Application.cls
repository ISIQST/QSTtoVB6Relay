VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private qST As Object 'Quasi97.Quasi97_Application

Public Property Get QuasiAddIn() As Boolean
   QuasiAddIn = True
End Property

Public Sub Initialize2(ByRef q As Object)
   Set qST = q
End Sub

Public Sub DesigEvents_DesignatorsChanged()
   Call MsgBox("Designator Events", vbSystemModal Or vbOKOnly)
End Sub

Public Sub SysEvents_PreampChipChanged(ByRef Preamp As String)
   Call MsgBox("Preamp Changed", vbSystemModal Or vbOKOnly)
End Sub

Private Sub Class_Terminate()
   Set qST = Nothing
End Sub
