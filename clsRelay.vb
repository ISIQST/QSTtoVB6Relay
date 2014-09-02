Imports System.Runtime.InteropServices

Public Class clsRelay
    'Implements Quasi97.iAddin.Initialize2
    Private qST As Quasi97.Application
    Private EvHandlerObj As Object
    Private WithEvents desig As Quasi97.Designators
    Private WithEvents SysEvents As Quasi97.clsSystem

    Public Sub Initialize2(ByRef q As Object) 'Implements System.IDisposable.Initialize2
        qST = q
        EvHandlerObj = CreateObject("projVB6.Application")
        EvHandlerObj.initialize2(q)
        desig = qST.DesignatorsParameters
        SysEvents = qST.SystemParameters
    End Sub

    Public Sub Dispose() 'Implements System.IDisposable.Dispose
        qST = Nothing
        EvHandlerObj = Nothing
        desig = Nothing
        SysEvents = Nothing
    End Sub

    Private Sub desig_DesignatorsChanged() Handles desig.DesignatorsChanged
        CallByName(EvHandlerObj, "DesigEvents_DesignatorsChanged", CallType.Method)
    End Sub

    Private Sub SysEvents_PreampChipChanged(ByRef Preamp As String) Handles SysEvents.PreampChipChanged
        CallByName(EvHandlerObj, "SysEvents_PreampChipChanged", CallType.Method, Preamp)
    End Sub

#Region "Other Required Properties"
    Public ReadOnly Property CustomStressSupport As Boolean 'Implements Quasi97.iAddin.CustomStressSupport
        Get
            Return False
        End Get
    End Property

    Public Property ModuleID As String = "" ' Implements Quasi97.iAddin.ModuleID

    Public ReadOnly Property QuasiAddIn As String ' Implements Quasi97.iAddin.QuasiAddIn
        Get
            Return True
        End Get
    End Property

    Public Sub RunCustomStress(sM As Quasi97.clsStress.StressMode, UniqueEventID As Short, ByRef EventParams As String) 'Implements Quasi97.iAddin.RunCustomStress

    End Sub

    Public Sub ValidateStressParam(eID As Short, ByRef ParamValue As String, ByRef RetValue As Boolean) 'Implements Quasi97.iAddin.ValidateStressParam

    End Sub
#End Region

End Class
