Attribute VB_Name = "Locals"
Option Explicit

Public Function GetPOM(sTypeOfDatabase As String) As IJDPOM
    ' initialize result
    Dim pPOM As IJDPOM: Set pPOM = Nothing
    
    Dim oAccessMiddle As IJDAccessMiddle
    If True Then
        Dim oContext As IJContext
        Set oContext = GetJContext()
        Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    End If
    
    On Error Resume Next
    Dim pUnknownOfPOM As IUnknown
    Set pUnknownOfPOM = oAccessMiddle.GetResourceManagerFromType(sTypeOfDatabase)
    If Err.Number = 0 Then Set pPOM = pUnknownOfPOM
    On Error GoTo 0
    
    ' return result
    Set GetPOM = pPOM
End Function
Public Function DoesGCTypeExist(ByVal sNameOfGCType As String, _
                                ByVal pGCEntitiesFactory As IJGeometricConstructionEntitiesFactory) As Boolean
    ' initialize result
    DoesGCTypeExist = False
    
    ' check if sNameOfGCType entity can be created
    Dim oGC As Object
    On Error Resume Next
        If Not pGCEntitiesFactory Is Nothing Then
            Set oGC = pGCEntitiesFactory.CreateEntity(sNameOfGCType, Nothing)
        End If
    On Error GoTo 0
    
    ' return result
    DoesGCTypeExist = Not oGC Is Nothing
End Function
Public Function GetGlobalShowDetails() As Long
    Dim oToolshelper As New PORTHELPERLib.ToolsHelper
    GetGlobalShowDetails = oToolshelper.GetGlobalShowDetails()
End Function

