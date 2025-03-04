Attribute VB_Name = "LogicalConnections"
Option Explicit

Public Function ObjectIsOkSurface(oInputObj As Object) As Boolean

    Const METHOD = "ObjectIsOkSurface"
    On Error GoTo errorHandler

    Dim oGeom As Object
    Dim oPort As IJPort
    Dim bOk As Boolean

    bOk = False

    If Not oInputObj Is Nothing Then

        If TypeOf oInputObj Is IJDynamicSurfaceFind Then
            Set oGeom = oInputObj
        ElseIf TypeOf oInputObj Is IJPort Then
            Set oPort = oInputObj
            Set oGeom = oPort.Geometry
        Else
            Set oGeom = oInputObj
        End If
        
        If oGeom Is Nothing Then
            bOk = False
        ElseIf TypeOf oGeom Is IJSurface Then
            bOk = True
        ElseIf TypeOf oGeom Is IJPlane Then
            bOk = True
        ElseIf TypeOf oGeom Is IJDynamicSurfaceFind Then
            bOk = True
        End If
    End If

    ' any surface port whose connectable is a MemberPart is not a valid Surface for FC's
    If Not oPort Is Nothing Then
        Set oGeom = oPort.Connectable
        If Not oGeom Is Nothing Then
            If TypeOf oGeom Is ISPSMemberPartPrismatic Then
                bOk = False
            End If
        End If
    End If

    ObjectIsOkSurface = bOk
    Exit Function

errorHandler:
    HandleError "LogicalConnections", METHOD
    Err.Clear
    ObjectIsOkSurface = False
    Exit Function
End Function


