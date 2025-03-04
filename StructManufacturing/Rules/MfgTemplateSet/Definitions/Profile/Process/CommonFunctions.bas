Attribute VB_Name = "CommonFunctions"

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
Public Function GetPositionFrames(ByVal pFrameSysDsp As Object, ByVal pProfilePart As IJProfilePart, ByVal pProcessSettings As Object, ByVal pMfgTemplateSetDisp As Object) As IJElements
Const METHOD = "GetPositionFrames"
On Error GoTo ErrorHandler

    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    Dim oPlateSideRule   As IJDMfgPlateUpSideRule
    Dim oSettingsHelper  As MfgSettingsHelper
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Dim strProgId As String

    Set oSettingsHelper = pProcessSettings
    Set oProcessSettings = pProcessSettings
    
    'TemplateSide
    strErrorMsg = "TemplateType failed"
    strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
    
    Dim oProfileSideRule   As IJDMfgProfileUpsideRule
    Dim lUpside As Long
    Set oProfileSideRule = SP3DCreateObject(strProgId)
    lUpside = oProfileSideRule.GetUpside(pProfilePart)
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = pProfilePart
    
    Dim oSurfacePort As IJPort
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfacePort = oProfileWrapper.GetSurfacePort(lUpside)
    Set oSurfaceBody = oSurfacePort.Geometry

    Dim strTemplateType As String
    strTemplateType = oProcessSettings.TemplateType

    Dim oDirectionVec As DVector
    Set oDirectionVec = New DVector

    Dim oFramesColl As IJElements
    Set oFramesColl = New JObjectCollection
    Dim oTempFrameColl As IJElements
    Set oTempFrameColl = New JObjectCollection
    Dim strDirection As String
    strDirection = oProcessSettings.TemplateDirection
    
    'Get Landing Curve vector
    Dim oLandingCurveVec As IJDVector
    Dim oLandingCurve As IJWireBody
    Dim oStartPos  As IJDPosition, oEndPos As IJDPosition
    Set oLandingCurve = oProfileWrapper.GetLandingCurve
    oLandingCurve.GetEndPoints oStartPos, oEndPos
    Set oLandingCurveVec = oEndPos.Subtract(oStartPos)
    
    'get avgnormal of the surface
    Dim dRootX As Double, dRootY As Double, dRootZ As Double, dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim oSurfaceNormal As IJDVector
    oGeomHelper.GetPlatePartAvgPointAvgNormal oSurfaceBody, False, dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ
    Set oSurfaceNormal = New DVector
    oSurfaceNormal.Set dNormalX, dNormalY, dNormalZ
              
    Dim oXVector As IJDVector, oYVector As IJDVector, oZVector As IJDVector
    Set oXVector = New DVector
    Set oYVector = New DVector
    Set oZVector = New DVector
    oXVector.Set 1, 0, 0
    oYVector.Set 0, 1, 0
    oZVector.Set 0, 0, 1
              
    Dim dXDot As Double, dYDot As Double, dZDot As Double
    If strDirection = "PerpToAxis" Then 'X - Direction(Buttock)
        'Find standard axis closest to LC
        dXDot = Abs(oLandingCurveVec.Dot(oXVector))
        dYDot = Abs(oLandingCurveVec.Dot(oYVector))
        dZDot = Abs(oLandingCurveVec.Dot(oZVector))
        
        If dXDot > dYDot And dXDot > dZDot Then
            Set oDirectionVec = oXVector
        ElseIf dYDot > dXDot And dYDot > dZDot Then
            Set oDirectionVec = oYVector
        ElseIf dZDot > dXDot And dZDot > dYDot Then
            Set oDirectionVec = oZVector
        End If
    ElseIf strDirection = "AlongAxis" Then 'Y - Direction(Frame)
        'get cross product vecor(this is perp to axis vector)
        Dim oPerpToAxisVec As IJDVector
        Set oPerpToAxisVec = oSurfaceNormal.Cross(oLandingCurveVec)
        'get standard axis closest to this
        dXDot = Abs(oPerpToAxisVec.Dot(oXVector))
        dYDot = Abs(oPerpToAxisVec.Dot(oYVector))
        dZDot = Abs(oPerpToAxisVec.Dot(oZVector))
        
        If dXDot > dYDot And dXDot > dZDot Then
            Set oDirectionVec = oXVector
        ElseIf dYDot > dXDot And dYDot > dZDot Then
            Set oDirectionVec = oYVector
        ElseIf dZDot > dXDot And dZDot > dYDot Then
            Set oDirectionVec = oZVector
        End If
    End If
    
     oGeomHelper.GetReferencePlanesInRange oSurfaceBody, oDirectionVec, oTempFrameColl
     oFramesColl.AddElements oTempFrameColl


    ''''''''''''''''Sort Frames

    Dim oFrameSet As IJElements
    Set oFrameSet = New JObjectCollection

    Dim nIndex As Long
    Dim oFrame As IHFrame
    Dim oFrameAxis As IHFrameAxis
    Dim oFrameSystem As IHFrameSystem

    If Not pFrameSysDsp Is Nothing Then
        For nIndex = 1 To oFramesColl.Count
            Set oFrame = oFramesColl.Item(nIndex)
            Set oFrameAxis = oFrame.FrameAxis
            Set oFrameSystem = oFrameAxis.FrameSystem

            If oFrameSystem Is pFrameSysDsp Then
                oFrameSet.Add oFramesColl.Item(nIndex)
            End If
        Next nIndex
    End If

      ' Send the data back
    Set GetPositionFrames = oFrameSet

  ' *************************************************************
  ' * End Example
  ' *************************************************************


    Set oGeomHelper = Nothing
    Set oProcessSettings = Nothing
    Set oSettingsHelper = Nothing
    Set oPlateSideRule = Nothing
    Set oSurfaceBody = Nothing
    Set oDirectionVec = Nothing
    Set oFramesColl = Nothing
    Set oFrameSet = Nothing

 Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetParallelAxis(ByVal oSurfaceBody As IJSurfaceBody) As IJDVector
Const METHOD = "GetParallelAxis"
    On Error GoTo ErrorHandler
    
     Dim oCenter             As IJDPosition
     Dim oNormal             As IJDVector
     Dim oFinalNormal        As IJDVector
     Dim oTopoLocate         As IJTopologyLocate
     Dim oTemplateHelper     As MfgTemplateHelper
     
     Set oTopoLocate = New TopologyLocate
     oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oCenter, oNormal
     Set oTopoLocate = Nothing
    
     Dim oXVector As IJDVector
     Dim oYVector As IJDVector
     Dim oZVector As IJDVector
     
     Set oXVector = New DVector
     Set oYVector = New DVector
     Set oZVector = New DVector
     
     oXVector.Set 1, 0, 0
     oYVector.Set 0, 1, 0
     oZVector.Set 0, 0, 1
     
     'normalize the vectors
     oXVector.Length = 1#
     oYVector.Length = 1#
     oZVector.Length = 1#
     
     Dim dX As Double
     Dim dY As Double
     Dim dZ As Double
     
     dX = Abs(oXVector.Dot(oNormal))
     dY = Abs(oYVector.Dot(oNormal))
     dZ = Abs(oZVector.Dot(oNormal))
     
     If (dX > dY) And (dX > dZ) Then
         Set GetParallelAxis = oXVector
     ElseIf (dY > dX) And (dY > dZ) Then
         Set GetParallelAxis = oYVector
     ElseIf (dZ > dX) And (dZ > dY) Then
         Set GetParallelAxis = oZVector
     End If

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
