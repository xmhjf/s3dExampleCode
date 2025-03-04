Attribute VB_Name = "CommonFuncs"
Option Explicit
Private Const sMODULE = "TemplateNamingRule.CommonFuncs"
Private Const E_FAIL = &H80004005      ' For the error object.
Private Const PART_SURFACE = 0
Private Const MOLDED_SURFACE = 1

Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"



Public Sub GetEndPoints(ByVal oEdge As IUnknown, ByRef oStartPosition As IJDPosition, ByRef oEndPosition As IJDPosition)
    Const sMETHOD = "GetEndPoints"
    Dim oErrors As IJEditErrors      ' To collect and propagate the errors.
    'Set oErrors = Nothing
    Set oErrors = CreateObject(ERRORPROGID)
    
    On Error Resume Next

    Dim oCurve As IJCurve
    Dim oWire As IJWireBody
    Set oCurve = oEdge

    If Not oCurve Is Nothing Then
        Set oStartPosition = New DPosition
        Set oEndPosition = New DPosition
        Dim dX1 As Double, dY1 As Double, dZ1 As Double, dX2 As Double, dY2 As Double, dZ2 As Double
        oCurve.EndPoints dX1, dY1, dZ1, dX2, dY2, dZ2
        oStartPosition.Set dX1, dY1, dZ1
        oEndPosition.Set dX2, dY2, dZ2
    End If
    
    Set oWire = oEdge
    
    If Not oWire Is Nothing Then
        oWire.GetEndPoints oStartPosition, oEndPosition
    End If
    
CleanUp:
    Set oCurve = Nothing
    Set oWire = Nothing
    Set oErrors = Nothing
    Exit Sub
    
ErrorHandler:
     oErrors.Add Err.Number, sMODULE & " - " & sMETHOD, Err.Description
     Err.Raise E_FAIL
     GoTo CleanUp
     
    
End Sub

Public Function GetClosestAndLowestFrame(ByVal pPlatePart As IJPlatePart, ByVal pTemplate As IJMfgTemplate, ByVal pTemplateSet As IJDMfgTemplateSet) As Object
    Const sMETHOD = "GetClosestAndLowestFrame"
    Dim oErrors As IJEditErrors      ' To collect and propagate the errors.
    
    Set oErrors = CreateObject(ERRORPROGID)
    
    On Error GoTo ErrorHandler

        Dim oGeomHelper As MfgGeomHelper
        Set oGeomHelper = New MfgGeomHelper
        
        Dim oProcessSettings As IJMfgTemplateProcessSettings
        Set oProcessSettings = pTemplateSet.GetProcessSettings
        
        Dim strDirection As String, strTemplateSide As String
        strDirection = oProcessSettings.TemplateDirection
        strTemplateSide = oProcessSettings.TemplateSide
             
        Dim oPlateSideRule   As IJDMfgPlateUpSideRule
        Dim Upside As enumPlateSide
        Dim oSettingsHelper As MfgSettingsHelper
        Dim strProgId As String
        
        Set oSettingsHelper = oProcessSettings
        strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
        Set oPlateSideRule = SP3DCreateObject(strProgId)
        Upside = oPlateSideRule.GetPlateUpSide(pPlatePart)
    
        Dim ePlateThickness As PlateThicknessSide
        If (Upside = BaseSide) Then
            ePlateThickness = PlateBaseSide
        ElseIf (Upside = OffsetSide) Then
            ePlateThickness = PlateOffsetSide
        End If
        
        Dim lSurfaceType As Long
        lSurfaceType = pTemplateSet.SurfaceType
        Dim esurfaceType As eStrMfgSurfaceType
        If lSurfaceType = PART_SURFACE Then
            esurfaceType = TRUE_PART
        Else
            esurfaceType = TRUE_MOLD
        End If
        
        Dim oSurface As Object
        Set oSurface = oGeomHelper.GetSurfaceFromPlate(pPlatePart, esurfaceType, ePlateThickness, 0)
        
        Dim oVector As IJDVector
        Set oVector = New DVector
        If strDirection = "Transversal" Then
            oVector.Set 1, 0, 0
        ElseIf strDirection = "Longitudinal" Then
            oVector.Set 0, 1, 0
        Else 'Waterline
            oVector.Set 0, 0, 1
        End If
        
        Dim oFrameColl As New Collection
        oGeomHelper.GetReferencePlanesInRange oSurface, oVector, oFrameColl
        
        Dim oControlLine As Object
        Set oControlLine = pTemplateSet.GetControlLine
        
        Dim oIntersectPoint As IJDPosition
        Dim oTemplateRpt As IJMfgTemplateReport
        Set oTemplateRpt = pTemplate
        Set oIntersectPoint = oTemplateRpt.GetCtrlLnMrkLnIntersectPt
        
        'Find out the closest and lowest Frame to current template
        Dim i As Integer
        Dim dDistance As Double, dTempDistance As Double
        Dim oFrame As Object
        dDistance = 100
        For i = 1 To oFrameColl.Count
            Dim oTempIntersectPt As IJDPosition
            Dim oTempFrame As Object
            Set oTempFrame = oFrameColl.Item(i)
            Set oTempIntersectPt = oGeomHelper.IntersectCurveWithPlane(oControlLine, oTempFrame)
            
            If Not oTempIntersectPt Is Nothing Then
                If strDirection = "Transversal" Then
                    If oTempIntersectPt.x <= oIntersectPoint.x Then
                        dTempDistance = oTempIntersectPt.DistPt(oIntersectPoint)
                        If dTempDistance < dDistance Then
                            Set oFrame = Nothing
                            dDistance = dTempDistance
                            Set oFrame = oTempFrame
                        End If
                    End If
                    
                ElseIf strDirection = "Longitudinal" Then
'                    If oTempIntersectPt.y <= oIntersectPoint.y Then
                        dTempDistance = oTempIntersectPt.DistPt(oIntersectPoint)
                        If dTempDistance < dDistance Then
                            Set oFrame = Nothing
                            dDistance = dTempDistance
                            Set oFrame = oTempFrame
                        End If
'                    End If
                Else 'Waterline
'                    If oTempIntersectPt.z <= oIntersectPoint.z Then
                        dTempDistance = oTempIntersectPt.DistPt(oIntersectPoint)
                        If dTempDistance < dDistance Then
                            Set oFrame = Nothing
                            dDistance = dTempDistance
                            Set oFrame = oTempFrame
                        End If
'                    End If
                End If
            End If
            
            Set oTempFrame = Nothing
            Set oTempIntersectPt = Nothing
            
        Next i
        Set GetClosestAndLowestFrame = oFrame
CleanUp:
        Set oGeomHelper = Nothing
        Set oProcessSettings = Nothing
        Set oFrameColl = Nothing
        Set oControlLine = Nothing
        Set oTemplateRpt = Nothing
        Set oIntersectPoint = Nothing
        Set oSettingsHelper = Nothing
        Set oPlateSideRule = Nothing
        Set oErrors = Nothing
            
    Exit Function
ErrorHandler:
    oErrors.Add Err.Number, sMODULE & " - " & sMETHOD, Err.Description
    Err.Raise E_FAIL
    GoTo CleanUp
    
    

End Function

 
Public Function GetParallelAxis(ByVal oSurface As Object) As IJDVector
Const METHOD = "GetParallelAxis"
    On Error GoTo ErrorHandler
    
     Dim oCenter             As IJDPosition
     Dim oNormal             As IJDVector
     Dim oFinalNormal        As IJDVector
     Dim oTopoLocate         As IJTopologyLocate
     Dim oTemplateHelper     As MfgTemplateHelper
     
     
     Dim oSurfaceBody As IJSurfaceBody
     Set oSurfaceBody = oSurface
     
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
     
     Dim dx As Double
     Dim dy As Double
     Dim dz As Double
     
     dx = Abs(oXVector.Dot(oNormal))
     dy = Abs(oYVector.Dot(oNormal))
     dz = Abs(oZVector.Dot(oNormal))
     
     If (dx > dy) And (dx > dz) Then
         Set GetParallelAxis = oXVector
     ElseIf (dy > dx) And (dy > dz) Then
         Set GetParallelAxis = oYVector
     ElseIf (dz > dx) And (dz > dy) Then
         Set GetParallelAxis = oZVector
     End If

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

