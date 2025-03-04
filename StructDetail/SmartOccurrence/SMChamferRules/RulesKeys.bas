Attribute VB_Name = "RulesKeys"
Option Explicit

'Inputs for Chamfer
Public Const CHAMFER_PART = "Part"
Public Const OPPOSITE_PART = "OppositePart"

Public Const gsChamferType = "ChamferType"
Public Const gsChamferWeld = "ChamferWeld"
Public Const gsChamferNumber = "ChamferNumber"
Const PI As Double = 3.141592654

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SMChamferRules\RuleKeys.bas"


'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

Public Function GetChamferCondition(oPart1 As StructDetailObjects.PlatePart, oPart2 As StructDetailObjects.PlatePart, bChamferSlope As Double) As ChamferCondition
Const METHOD = "::IsCanByCanCase"
On Error GoTo ErrorHandler
            'oPart1 is always a Chamfered part as it is set to oAssyConn.Chamfered
            Dim oParentObject As Object
            Dim oPlateSystem As IJPlateSystem
            Dim oSDO_PlatePart As StructDetailObjects.PlatePart
            Dim bFromBuiltUp1 As Boolean
            Dim bFromBuiltUp2 As Boolean
            Dim oSectionType1 As String
            Dim oSectionType2 As String
            
            Dim oSDO_PlateSystem As StructDetailObjects.PlateSystem
            Dim oBuiltupMember As ISPSDesignedMember
                        
            Dim oCan1 As New StructDetailObjects.MemberPart
            Dim oCan2 As New StructDetailObjects.MemberPart
            
            Dim oCanRule As ISPSCanRule
            Dim oSpsCanRuleStatus As SPSCanRuleStatus
              
            Dim oLineDir1 As IJDVector
            Dim pLine As IJLine
            Dim oMemberLine As IJLine
            
            Dim pPOM As IJDPOM
            Set pPOM = Nothing
            
            Dim oPrimaryMemberSystem As ISPSMemberSystem
            Set oPrimaryMemberSystem = Nothing
            
            Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
            Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate

            Dim oModelBody1 As IJModelBody
            Dim oModelBody2 As IJModelBody

            Dim oTopoIntersect As IJDTopologyIntersect
            Set oTopoIntersect = New DGeomOpsIntersect
            
'            check if only the chamfered part is a Tube portion of CAN. if yes, return TRUE.
            If TypeOf oPart1 Is PlatePart Then
                Set oParentObject = oPart1.ParentSystem
                If TypeOf oParentObject Is IJPlateSystem Then
                    Set oSDO_PlateSystem = New StructDetailObjects.PlateSystem
                    Set oSDO_PlateSystem.object = oParentObject
                    Set oParentObject = oSDO_PlateSystem.ParentSystem
                    Set oSDO_PlateSystem = Nothing
                    If TypeOf oParentObject Is IJPlateSystem Then
                        Set oPlateSystem = oParentObject
                        bFromBuiltUp1 = oPlateSystem.IsBuiltupPlateSystem
                        If bFromBuiltUp1 Then
                            Set oBuiltupMember = oPlateSystem.ParentBuiltup
                            Set oCan1.object = oBuiltupMember
                            oSectionType1 = oCan1.SectionType
                            If oSectionType1 = "BUCan" Then
                              bChamferSlope = oCan1.CanChamferSlope
                              Set oCanRule = oCan1.CanRule
                              oSpsCanRuleStatus = oCanRule.GetPrimaryMemberSystem(oPrimaryMemberSystem)
                            End If
                        End If
                    End If
                End If
            End If
           'consider the section type of oPart2.
           If TypeOf oPart2 Is PlatePart Then
                Set oParentObject = oPart2.ParentSystem
                If TypeOf oParentObject Is IJPlateSystem Then
                    Set oSDO_PlateSystem = New StructDetailObjects.PlateSystem
                    Set oSDO_PlateSystem.object = oParentObject
                    Set oParentObject = oSDO_PlateSystem.ParentSystem
                    Set oSDO_PlateSystem = Nothing
                    If TypeOf oParentObject Is IJPlateSystem Then
                        Set oPlateSystem = oParentObject
                        bFromBuiltUp2 = oPlateSystem.IsBuiltupPlateSystem
                        If bFromBuiltUp2 Then
                            Set oBuiltupMember = oPlateSystem.ParentBuiltup
                            Set oCan2.object = oBuiltupMember
                            oSectionType2 = oCan2.SectionType
                        End If
                    End If
                End If
            End If

            If Not oPrimaryMemberSystem Is Nothing Then
                'get the direction of the PrimaryMemberSystem
                Set pLine = oPrimaryMemberSystem.LogicalAxis.CurveGeometry
                Set oMemberLine = Line_FromPositions(pPOM, Position_FromLine(pLine, 0), Position_FromLine(pLine, 1))
                'This gives the Direction of the PrimaryMemberSystem along which the CAN is placed.
                Set oLineDir1 = Vector_FromLine(oMemberLine)
                
                Set oModelBody1 = oTopologyLocate.GetPlateParentBodyModel(oPart1.ParentSystem)
                Set oModelBody2 = oTopologyLocate.GetPlateParentBodyModel(oPart2.ParentSystem)

                Dim oSurfaceBody1 As IJSurfaceBody
                Dim oSurfaceBody2 As IJSurfaceBody

                Dim oShipGeomOps As GSCADShipGeomOps.SGOModelBodyUtilities
                Set oShipGeomOps = New GSCADShipGeomOps.SGOModelBodyUtilities

                Dim ppPointOnFirstBody As IJDPosition
                Dim ppPointOnSecondBody As IJDPosition
                Dim ppdistance As Double

                oShipGeomOps.GetClosestPointsBetweenTwoBodies oModelBody1, oModelBody2, ppPointOnFirstBody, ppPointOnSecondBody, ppdistance

                Dim ppNormal1 As IJDVector
                Dim ppNormal2 As IJDVector

                If TypeOf oModelBody1 Is IJSurfaceBody Then
                    Set oSurfaceBody1 = oModelBody1
                    oSurfaceBody1.GetNormalFromPosition ppPointOnFirstBody, ppNormal1
                End If

                If TypeOf oModelBody2 Is IJSurfaceBody Then
                    Set oSurfaceBody2 = oModelBody2
                    oSurfaceBody2.GetNormalFromPosition ppPointOnSecondBody, ppNormal2
                End If

                ppNormal1.Length = 1
                ppNormal2.Length = 1

                If Not oLineDir1 Is Nothing Then
                    If oSectionType2 = "BUCan" Then
                        If Round(oLineDir1.Dot(ppNormal1), 3) = 0 And Not Round(oLineDir1.Dot(ppNormal2), 3) = 0 Then
                            GetChamferCondition = CanCylinderToCone
                        ElseIf Not Round(oLineDir1.Dot(ppNormal1), 3) = 0 And Round(oLineDir1.Dot(ppNormal2), 3) = 0 Then
                            GetChamferCondition = CanConeToCylinder
                        End If
                    Else
                        If Not Round(oLineDir1.Dot(ppNormal1), 3) = 0 Then
                            GetChamferCondition = CanConeToMember
                        ElseIf Round(oLineDir1.Dot(ppNormal1), 3) = 0 And Round(oLineDir1.Dot(ppNormal2), 3) = 0 Then
                            GetChamferCondition = CanCylinderToMember
                        End If
                    End If
                Else
                    GetChamferCondition = Standard
                End If
            Else
                GetChamferCondition = Standard
            End If
            
            Set oParentObject = Nothing
            Set oPlateSystem = Nothing
            Set oSDO_PlatePart = Nothing
            Set oSDO_PlateSystem = Nothing
            Set oBuiltupMember = Nothing
            Set oCan1 = Nothing
            Set oCan2 = Nothing
            Set oPrimaryMemberSystem = Nothing
      Exit Function
           
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
Function Position_FromLine(pLine As IJLine, iIndex As Integer) As IJDPosition
    ' create new position
    Dim pPosition As IJDPosition
    Set pPosition = New DPosition
    
    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    If iIndex = 0 Then
        Call pLine.GetStartPoint(x, y, z)
    Else
        Call pLine.GetEndPoint(x, y, z)
    End If
    
    ' set coordinates
    Call pPosition.Set(x, y, z)
    
    ' return result
    Set Position_FromLine = pPosition
End Function
Public Function Line_FromPositions(pPOM As IJDPOM, pPositionOfStartPoint As IJDPosition, pPositionOfEndPoint As IJDPosition) As IJLine
    ' create new line
    Dim pGeometryFactory As New GeometryFactory
    Set Line_FromPositions = pGeometryFactory.Lines3d.CreateBy2Points(pPOM, _
        pPositionOfStartPoint.x, pPositionOfStartPoint.y, pPositionOfStartPoint.z, _
        pPositionOfEndPoint.x, pPositionOfEndPoint.y, pPositionOfEndPoint.z)
End Function
Function Vector_FromLine(pLine As IJLine) As IJDVector
    ' get vector
    Dim u As Double, v As Double, w As Double
    Call pLine.GetDirection(u, v, w)
    
    ' create result
    Dim pVector As New DVector
    Call pVector.Set(u, v, w)
    
    ' return result
    Set Vector_FromLine = pVector
End Function
Public Sub GetAdjustmentValue(dSlope As Double, dAngle As Double, dThicknessDiff As Double, dOffset As Double)
'dSLope is ChmaferSlope
'(90-dAngle) is Cone slope
Dim oLength As Double
Dim oConeSlope As Double
Dim oSlopeDiff As Double
Dim oChamferAngleWithEdge As Double
Dim dAngleG As Double

If dSlope > (PI / 2) - dAngle Then
    oConeSlope = (PI / 2) - (dAngle)
    oLength = dThicknessDiff / Cos(oConeSlope)
    oSlopeDiff = Abs(dSlope - oConeSlope)
    oChamferAngleWithEdge = PI - (oConeSlope + oSlopeDiff) - (dAngle / 2)
    dOffset = oLength * Sin(oSlopeDiff) / Sin(oChamferAngleWithEdge)
Else
    oConeSlope = (PI / 2) - (dAngle)
    oLength = dThicknessDiff / Cos(oConeSlope)
    oSlopeDiff = Abs(dSlope - oConeSlope)
    oChamferAngleWithEdge = PI - (oConeSlope) - (dAngle / 2)
    dAngleG = PI - oSlopeDiff - oChamferAngleWithEdge
    dOffset = oLength * Sin(oSlopeDiff) / Sin(dAngleG)
End If


Exit Sub
End Sub
Public Function DegToRad(dAngle As Double) As Double
    DegToRad = dAngle * PI / 180  'Radians=Degrees*Pi/180
End Function

