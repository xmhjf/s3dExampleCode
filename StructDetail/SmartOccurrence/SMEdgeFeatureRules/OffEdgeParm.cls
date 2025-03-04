VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OffEdgeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID & "EdgeFeatureRules.OffEdgeParm"
Const m_ParameterRuleName As String = CUSTOMERID & "EdgeFeatureRules.OffEdgeParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "StructDetail\SmartOccurrence\" & CUSTOMERID & "EdgeFeatureRules\OffEdgeParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  'Add port inputs
    pIH.SetInput INPUT_EDGE
    pIH.SetInput INPUT_POINT
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

 pOH.SetOutput "AdjustedOffset"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
    
    
    
    Dim oEdgeFeatureWrapper As StructDetailObjects.EdgeFeature
    Set oEdgeFeatureWrapper = New StructDetailObjects.EdgeFeature
    
    Set oEdgeFeatureWrapper.object = pPRL.SmartOccurrence
            
    Dim oEdgeHelper As StructDetailObjects.Helper
    Set oEdgeHelper = New StructDetailObjects.Helper
    
    'Getting offset
    Dim dOffset As Double
    Dim dAdjustedOffset As Double
    Dim dBevelDepth As Double
    Dim dClearanceOffBevel As Double
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String
    
    oEdgeHelper.GetCustomAttributeParamValue oEdgeFeatureWrapper.object, "IJUASmartEdgeFeature", "Offset", dOffset, _
                                                         DoubleType, sTableName, lCodeList, sShortDesc, sLongDesc
    
    dClearanceOffBevel = 0.005 'Clearance from Bevel
                                                         
                                                         
    Dim oPort As IJPort
    Dim oPoint As Object
    
    Set oPort = pPRL.InputObject(INPUT_EDGE)
    Set oPoint = pPRL.InputObject(INPUT_POINT)
    
    'Getting PC on which Edge Feature is dependent
    Dim oPC As StructDetailObjects.PhysicalConn
    Dim oPCObject As Object
    Set oPCObject = GetPCOnWhichEFIsDependent(oEdgeFeatureWrapper.object, oPort, oPoint)
    
    If Not oPCObject Is Nothing Then
        Set oPC = New StructDetailObjects.PhysicalConn
        Set oPC.object = oPCObject
    End If
    
    Dim dRefSideBevelDepth As Double
    Dim dAntiRefSideBevelDepth As Double
    Dim bIsEFOnRefPart As Boolean
    Dim sRefPartName As String
    dBevelDepth = 0
    bIsEFOnRefPart = False
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim strPCName As String
    
    'Checking in case of ButtWelds, whether EF is placed on Ref Part or Non Ref Part
    If Not oPC Is Nothing Then
        
        If oPC.ConnectionSubType = SMARTSUBTYPE_BUTTWELD Then
            Dim oNamedItem As IJNamedItem
            Set oNamedItem = oEdgeFeatureWrapper.GetPartObject
        
            oPC.GetSymbolParameterStringValue "RefPartName", sRefPartName
        
            If LCase$(Trim$(oNamedItem.Name)) = LCase$(Trim$(sRefPartName)) Then
                bIsEFOnRefPart = True
            Else
                bIsEFOnRefPart = False
            End If
        End If
    End If
    
    If Not oPC Is Nothing Then
        GetBevelDepth oPC, dRefSideBevelDepth, dAntiRefSideBevelDepth, bIsEFOnRefPart
        
        If GreaterThan(dRefSideBevelDepth, dAntiRefSideBevelDepth) Then
            dBevelDepth = dRefSideBevelDepth
        Else
            dBevelDepth = dAntiRefSideBevelDepth
        End If
    End If
    
    If GreaterThan((dBevelDepth + dClearanceOffBevel), dOffset) Then
        dAdjustedOffset = dBevelDepth + dClearanceOffBevel
    Else
        dAdjustedOffset = dOffset
    End If
                    
    pPRL.Add "AdjustedOffset", dAdjustedOffset
    
    
    Set oEdgeFeatureWrapper = Nothing
    Set oEdgeHelper = Nothing
    Set oPoint = Nothing
    Set oPort = Nothing
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

Private Function GetPCOnWhichEFIsDependent(oEdgeFeature As Object, oPort As IJPort, oPoint As Object) As Object
    
    
    'Retrieving input point
    Dim oP3d As IngrGeom3D.Point3d
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Set oP3d = New Point3d
    
    If TypeOf oPoint Is Point3d Then
        Set oP3d = oPoint
    Else 'Considering case when placed using seam or knuckle method
        Dim oTempPos As IJDPosition
        Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
        Set oTopologyLocate = New TopologyLocate
        
        Set oTempPos = oTopologyLocate.FindIntersectionPoint(oPort, oPoint)
        oP3d.SetPoint oTempPos.x, oTempPos.y, oTempPos.z
    End If
    
    'Input point
    oP3d.GetPoint dx, dy, dz
    
    'Getting target object for Edge Feature
    Dim oEdgeFeatureWrapper As New StructDetailObjects.EdgeFeature
    Dim oPartWithFeature As Object
    
    Set oEdgeFeatureWrapper.object = oEdgeFeature
    Set oPartWithFeature = oEdgeFeatureWrapper.GetPartObject
    
    'Variables and objects used for retrieving required PC
    Dim oPC As StructDetailObjects.PhysicalConn
    Dim lPCCount As Long
    Dim exitflag As Boolean
    Dim jCount As Long
    Dim oWB As IJWireBody
    Dim oComplexStrings As IJElements
    Dim oCurve As IJCurve
    Dim dMinDist As Double
    Dim dSrcX As Double
    Dim dSrcY As Double
    Dim dSrCZ As Double
    Dim dInX As Double
    Dim dInY As Double
    Dim dInZ As Double
    
    Set oPC = New StructDetailObjects.PhysicalConn
    
    exitflag = False
    
    'On Plates, Get LateralFacePort on which EF is dependent and then find PC where
    
    If TypeOf oPartWithFeature Is IJPlatePart Then
        Dim oTempPart As IJPlatePart
        Set oTempPart = oPort.Connectable
        Dim oFacePort As IJPort
        
        Dim bFacePortFound As Boolean
        bFacePortFound = False
    
        'Getting operator and operation ID of input port
        Dim lOperatorID As Long
        Dim lOperationID As Long
        Dim oStPort As IJStructPort
        Set oStPort = oPort
    
        lOperatorID = oStPort.OperatorID
        lOperationID = oStPort.OperationID
        
        If TypeOf oTempPart Is IJStructConnectable Then
            Dim oCollectionOfPorts As IJElements
            Dim oStructConnectable As IJStructConnectable
            Set oStructConnectable = oTempPart
                        
            oStructConnectable.enumConnectableTransientPorts oCollectionOfPorts, vbNullString, False, 0, JS_TOPOLOGY_FILTER_LCONNECT_PRT_LFACES, False
            
            If Not oCollectionOfPorts Is Nothing Then
                If oCollectionOfPorts.Count > 0 Then
                    Dim oTransientPort As IMSStructConnection.TransientPort
                    Dim oTempObject As Object
                    
                    Dim lFaceCount As Long
                    For lFaceCount = 1 To oCollectionOfPorts.Count
                        Set oStPort = oCollectionOfPorts.Item(lFaceCount)
                        If lOperationID = oStPort.OperationID And lOperatorID = oStPort.OperatorID Then
                            Set oTransientPort = oStPort
                            oTransientPort.GetPersistentPort oTempObject
                            Set oFacePort = oTempObject
                            If Not oFacePort Is Nothing Then
                                bFacePortFound = True
                                'Since Operation ID and Operator ID matched, we got required face port
                                Exit For
                            End If
                        End If
                    Next lFaceCount
                End If
            End If
        End If
        
        If Not oFacePort Is Nothing Then
            Dim oCollectionofPCs As IJElements
            oFacePort.enumConnections oCollectionofPCs, ConnectionPhysical, ConnectionStandard
            If Not oCollectionofPCs Is Nothing Then
                If oCollectionofPCs.Count > 0 Then
                    For lPCCount = 1 To oCollectionofPCs.Count
                        Set oPC.object = oCollectionofPCs(lPCCount)
        
                        If TypeOf oPC.object Is IJWireBody And _
                            (oPC.ConnectionSubType = SMARTSUBTYPE_BUTTWELD Or oPC.ConnectionSubType = SMARTSUBTYPE_TEEWELD) Then
                            Set oWB = oPC.object
            
                            oWB.GetComplexStrings Nothing, oComplexStrings
                            For jCount = 1 To oComplexStrings.Count
                                Set oCurve = oComplexStrings.Item(jCount)
                
                                exitflag = oCurve.IsPointOn(dx, dy, dz)
                                oCurve.DistanceBetween oP3d, dMinDist, dSrcX, dSrcY, dSrCZ, dInX, dInY, dInZ
                                If exitflag = False Then
                                    'if minimum distance between PC and point is less than 1 mm then given PC is desired PC
                                    If LessThan(dMinDist, 0.001) Then exitflag = True
                                    
                                End If
                                
                                'Since point lies on PC, desired PC is retrieved
                                If exitflag = True Then Exit For
                            Next jCount
                        End If
                        If exitflag = True Then Exit For
                        'Since desired PC is found, we can come out of loop
                    Next lPCCount
                End If
            End If
        End If
    ElseIf TypeOf oPartWithFeature Is IJProfilePart Then
        Dim nCount As Long
        Dim zConnectionData() As ConnectionData
        
        Dim oEdgeHelper As StructDetailObjects.Helper
        Set oEdgeHelper = New StructDetailObjects.Helper
        
        oEdgeHelper.Object_AppConnections oPort.Connectable, AppConnectionType_Physical, nCount, zConnectionData()
        
        For lPCCount = 1 To nCount
            Dim oStPCCon As IJStructPhysicalConnection
            Set oStPCCon = zConnectionData(lPCCount).AppConnection
            If Not oStPCCon Is Nothing Then
                Set oPC.object = oStPCCon
            
                If TypeOf oPC.object Is IJWireBody And _
                    (oPC.ConnectionSubType = SMARTSUBTYPE_BUTTWELD Or oPC.ConnectionSubType = SMARTSUBTYPE_TEEWELD) Then
                    
                    Set oWB = oPC.object
                    
                    oWB.GetComplexStrings Nothing, oComplexStrings
                    
                    For jCount = 1 To oComplexStrings.Count
                        Set oCurve = oComplexStrings.Item(jCount)
                        
                        oCurve.DistanceBetween oP3d, dMinDist, dSrcX, dSrcY, dSrCZ, dInX, dInY, dInZ
                        
                        'If point falls on PC wirebody then current PC is the one we are interested in
                        If LessThan(dMinDist, 0.001) Then exitflag = True
                        
                        If exitflag = True Then Exit For
                    Next jCount
                End If
               If exitflag = True Then Exit For
            End If
        Next lPCCount
    End If
            
    If exitflag = True Then
        Set GetPCOnWhichEFIsDependent = oPC.object
    Else
        Set GetPCOnWhichEFIsDependent = Nothing
    End If
    
    
End Function

Private Sub GetBevelDepth(oPC As StructDetailObjects.PhysicalConn, dRefSideBevelDepth As Double, dAntiRefSideBevelDepth As Double, bIsEFOnRefPart As Boolean)
    Dim oSmartItem As IJSmartItem
    Dim oSmartOcc As IJSmartOccurrence
    Dim strPCName As String
    Dim PI As Double
    
    PI = GetPI
    
    Dim dTolerance As Double
    dTolerance = 0.0001
    
    'Defining variables to retrive bevel parameters
    Dim dRootGap As Double
    Dim dRefSideFirstBevelDepth As Double
    Dim dRefSideFirstBevelAngle As Double
    Dim dAntiRefSideFirstBevelDepth As Double
    Dim dAntiRefSideFirstBevelAngle As Double
    Dim dRefSideSecondBevelDepth As Double
    Dim dRefSideSecondBevelAngle As Double
    Dim dAntiRefSideSecondBevelDepth As Double
    Dim dAntiRefSideSecondBevelAngle As Double
    Dim dNoseOrientationAngle As Double
    
    Dim dNRRootGap As Double
    Dim dNRRefSideFirstBevelDepth As Double
    Dim dNRRefSideFirstBevelAngle As Double
    Dim dNRAntiRefSideFirstBevelDepth As Double
    Dim dNRAntiRefSideFirstBevelAngle As Double
    Dim dNRRefSideSecondBevelDepth As Double
    Dim dNRRefSideSecondBevelAngle As Double
    Dim dNRAntiRefSideSecondBevelDepth As Double
    Dim dNRAntiRefSideSecondBevelAngle As Double
    Dim dNRNoseOrientationAngle As Double
    
    Dim sRefSideFirstBevelMethod As String
    Dim sAntiRefSideFirstBevelMethod As String
    Dim sRefSideSecondBevelMethod As String
    Dim sAntiRefSideSecondBevelMethod As String
        
    Dim dRootGapCorrection As Double
    Dim dTeeMountingAngle As Double
    Dim dCornerButtMountingAngle As Double
    Dim dAngle As Double
    
    Dim dRefSideBevelDepthCorrection As Double
    Dim dAntiRefSideBevelDepthCorrection As Double
    
    Dim dNRRefSideBevelDepthCorrection As Double
    Dim dNRAntiRefSideBevelDepthCorrection As Double
        
    If oPC.ConnectionSubType = SMARTSUBTYPE_TEEWELD Then
        oPC.GetBevelParameterValue "RefSideFirstBevelDepth", dRefSideFirstBevelDepth, DoubleType
        oPC.GetBevelParameterValue "RefSideFirstBevelAngle", dRefSideFirstBevelAngle, DoubleType
        oPC.GetBevelParameterValue "RefSideFirstBevelMethod", sRefSideFirstBevelMethod, StringType
        
        oPC.GetBevelParameterValue "AntiRefSideFirstBevelDepth", dAntiRefSideFirstBevelDepth, DoubleType
        oPC.GetBevelParameterValue "AntiRefSideFirstBevelAngle", dAntiRefSideFirstBevelAngle, DoubleType
        oPC.GetBevelParameterValue "AntiRefSideFirstBevelMethod", sAntiRefSideFirstBevelMethod, StringType
        
        oPC.GetBevelParameterValue "RefSideSecondBevelDepth", dRefSideSecondBevelDepth, DoubleType
        oPC.GetBevelParameterValue "RefSideSecondBevelAngle", dRefSideSecondBevelAngle, DoubleType
        oPC.GetBevelParameterValue "RefSideSecondBevelMethod", sRefSideSecondBevelMethod, StringType
        
        oPC.GetBevelParameterValue "AntiRefSideSecondBevelDepth", dAntiRefSideSecondBevelDepth, DoubleType
        oPC.GetBevelParameterValue "AntiRefSideSecondBevelAngle", dAntiRefSideSecondBevelAngle, DoubleType
        oPC.GetBevelParameterValue "AntiRefSideSecondBevelMethod", sAntiRefSideSecondBevelMethod, StringType
        
        oPC.GetBevelParameterValue "RootGap", dRootGap, DoubleType
        oPC.GetBevelParameterValue "NoseOrientationAngle", dNoseOrientationAngle, DoubleType
        
        
        dTeeMountingAngle = oPC.TeeMountingAngle
        
        If GreaterThan(dTeeMountingAngle, PI, dTolerance) Then
            dTeeMountingAngle = 2 * PI - dTeeMountingAngle
        End If
        
        'Changing Bevel angle parameters from Varying method to constant
        If LCase$(Trim$(sRefSideFirstBevelMethod)) = LCase$(Trim$("Varying")) _
                        And GreaterThanZero(dRefSideFirstBevelDepth) Then
                dRefSideFirstBevelAngle = PI / 2 + dRefSideFirstBevelAngle - dTeeMountingAngle
        End If
        
        If LCase$(Trim$(sAntiRefSideFirstBevelMethod)) = LCase$(Trim$("Varying")) _
                        And GreaterThanZero(dAntiRefSideFirstBevelDepth) Then
                dAntiRefSideFirstBevelAngle = dAntiRefSideFirstBevelAngle + dTeeMountingAngle - PI / 2
        End If
        
        If LCase$(Trim$(sRefSideSecondBevelMethod)) = LCase$(Trim$("Varying")) _
                        And GreaterThanZero(dRefSideSecondBevelDepth) Then
                dRefSideSecondBevelAngle = PI / 2 + dRefSideSecondBevelAngle - dTeeMountingAngle
        End If
        
        If LCase$(Trim$(sAntiRefSideSecondBevelMethod)) = LCase$(Trim$("Varying")) _
                        And GreaterThanZero(dAntiRefSideSecondBevelDepth) Then
                dAntiRefSideSecondBevelAngle = dAntiRefSideSecondBevelAngle + dTeeMountingAngle - PI / 2
        End If
        
        dRefSideBevelDepth = 0
        dRootGapCorrection = 0
        dAntiRefSideBevelDepth = 0
        
        If GreaterThanZero(Sin(dNoseOrientationAngle)) Then
        'if dNoseOrientationAngle is zero then RootGap correction will be automatically
        'incorporated in RefSideBevelDepth or AntiRefSideBevelDepth
            dRootGapCorrection = dRootGap / Sin(dNoseOrientationAngle)
        End If
                        
        'Since BevelDepths are calculated as if bevel portions are projected along bounded part
        'Another correction is required to take bevel projections perpendicular to bounded part
        'and mounting angle into account
        
        If Not Equal(dTeeMountingAngle, PI / 2, dTolerance) Then
            dRefSideBevelDepthCorrection = -1 * (dRefSideFirstBevelDepth + dRefSideSecondBevelDepth) / Tan(dTeeMountingAngle)
            dAntiRefSideBevelDepthCorrection = (dAntiRefSideFirstBevelDepth + dAntiRefSideSecondBevelDepth) / Tan(dTeeMountingAngle)
        Else
            dRefSideBevelDepthCorrection = 0
            dAntiRefSideBevelDepthCorrection = 0
        End If
        
        'RefSide Bevel Depth
        If Equal(dRefSideFirstBevelAngle, PI / 2, dTolerance) Then
            dRefSideBevelDepth = dRefSideBevelDepth + dRefSideFirstBevelDepth
        Else
            dRefSideBevelDepth = dRefSideBevelDepth + dRefSideFirstBevelDepth * Tan(dRefSideFirstBevelAngle)
        End If
        
        If Equal(dRefSideSecondBevelAngle, PI / 2, dTolerance) Then
            dRefSideBevelDepth = dRefSideBevelDepth + dRefSideSecondBevelDepth
        Else
            dRefSideBevelDepth = dRefSideBevelDepth + dRefSideSecondBevelDepth * Tan(dRefSideSecondBevelAngle)
        End If
        
        dRefSideBevelDepth = dRefSideBevelDepth + dRootGapCorrection + dRefSideBevelDepthCorrection
        
        'AntiRefSide Bevel Depth
        If Equal(dAntiRefSideFirstBevelAngle, PI / 2, dTolerance) Then
            dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dAntiRefSideFirstBevelDepth
        Else
            dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dAntiRefSideFirstBevelDepth * Tan(dAntiRefSideFirstBevelAngle)
        End If
        
        If Equal(dAntiRefSideSecondBevelAngle, PI / 2, dTolerance) Then
            dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dRefSideSecondBevelDepth
        Else
            dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dRefSideSecondBevelDepth * Tan(dAntiRefSideSecondBevelAngle)
        End If
        
        dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dRootGapCorrection + dAntiRefSideBevelDepthCorrection
        
    ElseIf oPC.ConnectionSubType = SMARTSUBTYPE_BUTTWELD Then
        If bIsEFOnRefPart = True Then
            oPC.GetBevelParameterValue "RefSideFirstBevelDepth", dRefSideFirstBevelDepth, DoubleType
            oPC.GetBevelParameterValue "RefSideFirstBevelAngle", dRefSideFirstBevelAngle, DoubleType
                        
            oPC.GetBevelParameterValue "AntiRefSideFirstBevelDepth", dAntiRefSideFirstBevelDepth, DoubleType
            oPC.GetBevelParameterValue "AntiRefSideFirstBevelAngle", dAntiRefSideFirstBevelAngle, DoubleType
                        
            oPC.GetBevelParameterValue "RefSideSecondBevelDepth", dRefSideSecondBevelDepth, DoubleType
            oPC.GetBevelParameterValue "RefSideSecondBevelAngle", dRefSideSecondBevelAngle, DoubleType
                        
            oPC.GetBevelParameterValue "AntiRefSideSecondBevelDepth", dAntiRefSideSecondBevelDepth, DoubleType
            oPC.GetBevelParameterValue "AntiRefSideSecondBevelAngle", dAntiRefSideSecondBevelAngle, DoubleType
                        
            oPC.GetBevelParameterValue "RootGap", dRootGap, DoubleType
            oPC.GetBevelParameterValue "NoseOrientationAngle", dNoseOrientationAngle, DoubleType
            
            dCornerButtMountingAngle = oPC.CornerButtMountingAngle
        
            If GreaterThan(dCornerButtMountingAngle, PI / 2, dTolerance) Then
                dCornerButtMountingAngle = PI - dCornerButtMountingAngle
            End If
            
            'Angle between Plates and Nose
            dAngle = (PI / 2 - (dCornerButtMountingAngle / 2))
            
            dRefSideBevelDepth = 0
            dRootGapCorrection = 0
            dAntiRefSideBevelDepth = 0
            
            dRefSideBevelDepthCorrection = 0
            dAntiRefSideBevelDepthCorrection = 0
                    
            If GreaterThanZero(Sin(dNoseOrientationAngle)) Then
            'if dNoseOrientationAngle is zero then RootGap correction will be automatically
            'incorporated in RefSideBevelDepth or AntiRefSideBevelDepth
                dRootGapCorrection = dRootGap / Sin(dNoseOrientationAngle)
            End If
            
            'Since BevelDepths are calculated as if bevel portions are projected along bounded part
            'Another correction is required to take bevel projections perpendicular to bounded part
            'and mounting angle into account
            If Not Equal(dAngle, 0, 0.0001) Then
                If Not Equal(dAngle, PI / 2, dTolerance) Then
                    dRefSideBevelDepthCorrection = Abs((dRefSideFirstBevelDepth + dRefSideSecondBevelDepth) / Tan(dAngle))
                    dAntiRefSideBevelDepthCorrection = Abs((dAntiRefSideFirstBevelDepth + dAntiRefSideSecondBevelDepth) / Tan(dAngle))
                End If
            End If
            
            If oPC.ConnectedObject1Type = SDOBJECT_PLATE And oPC.ConnectedObject2Type = SDOBJECT_PLATE Then
            
                'Checking if Base in inside face of knuckle or not
                Dim oPartInfo As New GSCADStructGeomUtilities.PartInfo
                If (oPartInfo.IsBaseFaceInsideOfKnuckleAtPhysicalConnection(oPC.object)) Then
                    dRefSideBevelDepth = dRefSideBevelDepth * -1
                Else
                    dAntiRefSideBevelDepth = dAntiRefSideBevelDepth * -1
                End If
                Set oPartInfo = Nothing
            Else
                'We will use approximation
                If GreaterThan(dRefSideBevelDepthCorrection, dAntiRefSideBevelDepthCorrection) Then
                    dAntiRefSideBevelDepthCorrection = dRefSideBevelDepthCorrection
                Else
                    dRefSideBevelDepthCorrection = dAntiRefSideBevelDepthCorrection
                End If
            End If
                        
            'RefSide Bevel Depth
            If Equal(dRefSideFirstBevelAngle, PI / 2, dTolerance) Then
                dRefSideBevelDepth = dRefSideBevelDepth + dRefSideFirstBevelDepth
            Else
                dRefSideBevelDepth = dRefSideBevelDepth + dRefSideFirstBevelDepth * Tan(dRefSideFirstBevelAngle)
            End If
            
            If Equal(dRefSideSecondBevelAngle, PI / 2, dTolerance) Then
                dRefSideBevelDepth = dRefSideBevelDepth + dRefSideSecondBevelDepth
            Else
                dRefSideBevelDepth = dRefSideBevelDepth + dRefSideSecondBevelDepth * Tan(dRefSideSecondBevelAngle)
            End If
            
            dRefSideBevelDepth = dRefSideBevelDepth + dRootGapCorrection + dRefSideBevelDepthCorrection
            
            'AntiRefSide Bevel Depth
            If Equal(dAntiRefSideFirstBevelAngle, PI / 2, dTolerance) Then
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dAntiRefSideFirstBevelDepth
            Else
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dAntiRefSideFirstBevelDepth * Tan(dAntiRefSideFirstBevelAngle)
            End If
            
            If Equal(dAntiRefSideSecondBevelAngle, PI / 2, dTolerance) Then
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dRefSideSecondBevelDepth
            Else
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dRefSideSecondBevelDepth * Tan(dAntiRefSideSecondBevelAngle)
            End If
            
            dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dRootGapCorrection + dAntiRefSideBevelDepthCorrection
           
        Else
            oPC.GetBevelParameterValue "NRRefSideFirstBevelDepth", dNRRefSideFirstBevelDepth, DoubleType
            oPC.GetBevelParameterValue "NRRefSideFirstBevelAngle", dNRRefSideFirstBevelAngle, DoubleType
            
            oPC.GetBevelParameterValue "NRAntiRefSideFirstBevelDepth", dNRAntiRefSideFirstBevelDepth, DoubleType
            oPC.GetBevelParameterValue "NRAntiRefSideFirstBevelAngle", dNRAntiRefSideFirstBevelAngle, DoubleType
            
            oPC.GetBevelParameterValue "NRRefSideSecondBevelDepth", dNRRefSideSecondBevelDepth, DoubleType
            oPC.GetBevelParameterValue "NRRefSideSecondBevelAngle", dNRRefSideSecondBevelAngle, DoubleType
            
            oPC.GetBevelParameterValue "NRAntiRefSideSecondBevelDepth", dNRAntiRefSideSecondBevelDepth, DoubleType
            oPC.GetBevelParameterValue "NRAntiRefSideSecondBevelAngle", dNRAntiRefSideSecondBevelAngle, DoubleType
            
            oPC.GetBevelParameterValue "NRRootGap", dNRRootGap, DoubleType
            oPC.GetBevelParameterValue "NRNoseOrientationAngle", dNRNoseOrientationAngle, DoubleType
        
            dCornerButtMountingAngle = oPC.CornerButtMountingAngle
        
            If GreaterThan(dCornerButtMountingAngle, PI / 2, dTolerance) Then
                dCornerButtMountingAngle = PI - dCornerButtMountingAngle
            End If
            
            'Angle between Plates and Nose
            dAngle = (PI / 2 - (dCornerButtMountingAngle / 2))
            
            dRefSideBevelDepth = 0
            dRootGapCorrection = 0
            dAntiRefSideBevelDepth = 0
            dNRRefSideBevelDepthCorrection = 0
            dNRAntiRefSideBevelDepthCorrection = 0
        
            If GreaterThanZero(Sin(dNoseOrientationAngle)) Then
            'if dNoseOrientationAngle is zero then RootGap correction will be automatically
            'incorporated in RefSideBevelDepth or AntiRefSideBevelDepth
                dRootGapCorrection = dRootGap / Sin(dNoseOrientationAngle)
            End If
            
            'Since BevelDepths are calculated as if bevel portions are projected along bounded part
            'Another correction is required to take bevel projections perpendicular to bounded part
            'and mounting angle into account
            If Not Equal(dAngle, 0, 0.0001) Then
                If Not Equal(dAngle, PI / 2, dTolerance) Then
                    dNRRefSideBevelDepthCorrection = Abs((dNRRefSideFirstBevelDepth + dNRRefSideSecondBevelDepth) / Tan(dAngle))
                    dNRAntiRefSideBevelDepthCorrection = Abs((dNRAntiRefSideFirstBevelDepth + dNRAntiRefSideSecondBevelDepth) / Tan(dAngle))
                End If
            End If
            
            
            If oPC.ConnectedObject1Type = SDOBJECT_PLATE And oPC.ConnectedObject2Type = SDOBJECT_PLATE Then
            
                'Checking if Base in inside face of knuckle or not
                Dim oNRPartinfo As New GSCADStructGeomUtilities.PartInfo
                If (oNRPartinfo.IsBaseFaceInsideOfKnuckleAtPhysicalConnection(oPC.object)) Then
                    dRefSideBevelDepth = dRefSideBevelDepth * -1
                Else
                    dAntiRefSideBevelDepth = dAntiRefSideBevelDepth * -1
                End If
                Set oNRPartinfo = Nothing
            Else
                'We will use approximation
                If GreaterThan(dRefSideBevelDepthCorrection, dAntiRefSideBevelDepthCorrection) Then
                    dAntiRefSideBevelDepthCorrection = dRefSideBevelDepthCorrection
                Else
                    dRefSideBevelDepthCorrection = dAntiRefSideBevelDepthCorrection
                End If
            End If
            
            'RefSide Bevel Depth
            If Equal(dNRRefSideFirstBevelAngle, PI / 2, dTolerance) Then
                dRefSideBevelDepth = dRefSideBevelDepth + dNRRefSideFirstBevelDepth
            Else
                dRefSideBevelDepth = dRefSideBevelDepth + dNRRefSideFirstBevelDepth * Tan(dNRRefSideFirstBevelAngle)
            End If
            
            If Equal(dNRRefSideSecondBevelAngle, PI / 2, dTolerance) Then
                dRefSideBevelDepth = dRefSideBevelDepth + dNRRefSideSecondBevelDepth
            Else
                dRefSideBevelDepth = dRefSideBevelDepth + dNRRefSideSecondBevelDepth * Tan(dNRRefSideSecondBevelAngle)
            End If
            
            dRefSideBevelDepth = dRefSideBevelDepth + dRootGapCorrection + dNRRefSideBevelDepthCorrection
                        
            'AntiRefSide Bevel Depth
            If Equal(dNRAntiRefSideFirstBevelAngle, PI / 2, dTolerance) Then
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dNRAntiRefSideFirstBevelDepth
            Else
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dNRAntiRefSideFirstBevelDepth * Tan(dNRAntiRefSideFirstBevelAngle)
            End If
            
            If Equal(dNRAntiRefSideSecondBevelAngle, PI / 2, dTolerance) Then
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dNRRefSideSecondBevelDepth
            Else
                dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dNRRefSideSecondBevelDepth * Tan(dNRAntiRefSideSecondBevelAngle)
            End If
            
            dAntiRefSideBevelDepth = dAntiRefSideBevelDepth + dRootGapCorrection + dNRAntiRefSideBevelDepthCorrection
        End If
    End If
    
End Sub

' ** End CM **


' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation
 
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractParameterRule pPR
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


