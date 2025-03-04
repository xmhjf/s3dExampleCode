Attribute VB_Name = "ShrinkageHelper"
Option Explicit
Const MODULE = "ShrinkageHelper.bas"
Private Const TOLERANCE = 0.01
Public Const ADD_SHRINKAGES = 1
Public Const MULTIPLY_SHRINKAGES = 2
Public Const AVERAGE_SHRINKAGES = 3

Public Function InSameDirection(oVector1 As IJDVector, oVector2 As IJDVector) As Boolean

Dim dDotProduct As Double
dDotProduct = Abs(oVector1.Dot(oVector2))

If dDotProduct > (1 - TOLERANCE) And dDotProduct < (1 + TOLERANCE) Then
    InSameDirection = True
End If

End Function

Private Function CalculateFactorInVecDirection(oVector As IJDVector, oPrimVector As IJDVector, dPrimFactor As Double, oSecVector As IJDVector, dSecFactor As Double, oTertVector As IJDVector, dTertFactor As Double) As Double
Const METHOD = "CalculateFactorInRangeBoxVecDirection"
On Error GoTo ErrorHandler
   
    Dim dProductPrim As Double, dProductSec As Double, dProductTert As Double
    
    ' Calculate Cos(angle between Global vec(X/Y/Z) and Range Box Vec)
    If Not oPrimVector Is Nothing Then
        dProductPrim = Abs(oVector.Dot(oPrimVector))
    End If
    
    If Not oSecVector Is Nothing Then
        dProductSec = Abs(oVector.Dot(oSecVector))
    End If
    
    If Not oTertVector Is Nothing Then
        dProductTert = Abs(oVector.Dot(oTertVector))
    End If
    
    'If A is the angle between profile landing curve vector and assembly primary direction vector = dProductPrim in this case
    'If B is the angle between profile landing curve vector and assembly secondary direction vector = dProductSec in this case
    'If C is the angle between profile landing curve vector and assembly tertiary direction vector = dProductTert in this case
    'and Sx is the assembly primary factor, Sy is the assembly secondary factor and Sz is assembly tertiary factor, then
    'the factor in the required direction would be equal to Sx * cosine(A) * cosine(A) + Sy * cosine(B) * cosine(B) + Sy * cosine(C) * cosine(C)
    CalculateFactorInVecDirection = (dPrimFactor * dProductPrim * dProductPrim) + (dSecFactor * dProductSec * dProductSec) + (dTertFactor * dProductTert * dProductTert)
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetVectorFromAxisOrPort(oDirection As Object) As IJDVector
Const METHOD = "GetVectorFromAxisOrPort"
On Error GoTo ErrorHandler
   
    If oDirection Is Nothing Then Exit Function
    
    If TypeOf oDirection Is IJPort Then
        Dim oPort As IJPort
        Set oPort = oDirection
        
        Dim oStartPos As IJDPosition
        Dim oEndPos As IJDPosition
        Dim oWireBody As IJWireBody
        
        Dim oEntityHelper As MfgEntityHelper
        Set oEntityHelper = New MfgEntityHelper
        
        If TypeOf oPort.Connectable Is IJPlatePart Then
            Dim oEdgePort As IJPort
            Set oEdgePort = oEntityHelper.GetEdgePortGivenFacePort(oPort, CTX_BASE)
                                    
            Set oWireBody = oEdgePort.Geometry
            oWireBody.GetEndPoints oStartPos, oEndPos
        ElseIf TypeOf oPort.Connectable Is IJProfilePart Then
            Dim oLandingCurve               As IJWireBody
            Dim oThicknessDir               As IJDVector
            Dim bThicknessCentered          As Boolean
            
            Dim oSDProfilePart As New StructDetailObjects.ProfilePart
            Set oSDProfilePart.object = oPort.Connectable
            
            On Error Resume Next
            oSDProfilePart.LandingCurve oLandingCurve, oThicknessDir, bThicknessCentered
            On Error GoTo ErrorHandler
            
            If oLandingCurve Is Nothing Then Exit Function
        
            oLandingCurve.GetEndPoints oStartPos, oEndPos
            Set oWireBody = oLandingCurve
        Else
            Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
            Dim oAcisCurve As IJCurve
            Dim dStartX As Double, dStartY As Double, dStartZ As Double
            Dim dEndX As Double, dEndY As Double, dEndZ As Double
            
            Set oMemberPartPrismatic = oPort.Connectable
            Set oAcisCurve = oMemberPartPrismatic.Axis
            oAcisCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
            
            Set oStartPos = New DPosition
            oStartPos.Set dStartX, dStartY, dStartZ
            
            Set oEndPos = New DPosition
            oEndPos.Set dEndX, dEndY, dEndZ
        End If
        
        Dim oDirectionVec As IJDVector
        Set oDirectionVec = New DVector
    
        'set primary direction vector
        oDirectionVec.Set oEndPos.x - oStartPos.x, oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z
        oDirectionVec.length = 1
        
        Set GetVectorFromAxisOrPort = oDirectionVec
    ElseIf TypeOf oDirection Is IJDVector Then
        Set GetVectorFromAxisOrPort = oDirection
    Else
        Dim oXVector As IJDVector
        Dim oYVector As IJDVector
        Dim oZVector As IJDVector

        Set oXVector = New DVector
        Set oYVector = New DVector
        Set oZVector = New DVector

        oXVector.Set 1, 0, 0
        oYVector.Set 0, 1, 0
        oZVector.Set 0, 0, 1

        oXVector.length = 1
        oYVector.length = 1
        oZVector.length = 1

        If InSameDirection(oXVector, oDirection.UnitVector) = True Then
            Set GetVectorFromAxisOrPort = oXVector
        ElseIf InSameDirection(oYVector, oDirection.UnitVector) = True Then
            Set GetVectorFromAxisOrPort = oYVector
        ElseIf InSameDirection(oZVector, oDirection.UnitVector) = True Then
            Set GetVectorFromAxisOrPort = oZVector
        End If
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetGlobalCoordinateSystem(ByVal pDispObj As Object) As Object
    Const METHOD = "GetGlobalCoordinateSystem"
    On Error GoTo ErrorHandler
    
    Dim oSysUtil As IJDMfgFrameSysUtil
    Set oSysUtil = New CMfgCoordSys
    
    Dim oParentFS As IHFrameSystem
    Set oParentFS = oSysUtil.SystemDefaultCS(pDispObj)
    
    If Not oParentFS Is Nothing Then
        Set GetGlobalCoordinateSystem = oParentFS
    Else
        Dim oTempFSColl As IJElements
        Set oTempFSColl = oSysUtil.FrameSystemsInRange(pDispObj)
        
        If oTempFSColl.Count = 0 Then
            Set oTempFSColl = oSysUtil.FrameSystemsFromCatalogRule("CScalingShr_FrameSystem", pDispObj)
        End If
        
        Dim nIndex As Long
        Dim FS As IHFrameSystem
    
        For nIndex = 1 To oTempFSColl.Count
            Set FS = oTempFSColl(nIndex)
            If Not FS Is Nothing Then
                If "CS_0" = FS.Name Then
                    Set GetGlobalCoordinateSystem = FS
                    Exit For
                End If
            End If
        Next nIndex
    
        If GetGlobalCoordinateSystem Is Nothing Then
            If oTempFSColl.Count > 0 Then
                Set GetGlobalCoordinateSystem = oTempFSColl(1)
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AccumulateShrinkageParameters(ByVal oShrinkageColl As IJElements, _
                ByVal oPrimVector As IJDVector, _
                ByVal oSecVector As IJDVector, _
                ByVal eAccumulationMethod As Integer, _
                PrimaryFactor As Double, _
                SecondaryFactor As Double)
    Const METHOD = "IJDShrinkageRule_GetShrinkageParameters"
    On Error GoTo ErrorHandler
    
    'compute factors in each direction
    Dim iShrCount As Long
    iShrCount = oShrinkageColl.Count
    Dim dPrimFactor As Double, dSecFactor As Double
    
    Dim dSumOfPriShrVal As Double, dSumOfSecShrVal As Double
    Dim iNumOfPriShrVal As Long, iNumOfSecShrVal As Long
    
    dPrimFactor = 1#
    dSecFactor = 1#
    
    dSumOfPriShrVal = 0#
    dSumOfSecShrVal = 0#
    iNumOfPriShrVal = 0
    iNumOfSecShrVal = 0
    
    Dim i As Long
    
    If iShrCount > 0 Then
        For i = 1 To iShrCount
            'get the assemly shrinkage
            Dim oShrinkage As IJScalingShr
            Set oShrinkage = oShrinkageColl.Item(i)
    
            Dim oShrParent As Object
            Dim dThisPrimaryFactor As Double, dThisSecondaryFactor As Double, dThisTertiaryFactor As Double
            Dim oThisPrimaryObj As Object, oThisSecondaryObj As Object, oThisTertiaryObj As Object
            
            'Get primary and secondary directions
            oShrinkage.GetInputs oShrParent, oThisPrimaryObj, oThisSecondaryObj
            
            Dim oThisPrimaryDir As IJDVector, oThisSecondaryDir As IJDVector, oThisTertiaryDir As IJDVector
            
            'Get vectors in primary and secondary directions
            Set oThisPrimaryDir = GetVectorFromAxisOrPort(oThisPrimaryObj)
            Set oThisSecondaryDir = GetVectorFromAxisOrPort(oThisSecondaryObj)
            
            'get tertiary direction vector
            oShrinkage.GetTertiaryScalingShrinkage dThisTertiaryFactor, oThisTertiaryDir, oThisTertiaryObj
    
            'Get factors in all three directions
            dThisPrimaryFactor = oShrinkage.PrimaryFactor - 1#
            dThisSecondaryFactor = oShrinkage.SecondaryFactor - 1#
            dThisTertiaryFactor = oShrinkage.TertiaryFactor - 1#
                                        
            'Use the different factors in three directions to get the factor in the required directions
            Dim dTempFactor  As Double
            dTempFactor = 0
            If oPrimVector.length > 0.0001 Then ' Make sure this is not a NULL vector
                dTempFactor = CalculateFactorInVecDirection(oPrimVector, oThisPrimaryDir, _
                                                            dThisPrimaryFactor, oThisSecondaryDir, _
                                                            dThisSecondaryFactor, oThisTertiaryDir, _
                                                            dThisTertiaryFactor)
            Else
                dTempFactor = dThisPrimaryFactor
            End If
            
            If eAccumulationMethod = ADD_SHRINKAGES Then
                dPrimFactor = dPrimFactor + dTempFactor
            ElseIf eAccumulationMethod = MULTIPLY_SHRINKAGES Then
                dPrimFactor = dPrimFactor * (1 + dTempFactor)
            Else
                dSumOfPriShrVal = dSumOfPriShrVal + dTempFactor
                iNumOfPriShrVal = iNumOfPriShrVal + 1
            End If
            
            dTempFactor = 0
                        
            If Not oSecVector Is Nothing Then
                If oSecVector.length > 0.0001 Then ' Make sure this is not a NULL vector
                    dTempFactor = CalculateFactorInVecDirection(oSecVector, oThisPrimaryDir, _
                                                                dThisPrimaryFactor, oThisSecondaryDir, _
                                                                dThisSecondaryFactor, oThisTertiaryDir, _
                                                                dThisTertiaryFactor)
                Else
                    dSecFactor = 0#
                End If
                
                If eAccumulationMethod = ADD_SHRINKAGES Then
                    dSecFactor = dSecFactor + dTempFactor
                ElseIf eAccumulationMethod = MULTIPLY_SHRINKAGES Then
                    dSecFactor = dSecFactor * (1 + dTempFactor)
                Else
                    dSumOfSecShrVal = dSumOfSecShrVal + dTempFactor
                    iNumOfSecShrVal = iNumOfSecShrVal + 1
                End If
            End If
        Next
    End If

    If eAccumulationMethod = AVERAGE_SHRINKAGES Then
        If (iNumOfPriShrVal > 0) Then
            PrimaryFactor = dPrimFactor + (dSumOfPriShrVal / iNumOfPriShrVal)
        End If
        If (iNumOfSecShrVal > 0) Then
            SecondaryFactor = dSecFactor + (dSumOfSecShrVal / iNumOfSecShrVal)
        End If
    Else
        PrimaryFactor = dPrimFactor
        SecondaryFactor = dSecFactor
    End If

    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7005, , "RULES")
End Sub

Public Function GetStringFromStiffDirectionEnum(eStiffDirection As StructProfileType) As String
Const sMETHOD = "GetStringFromStiffDirectionEnum"
On Error GoTo ErrorHandler

    If eStiffDirection = sptLongitudinal Then
        GetStringFromStiffDirectionEnum = "X-Direction"
    ElseIf eStiffDirection = sptTransversal Then
        GetStringFromStiffDirectionEnum = "Y-Direction"
    ElseIf eStiffDirection = sptVertical Then
        GetStringFromStiffDirectionEnum = "Z-Direction"
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
