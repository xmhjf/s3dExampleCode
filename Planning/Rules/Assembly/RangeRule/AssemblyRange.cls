VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation. All rights reserved.
'
'  Project     : Planning\Data\Rules\AssemblyRange\
'  File        : RangeInfo.cls
'
'  Description :
'
'  History     :
'   6th Nov 2012      Siva     Initial creation
'**************************************************************************************

Option Explicit
Private Const MODULE = "RangeInfo"
Private Const IID_IJScalingShr = "{DE77050C-3300-11D5-BA1A-0090276F4279}"
Private Const IID_IJAssyMarginParent = "{A998DAF5-7AB0-4964-8513-188788F40677}"
Implements IJPlnAssemblyRangeRule

Private Sub IJPlnAssemblyRangeRule_GetCustomRange(ByVal lRangeTypeID As Long, ByVal oAssembly As Object, dLength As Double, dWidth As Double, dHeight As Double)
    Const METHOD = "IJPlnAssemblyRangeRule_GetCustomRange"
    On Error GoTo ErrorHandler
    
    ' Algorithm:
    ' Call middle APIs to calculate the range box of Assembly/Block based on Range type inputs
    ' Apply Assembly shrinkage and Assembly Margin on range box
    ' Return Length, Width and Height of range box
            
    ' Check the inputs
    If lRangeTypeID < 10000 Or oAssembly Is Nothing Then
        Err.Description = "Invaild inputs to the rule"
        Exit Sub
    End If
    
    ' Get the Range Type IDs from the Assembly Range table in catalog
    Dim lCount As Long, iIndex  As Long, lRangeInput1 As Long, lRangeInput2 As Long, lRangeInput3 As Long, lRangeInput4 As Long, lRangeInput5 As Long
    RunQueryOnAssemblyRangeLookUp lRangeTypeID, lRangeInput1, lRangeInput2, lRangeInput3, lRangeInput4, lRangeInput5
        
    Dim oBoxPoints      As IJElements
    Dim oPoints(1 To 4) As IJDPosition
    Dim oRootPoint      As IJDPosition
    Dim oXAxis          As IJDVector
    Dim oYAxis          As IJDVector
    Dim oZAxis          As IJDVector
    Dim bReduceRange    As Boolean
    Dim oRange          As IJRangeAlias
    
    If lRangeInput1 = 2 Then ' Nominal range
        Set oRootPoint = New DPosition
        Set oXAxis = New DVector
        Set oYAxis = New DVector
        Set oZAxis = New DVector
        
        If TypeOf oAssembly Is IJRangeAlias Then
            
            Set oRange = oAssembly
            
            Dim oGBox           As GBox
            oGBox = oRange.GetRange
            
            If Not TypeName(oAssembly) = "IJBlock" Then
                ' For block, internal geometries are surface body(not GTypes)
                ' Need to remove 2mm to the range as math routine adds 1mm on each side
                ' This happens only when range is fetched from the part directly(using CoreSpatialIndex table)
                bReduceRange = True
            End If

            oXAxis.Set oGBox.m_high.x - oGBox.m_low.x, 0, 0
            oYAxis.Set 0, oGBox.m_high.y - oGBox.m_low.y, 0
            oZAxis.Set 0, 0, oGBox.m_high.z - oGBox.m_low.z
        Else
            Exit Sub
        End If
            
    Else ' Actual range
        
        Dim oPlnIntHelper As IJDPlnIntHelper
        Set oPlnIntHelper = New CPlnIntHelper
        
        ' Get all the children from the input Assembly/Block
        Dim pChildElems     As IJElements
        Set pChildElems = oPlnIntHelper.GetStoredProcAssemblyChildren(oAssembly, "IJAssemblyChild", False, Nothing, False)
        
        If Not pChildElems Is Nothing Then
            lCount = pChildElems.Count
        End If
        
        If lCount < 1 Then
            Exit Sub
        Else
            Dim oSkippedParts As IJElements
            Set oSkippedParts = New JObjectCollection
            
            For iIndex = 1 To lCount
                Dim oAssemblyChild As IJAssemblyChild
                Set oAssemblyChild = pChildElems.Item(iIndex)
                
                Dim strTypeName As String
                
                If Not oAssemblyChild Is Nothing Then
                    Dim oParentunk   As Object
                    Set oParentunk = oAssemblyChild.Parent
                    
                    If Not oParentunk Is Nothing Then
                        strTypeName = TypeName(oParentunk)
                        If strTypeName = "IJPlnUnprocessedParts" Or strTypeName = "IJPlnFailedParts" Then
                            'Ignore any part that is under an UnAssigned parts or Failed Parts folder
                            oSkippedParts.Add oAssemblyChild
                        End If
                    End If
                End If
                
                strTypeName = TypeName(oAssemblyChild)
                
                If strTypeName = "IJPlanningAssembly" Or _
                   strTypeName = "IJAssemblyBlock" Or _
                   strTypeName = "IJBlock" Or _
                   strTypeName = "IJAssyMarginParent" Or _
                   strTypeName = "IJPinJig" Then
                    ' Skip Assembly, AssemblyBlock, Block, Assembly Margin and Assembly PinJig objects
                    oSkippedParts.Add oAssemblyChild
                End If
            Next
            
            ' Remove the unassigned and failed parts from the collection
            pChildElems.RemoveElements oSkippedParts
            
            lCount = pChildElems.Count
            If lCount < 1 Then
                Exit Sub
            End If
        End If

        If lRangeInput2 = 1 Then ' Global Orientation
            
            Dim gAssyRangeBox As GBox
            Dim bFirstPart As Boolean
            bFirstPart = True
            
            ' Get the RangeBox of the parts
            For iIndex = 1 To lCount
                'Get all the graphic entities
                Set oRange = pChildElems.Item(iIndex)

                If bFirstPart = True Then
                    gAssyRangeBox = oRange.GetRange
                    bFirstPart = False
                Else
                    Dim gTempBox As GBox
                    gTempBox = oRange.GetRange
                    
                    If gAssyRangeBox.m_high.x < gTempBox.m_high.x Then
                        gAssyRangeBox.m_high.x = gTempBox.m_high.x
                    End If
                    If gAssyRangeBox.m_high.y < gTempBox.m_high.y Then
                        gAssyRangeBox.m_high.y = gTempBox.m_high.y
                    End If
                    If gAssyRangeBox.m_high.z < gTempBox.m_high.z Then
                        gAssyRangeBox.m_high.z = gTempBox.m_high.z
                    End If
                    If gAssyRangeBox.m_low.x > gTempBox.m_low.x Then
                        gAssyRangeBox.m_low.x = gTempBox.m_low.x
                    End If
                    If gAssyRangeBox.m_low.y > gTempBox.m_low.y Then
                        gAssyRangeBox.m_low.y = gTempBox.m_low.y
                    End If
                    If gAssyRangeBox.m_low.z > gTempBox.m_low.z Then
                        gAssyRangeBox.m_low.z = gTempBox.m_low.z
                    End If
                End If
                
                Set oRange = Nothing
            Next
            
            Set oRootPoint = New DPosition
            Set oXAxis = New DVector
            Set oYAxis = New DVector
            Set oZAxis = New DVector
            
            oXAxis.Set gAssyRangeBox.m_high.x - gAssyRangeBox.m_low.x, 0, 0
            oYAxis.Set 0, gAssyRangeBox.m_high.y - gAssyRangeBox.m_low.y, 0
            oZAxis.Set 0, 0, gAssyRangeBox.m_high.z - gAssyRangeBox.m_low.z
            
        ElseIf lRangeInput2 = 2 Then ' Assembly Orientation
        
            Dim oAssemblyLCS    As IJLocalCoordinateSystem
            Set oAssemblyLCS = oAssembly
            
            Dim oVecElements    As IJElements
            Set oVecElements = New JObjectCollection
            
            oVecElements.Add oAssemblyLCS.XAxis
            oVecElements.Add oAssemblyLCS.YAxis
            oVecElements.Add oAssemblyLCS.ZAxis
            
            Set oBoxPoints = oPlnIntHelper.GetGeometryMinBoxByVectors(pChildElems, oVecElements)

        ElseIf lRangeInput2 = 3 Then ' Min Bounding Box
        
            Set oBoxPoints = oPlnIntHelper.GetGeometryMinBoxByVectors(pChildElems)
            
        Else
            ' Do nothing
        End If
        
    End If
    
    ' Compute Length, Width and Height of Range box
    If Not oBoxPoints Is Nothing Then
        dLength = oPoints(1).DistPt(oPoints(2))
        dWidth = oPoints(2).DistPt(oPoints(3))
        dHeight = oPoints(1).DistPt(oPoints(4))
    ElseIf Not oRootPoint Is Nothing Then
        dLength = oXAxis.Length
        dWidth = oYAxis.Length
        dHeight = oZAxis.Length
        
        If bReduceRange = True Then   ' for Nominal and Assembly Orientation
            ' Need to remove 2mm to the range as math routine adds 1mm on each side
            ' This happens only when range is fetched from the part directly(using CoreSpatialIndex table)
            dLength = dLength - 0.002
            dWidth = dWidth - 0.002
            dHeight = dHeight - 0.002
        End If
    End If
    
    If lRangeInput3 = 1 Or lRangeInput4 = 1 Then    ' Apply Assembly Shrinkage, Assembly Margin
        
        Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
        Dim oMfgObjColl As IJElements
        
        Dim oRangeBoxXVec As IJDVector
        Dim oRangeBoxYVec As IJDVector
        Dim oRangeBoxZVec As IJDVector
        
        If Not oBoxPoints Is Nothing Then
            Set oRangeBoxXVec = oPoints(1).Subtract(oPoints(2))
            Set oRangeBoxYVec = oPoints(2).Subtract(oPoints(3))
            Set oRangeBoxZVec = oPoints(1).Subtract(oPoints(4))
        ElseIf Not oRootPoint Is Nothing Then
            Set oRangeBoxXVec = oXAxis
            Set oRangeBoxYVec = oYAxis
            Set oRangeBoxZVec = oZAxis
        End If
        
        oRangeBoxXVec.Length = 1
        oRangeBoxYVec.Length = 1
        oRangeBoxZVec.Length = 1
            
        If lRangeInput3 = 1 Then ' Apply Assembly Shrinkage
            Set oMfgObjColl = oStructMfgGlobals.GetMfgPart(oAssembly, IID_IJScalingShr)
            
            If Not oMfgObjColl Is Nothing Then
                If oMfgObjColl.Count > 0 Then
                    Dim oMfgAssemblyShrinkage   As IJScalingShr
                    Set oMfgAssemblyShrinkage = oMfgObjColl.Item(1)
                    
                    Dim dPrimaryFactor  As Double
                    Dim dSecondaryFactor  As Double
                    Dim dTertiaryFactor As Double
            
                    ' Get the Assembly Shrinkage factors
                    dPrimaryFactor = oMfgAssemblyShrinkage.PrimaryFactor
                    dSecondaryFactor = oMfgAssemblyShrinkage.SecondaryFactor
                    dTertiaryFactor = oMfgAssemblyShrinkage.TertiaryFactor
                    
                    ' Calculate the shrinkage factor to be applied in box orintation
                    Dim dXDirShrFactor  As Double
                    dXDirShrFactor = CalculateFactorInRangeBoxVecDirection(oRangeBoxXVec, dPrimaryFactor, dSecondaryFactor, dTertiaryFactor)
                    
                    Dim dYDirShrFactor  As Double
                    dYDirShrFactor = CalculateFactorInRangeBoxVecDirection(oRangeBoxYVec, dPrimaryFactor, dSecondaryFactor, dTertiaryFactor)
                    
                    Dim dZDirShrFactor  As Double
                    dZDirShrFactor = CalculateFactorInRangeBoxVecDirection(oRangeBoxZVec, dPrimaryFactor, dSecondaryFactor, dTertiaryFactor)
                    
                    ' Calculate new dimensions based on shrinkage factors in box orientation
                    dLength = dLength * (1 + dXDirShrFactor / 1000)
                    dWidth = dWidth * (1 + dYDirShrFactor / 1000)
                    dHeight = dHeight * (1 + dZDirShrFactor / 1000)
                End If
            End If
        End If
        
        If lRangeInput4 = 1 Then ' Apply Assembly Margin
            Set oMfgObjColl = oStructMfgGlobals.GetMfgPart(oAssembly, IID_IJAssyMarginParent)
            
            If Not oMfgObjColl Is Nothing Then
                lCount = oMfgObjColl.Count
                For iIndex = 1 To lCount
                
                    Dim oMfgAssemblyMargin  As IJAssyMarginParent
                    Set oMfgAssemblyMargin = oMfgObjColl.Item(iIndex)
                    
                    Dim oAssyMarginChildElems As IJElements
                    Set oAssyMarginChildElems = oMfgAssemblyMargin.GetAssyMarginChildren
                    
                    Dim dMaxMarginVal   As Double
                    Dim jIndex          As Long
                    Dim lMarginCnt      As Long
                    Dim oRelatedPort    As IJPort
                    
                    If Not oAssyMarginChildElems Is Nothing Then
                        lMarginCnt = oAssyMarginChildElems.Count
                        
                        ' Get the Margin value that is larger and related port
                        For jIndex = 1 To lMarginCnt
                            Dim oConstantMargin As IJConstMargin
                            Set oConstantMargin = oAssyMarginChildElems.Item(jIndex)
                            
                            Dim dMarginValue As Double
                            dMarginValue = oConstantMargin.Value
                            
                            If dMaxMarginVal < dMarginValue Then
                                Dim oMfgDefinition As IJMfgDefinition
                                Set oMfgDefinition = oConstantMargin
                                Set oRelatedPort = oMfgDefinition.GetPort
                                
                                dMaxMarginVal = dMarginValue
                            End If
                        Next
                        
                        ' Get the port normal to apply margin on the range box in the port normal direction
                        Dim oSurfaceBody    As IJSurfaceBody
                        Set oSurfaceBody = oRelatedPort.Geometry
                        
                        Dim oCOGPos As IJDPosition
                        oSurfaceBody.GetCenterOfGravity oCOGPos
                        
                        Dim oPortNormalVec As IJDVector
                        oSurfaceBody.GetNormalFromPosition oCOGPos, oPortNormalVec
                        
                        oPortNormalVec.Length = 1
                        
                        ' Compute Margin value in each direction of range box
                        Dim dRangeXDirMargin    As Double
                        dRangeXDirMargin = CalculateMarginValInRangeBoxVecDirection(oRangeBoxXVec, oPortNormalVec, dMaxMarginVal)
                        
                        Dim dRangeYDirMargin  As Double
                        dRangeYDirMargin = CalculateMarginValInRangeBoxVecDirection(oRangeBoxYVec, oPortNormalVec, dMaxMarginVal)
                        
                        Dim dRangeZDirMargin  As Double
                        dRangeZDirMargin = CalculateMarginValInRangeBoxVecDirection(oRangeBoxZVec, oPortNormalVec, dMaxMarginVal)
                        
                        ' Calculate new dimensions based on Margin values in box orientation
                        dLength = dLength + dRangeXDirMargin
                        dWidth = dWidth + dRangeYDirMargin
                        dHeight = dHeight + dRangeZDirMargin
                    End If
                Next
            End If
        End If
    End If
    
Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub RunQueryOnAssemblyRangeLookUp(lRangeTypeID As Long, lRangeInput1 As Long, lRangeInput2 As Long, lRangeInput3 As Long, lRangeInput4 As Long, lRangeInput5 As Long)
    Const METHOD = "RunQueryOnBUProfileAssemblyLookUp"
    On Error GoTo ErrorHandler
    
    Dim strQuery As String
    Dim oQueryOutputValues() As Variant
    
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
    
    ' Construct the query to get the Range Input values
    strQuery = "SELECT  RangeInput1 , RangeInput2, RangeInput3, RangeInput4, RangeInput5 FROM JPlnAssemblyRangeControl WHERE RangeTypeID = " + CStr(lRangeTypeID) + ""
        
    On Error Resume Next
    oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
    On Error GoTo ErrorHandler
    
    If (UBound(oQueryOutputValues) > 0) Then
        lRangeInput1 = oQueryOutputValues(0)
        lRangeInput2 = oQueryOutputValues(1)
        lRangeInput3 = oQueryOutputValues(2)
        lRangeInput4 = oQueryOutputValues(3)
        lRangeInput5 = oQueryOutputValues(4)
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function CalculateFactorInRangeBoxVecDirection(oRangeBoxVec As IJDVector, dXFactor As Double, dYFactor As Double, dZFactor As Double) As Double
Const METHOD = "CalculateFactorInRangeBoxVecDirection"
On Error GoTo ErrorHandler
   
    Dim dProductX As Double, dProductY As Double, dProductZ As Double
    Dim oXVector As IJDVector, oYVector As IJDVector, oZVector As IJDVector
    
    Set oXVector = New DVector
    Set oYVector = New DVector
    Set oZVector = New DVector
    
    oXVector.Set 1, 0, 0
    oYVector.Set 0, 1, 0
    oZVector.Set 0, 0, 1
    
    ' Calculate Cos(angle between Global vec(X/Y/Z) and Range Box Vec)
    dProductX = Abs(oRangeBoxVec.Dot(oXVector))
    dProductY = Abs(oRangeBoxVec.Dot(oYVector))
    dProductZ = Abs(oRangeBoxVec.Dot(oZVector))
    
    ' Get the component of the length in three directions...
    ' Use the above cos of angle and multiple it with total lengths to get the component in each of the directions (dCompTotalLength)
    
    'dxCompTotalLength = x component of length plus added shrinkage in x direction
    Dim dXCompTotalLength As Double
    Dim dYCompTotalLength As Double
    Dim dZCompTotalLength As Double
    
    dXCompTotalLength = dProductX + (dProductX * dXFactor)
    dYCompTotalLength = dProductY + (dProductY * dYFactor)
    dZCompTotalLength = dProductZ + (dProductZ * dZFactor)
    
    ' Total primary port length after adding shrinkage = sqrt(sum of squares of components)
    Dim dPortNewLength As Double
    dPortNewLength = Sqr((dXCompTotalLength * dXCompTotalLength) + (dYCompTotalLength * dYCompTotalLength) + (dZCompTotalLength * dZCompTotalLength))
    
    ' Effective factor = (New Length - Old Length/Old Length) = (New Length/Old Length) - 1
    CalculateFactorInRangeBoxVecDirection = dPortNewLength - 1
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function CalculateMarginValInRangeBoxVecDirection(oRangeBoxVec As IJDVector, oPortNormalVec As IJDVector, oMaxMarginVal As Double) As Double
Const METHOD = "CalculateMarginValInRangeBoxVecDirection"
On Error GoTo ErrorHandler
    
    ' dproduct is value of cos(angle between primary dir and Port normal vector)
    Dim dProduct As Double
    dProduct = Abs(oRangeBoxVec.Dot(oPortNormalVec))
    
    ' Return component of length plus added Margin in Range box vec direction
    CalculateMarginValInRangeBoxVecDirection = dProduct * oMaxMarginVal
    
Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
