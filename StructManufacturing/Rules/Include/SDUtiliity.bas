Attribute VB_Name = "SDUtiliity"
Option Explicit

Const MODULE = "CustomReports.SDUtility::"

'Set up some constant Tolerance Values
Private Const C_TOL001 = 0.001
Private Const C_TOL0001 = 0.0001
Private Const C_TOL00001 = 0.00001
Private Const C_TOL000001 = 0.000001

'The Enum type of seam orientation
Public Enum eSeamDirectionType
   UnknownType = -1
   Upper = 1
   Lower = 2
   Aft = 3
   Fore = 4
End Enum

Public Sub GetOrderPortsList(oPortList As IJElements, _
                             oOrderIndex() As Long)
Const METHOD = "GetOrderPortsList"
On Error GoTo ErrorHandler


    
    Dim nPorts As Long
    Dim iIndex1 As Long
    Dim iIndex2 As Long
    Dim iIndex3 As Long
    Dim iMinIndex As Long
    Dim iMinOrder As Long
    
    Dim dValue1 As Double
    Dim dValue2 As Double
    Dim dValue3 As Double
    
    Dim bContinue As Boolean
    
    Dim aOrderIndex() As Long
    Dim aOrderFlags() As Long
    Dim aPortXvalue() As Double
    Dim aPortYvalue() As Double
    Dim aPortZvalue() As Double
    Dim aPortMinXvalue() As Double
    Dim aPortMinYvalue() As Double
    Dim aPortMinZvalue() As Double
    Dim aPortMaxXvalue() As Double
    Dim aPortMaxYvalue() As Double
    Dim aPortMaxZvalue() As Double
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    Dim oLowPos As IJDPosition
    Dim oHighPos As IJDPosition
    Dim oIntersectObject1 As IUnknown
    
    Dim oNameUtil As SDNameRulesUtilHelper
    Dim oIntersect As IIntersect
    Dim oTopologyIntersect As IJDTopologyIntersect
    
    On Error Resume Next
    
    nPorts = oPortList.Count
    ReDim oOrderIndex(nPorts)
    
    Dim sText As String
    sText = "nPorts = " & Str(nPorts)
    
    Set oNameUtil = New GSCADSDNameRulesUtil.SDNameRulesUtilHelper
    Set oIntersect = New Intersect
    Set oTopologyIntersect = New DGeomOpsIntersect
    
    ' Initialize data arrays for each Port
    ' For each Edge Port,
    '   retieve it's min/max Range Points
    '   calculate middle of Port's Range Box
    ReDim aOrderIndex(nPorts)
    ReDim aOrderFlags(nPorts)
    ReDim aPortXvalue(nPorts)
    ReDim aPortYvalue(nPorts)
    ReDim aPortZvalue(nPorts)
    ReDim aPortMinXvalue(nPorts)
    ReDim aPortMinYvalue(nPorts)
    ReDim aPortMinZvalue(nPorts)
    ReDim aPortMaxXvalue(nPorts)
    ReDim aPortMaxYvalue(nPorts)
    ReDim aPortMaxZvalue(nPorts)
    
    For iIndex1 = 1 To nPorts
        aOrderFlags(iIndex1) = 0
        aOrderIndex(iIndex1) = 0
        oNameUtil.GetRangeCorners oPortList.Item(iIndex1), oLowPos, oHighPos
        aPortMinXvalue(iIndex1) = oLowPos.x
        aPortMinYvalue(iIndex1) = oLowPos.y
        aPortMinZvalue(iIndex1) = oLowPos.z
        
        aPortMaxXvalue(iIndex1) = oHighPos.x
        aPortMaxYvalue(iIndex1) = oHighPos.y
        aPortMaxZvalue(iIndex1) = oHighPos.z
        
        aPortXvalue(iIndex1) = (oLowPos.x + oHighPos.x) / 2#
        aPortYvalue(iIndex1) = (oLowPos.y + oHighPos.y) / 2#
        aPortZvalue(iIndex1) = (oLowPos.z + oHighPos.z) / 2#
        
        Set oLowPos = Nothing
        Set oHighPos = Nothing
    Next iIndex1
    
    ' Starting with the first Port Edge
    ' Find an Edge that intersects the First Port Edge
    iIndex1 = 1
    bContinue = True
    While bContinue
    
        iIndex3 = 0
        bContinue = False
        aOrderFlags(iIndex1) = -1
        Set oPort1 = oPortList.Item(iIndex1)
        For iIndex2 = 1 To nPorts
            ' Skip the current Edge Port if it has already been Ordered
            If aOrderFlags(iIndex2) = 0 Then
                iIndex3 = iIndex2
                Set oPort2 = oPortList.Item(iIndex2)
                oIntersect.GetCommonGeometry oPort1.Geometry, _
                                              oPort2.Geometry, _
                                              oIntersectObject1, _
                                              False
                If oIntersectObject1 Is Nothing Then
                    ' Expect the GetCommonGeometry to return valid geometry
                    ' from adjacent Ports but if not,
                    ' see if the Ports intersect
                    oTopologyIntersect.PlaceIntersectionObject Null, _
                                                            oPort1.Geometry, _
                                                            oPort2.Geometry, _
                                                            Null, _
                                                            oIntersectObject1
                End If
            
                Set oPort2 = Nothing
                ' if the current Port Edge (Index2) intersects the Base Port Edge (Index1)
                '   set the Ordered flag for the Base Port Edge to
                '   point to the current Port Edge
                '   Use current Port Edge as the next Base Port Edge
                If Not oIntersectObject1 Is Nothing Then
                    Set oIntersectObject1 = Nothing
                    aOrderFlags(iIndex1) = iIndex2
                    iIndex1 = iIndex2
                    bContinue = True
                    Exit For
                End If
            End If
        Next iIndex2
        
        Set oPort1 = Nothing
        
    Wend
            
    ' Verify That all Port Edges have been Ordered
    bContinue = True
    For iIndex1 = 1 To nPorts
        If aOrderFlags(iIndex1) = 0 Then
            ' not all Port Edges have been ordered
            ' default to input order
            bContinue = False
        
            For iIndex2 = 1 To nPorts
                oOrderIndex(iIndex2) = iIndex2
            Next iIndex2
        
            Exit Sub
        End If
    Next iIndex1
    
    ' Determine the Port Edge at Minimum (X,|Y|,Z) Point
    iIndex1 = 1
    iIndex2 = 1
    iMinIndex = 1
    iMinOrder = 1
    aOrderIndex(iIndex2) = iIndex1
    Set oLowPos = New automath.DPosition
    Set oHighPos = New automath.DPosition
        
    bContinue = True
    While bContinue
        bContinue = False
        iIndex3 = aOrderFlags(iIndex1)
        If iIndex3 > 0 Then
            iIndex2 = iIndex2 + 1
            aOrderIndex(iIndex2) = iIndex3
            oLowPos.Set aPortMinXvalue(iMinOrder), _
                        aPortMinYvalue(iMinOrder), aPortMinZvalue(iMinOrder)
            oHighPos.Set aPortMinXvalue(iIndex3), _
                         aPortMinYvalue(iIndex3), aPortMinZvalue(iIndex3)
            dValue1 = oLowPos.DistPt(oHighPos)
            
            ' Check if current Edge
            '   has minimum X value / has minimum abs(Y) value / has minimum Z value
            If dValue1 > C_TOL001 Then
                If IsMinimumPoint(oHighPos, oLowPos) Then
                    iMinIndex = iIndex2
                    iMinOrder = iIndex3
                End If
            Else
                ' The current Edges are at the same minimum Corner
                ' want the Edge that has least change in Y direction
                dValue1 = aPortMaxYvalue(iIndex3) - aPortMinYvalue(iIndex3)
                dValue2 = aPortMaxYvalue(iMinOrder) - aPortMinYvalue(iMinOrder)
                dValue3 = Abs(dValue1) - Abs(dValue2)
                If dValue3 >= C_TOL001 And dValue1 < dValue2 Then
                    iMinIndex = iIndex2
                    iMinOrder = iIndex3
                Else
                    ' The current Edges are at the same minimum Corner
                    ' The current Edges have identical change in Y direction
                    ' want the Edge at minimum Z middle value
                    dValue1 = Abs(aPortZvalue(iIndex3)) - Abs(aPortZvalue(iMinOrder))
                    If dValue1 >= C_TOL001 And _
                        aPortZvalue(iIndex3) < aPortZvalue(iMinOrder) Then
                        iMinIndex = iIndex2
                        iMinOrder = iIndex3
                    Else
                        dValue1 = Abs(aPortXvalue(iIndex3)) - Abs(aPortXvalue(iMinOrder))
                        If dValue1 >= C_TOL001 And _
                            aPortXvalue(iIndex3) < aPortXvalue(iMinOrder) Then
                            iMinIndex = iIndex2
                            iMinOrder = iIndex3
                        End If
                    End If
                End If
                
            End If
            
            bContinue = True
            iIndex1 = iIndex3
        End If
    Wend
        
    ' Have determined Edge Port with Minimum X,|Y|,Z values
    ' Fill the OrderIndex array with Port Indexes indicating the Order
    iIndex2 = 0
    If iMinIndex < nPorts Then
        For iIndex1 = iMinIndex To nPorts
            iIndex2 = iIndex2 + 1
            oOrderIndex(iIndex2) = aOrderIndex(iIndex1)
        Next iIndex1
    End If
        
    If iMinIndex > 1 Then
        For iIndex1 = 1 To iMinIndex - 1
            iIndex2 = iIndex2 + 1
            oOrderIndex(iIndex2) = aOrderIndex(iIndex1)
        Next iIndex1
    End If
        
        
    ' Determine if the Port Edges have been ordered in 'anti- Clockwise' (Left Hand)
    ' or ordered in 'Clockwise' (Right hand) order
    ' Return the List in 'Clockwise' (Right Hand) Order
   
    Dim dDot1 As Double
    Dim bReverse As Boolean
    
    Dim oCross2 As IJDVector
    Dim oVector12 As IJDVector
    Dim oVector23 As IJDVector
    Dim oPostion1 As IJDPosition
    Dim oPostion2 As IJDPosition
    Dim oPostion3 As IJDPosition
        
    Set oPostion1 = New automath.DPosition
    Set oPostion2 = New automath.DPosition
    Set oPostion3 = New automath.DPosition
        
    dDot1 = 0#
    bReverse = False
    
    iIndex2 = oOrderIndex(1)
    oPostion1.Set aPortXvalue(iIndex2), aPortYvalue(iIndex2), aPortZvalue(iIndex2)
    
    iIndex2 = oOrderIndex(2)
    oPostion2.Set aPortXvalue(iIndex2), aPortYvalue(iIndex2), aPortZvalue(iIndex2)
    
    For iIndex1 = 3 To nPorts + 2
        If iIndex1 > nPorts Then
            iIndex2 = oOrderIndex(iIndex1 - nPorts)
        Else
            iIndex2 = oOrderIndex(iIndex1)
        End If
        
        oPostion3.Set aPortXvalue(iIndex2), aPortYvalue(iIndex2), aPortZvalue(iIndex2)
            
        Set oVector12 = oPostion2.Subtract(oPostion1)
        Set oVector23 = oPostion3.Subtract(oPostion2)
            
        oVector12.Length = 1#
        oVector23.Length = 1#
            
        ' if the Cross Product is mainly Positive
        ' Assume Counter ColckWise direction, want Clockwise direction
        Set oCross2 = oVector12.Cross(oVector23)
        oCross2 = 1#
        If Abs(oCross2.z) > Abs(oCross2.x) Then
            If Abs(oCross2.z) > Abs(oCross2.y) Then
                If oCross2.z >= C_TOL001 Then
                    bReverse = True
                    Exit For
                End If
            Else
                If oCross2.y >= C_TOL001 Then
                    bReverse = True
                    Exit For
                End If
            End If
        ElseIf Abs(oCross2.y) > Abs(oCross2.x) Then
            If oCross2.y >= C_TOL001 Then
                bReverse = True
                Exit For
            End If
        Else
            If oCross2.x >= C_TOL001 Then
                bReverse = True
                Exit For
            End If
        End If
        
        Set oCross2 = Nothing
        oPostion1.Set oPostion2.x, oPostion2.y, oPostion2.z
        oPostion2.Set oPostion3.x, oPostion3.y, oPostion3.z
    Next iIndex1
        
    If bReverse Then
        For iIndex1 = 1 To nPorts
            aOrderFlags(nPorts - iIndex1 + 1) = oOrderIndex(iIndex1)
        Next iIndex1
        
        For iIndex1 = 1 To nPorts
            oOrderIndex(iIndex1) = aOrderFlags(iIndex1)
        Next iIndex1
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
            
End Sub

Public Sub Get_ObjectTypeData(oObject As Object, _
                              sTypeObject As String, sObjectType As String)
Const METHOD = "Get_ObjectTypeData"
On Error GoTo ErrorHandler

Dim lType As Long
Dim oBeam As IJBeam
Dim oSeam As IJSeam
Dim oPlate As IJPlate
Dim oSmartPlate As IJSmartPlate
Dim oStiffener As IJStiffener
Dim oCrossSection As IJCrossSection
Dim oProfileSection As IJDProfileSection

Dim oAutomationSeam As StructDetailObjects.SEAM
Dim oAutomationBracket As StructDetailObjects.Bracket
Dim oAutomationChamfer As StructDetailObjects.Chamfer
Dim oAutomationPlatePart As StructDetailObjects.PlatePart
Dim oAutomationProfileSystem As StructDetailObjects.ProfileSystem
Dim oAutomationEdgeReinforcement As StructDetailObjects.EdgeReinforcement
    
    On Error Resume Next
    
    If TypeOf oObject Is IJAssemblyConnection Then
        Set oBeam = oObject
        sTypeObject = "AssemblyConnection"
        sObjectType = Get_SmartOccurrenceClassName(oObject)
    
    ElseIf TypeOf oObject Is IJSmartPlate Then
        Set oSmartPlate = oObject
        Set oAutomationBracket = New StructDetailObjects.Bracket
        Set oAutomationBracket.object = oObject
        sTypeObject = "Bracket"
        sObjectType = oAutomationBracket.ClassName
        
    ElseIf TypeOf oObject Is IJChamfer Then
        Set oAutomationChamfer = New StructDetailObjects.Chamfer
        Set oAutomationChamfer.object = oObject
        sTypeObject = "Chamfer"
        sObjectType = oAutomationChamfer.ChamferType
        
    ElseIf TypeOf oObject Is IJCollarPart Then
        Set oPlate = oObject
        sTypeObject = "CollarPart"
        sObjectType = oPlate.plateType
        
    'IJStructFeature supports: FlangeCut, Slot, WebCut
    ElseIf TypeOf oObject Is IJStructFeature Then
        sTypeObject = "IJStructFeature"
        sObjectType = Get_SmartOccurrenceClassName(oObject)
        
    ElseIf TypeOf oObject Is IJStructPhysicalConnection Then
        sTypeObject = "PhysicalConnection"
        sObjectType = Get_SmartOccurrenceClassName(oObject)
    
    ElseIf TypeOf oObject Is IJPlatePart Then
        Set oAutomationPlatePart = New StructDetailObjects.PlatePart
        Set oAutomationPlatePart.object = oObject
        sTypeObject = "PlatePart"
        sObjectType = Get_StructPlateType(oAutomationPlatePart.plateType)
    
    ElseIf TypeOf oObject Is IJPlateSystem Then
        Set oAutomationPlatePart = New StructDetailObjects.PlatePart
        Set oAutomationPlatePart.object = oObject
        sTypeObject = "PlateSystem"
        sObjectType = Get_StructPlateType(oAutomationPlatePart.plateType)
    
    ElseIf TypeOf oObject Is IJPlane Then
        Set oPlate = oObject
        sTypeObject = GetObjectName(oObject, False)
        sObjectType = "RefPlane"
    
    ElseIf TypeOf oObject Is IJLandCurve Then
        Set oPlate = oObject
        sTypeObject = "IJLandCurve"
        sObjectType = "Geometry"
    
    Else
        sTypeObject = TypeName(oObject)
        sObjectType = "Unknown"
    End If


Exit Sub
    
ErrorHandler:
      Err.Raise Err.Number
            
End Sub

Public Function IsMinimumPoint(oPointToCheck As IJDPosition, _
                               oMinPoint As IJDPosition) As Boolean
Const METHOD = "IsMinimumPoint"
On Error GoTo ErrorHandler
    
    On Error Resume Next
    IsMinimumPoint = False
            
    Dim dValue1 As Double
    Dim dValue2 As Double
    Dim dValue3 As Double
    
    dValue1 = Abs(oPointToCheck.x - oMinPoint.x)
    dValue2 = Abs(oPointToCheck.y - oMinPoint.y)
    dValue3 = Abs(oPointToCheck.z - oMinPoint.z)
                
    ' Check if PointToCheck
    '   has minimum X value / has minimum abs(Y) value / has minimum Z value
    If dValue1 >= C_TOL001 And oPointToCheck.x < oMinPoint.x Then
        IsMinimumPoint = True
    ElseIf dValue2 >= C_TOL001 And Abs(oPointToCheck.y) < Abs(oMinPoint.y) Then
        IsMinimumPoint = True
    ElseIf dValue3 >= C_TOL001 And oPointToCheck.z < oMinPoint.z Then
        IsMinimumPoint = True
    End If
            

Exit Function

ErrorHandler:
    Err.Raise Err.Number

End Function

Public Function Get_SmartOccurrenceClassName(oObject As Object) As String
Const METHOD = "Congifuration.Get_StructFeatureClassName"
On Error GoTo ErrorHandler

    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Set oSmartOccurrence = oObject
    If Not oSmartOccurrence Is Nothing Then
        Set oSmartItem = oSmartOccurrence.ItemObject
        If Not oSmartItem Is Nothing Then
            Set oSmartClass = oSmartItem.Parent
        Else
            Set oSmartClass = oSmartOccurrence.ItemObject
        End If
    End If
    
    If Not oSmartClass Is Nothing Then
        Get_SmartOccurrenceClassName = oSmartClass.SCName
    End If
    
    Set oSmartItem = Nothing
    Set oSmartClass = Nothing
    Set oSmartOccurrence = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
            
End Function

Public Function Get_StructPlateType(eType As Long) As String
Const METHOD = "Congifuration.Get_StructPlateType"
On Error GoTo ErrorHandler
Dim sType As String

    '   Convert Enum of StructPlateType to "known" string
    If eType = DeckPlate Then
        sType = "DeckPlate"
    ElseIf eType = Hull Then
        sType = "Hull"
    ElseIf eType = LBulkheadPlate Then
        sType = "LBulkheadPlate"
    ElseIf eType = TBulkheadPlate Then
        sType = "TBulkheadPlate"
    ElseIf eType = UntypedPlate Then
        sType = "UntypedPlate"
    Else
        sType = Trim(Str(eType))
    End If

    Get_StructPlateType = sType
Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
            
End Function

Public Function GetObjectName(oObject As Object, bIncludeType As Boolean) As String
On Error GoTo ErrorHandler
    Const METHOD = "GetObjectName"

    Dim sText As String
    Dim oJNamedItem As IJNamedItem
    Dim EntityNaming As IJDStructEntityNaming
        
    On Error Resume Next
    GetObjectName = ""
    
    sText = ""
    Set EntityNaming = oObject
    If EntityNaming Is Nothing Then
        Set oJNamedItem = oObject
        If oJNamedItem Is Nothing Then
            sText = sText & "??Name??"
        Else
            sText = sText & oJNamedItem.name
        End If
    Else
        sText = sText & EntityNaming.name
    End If
            
    If bIncludeType Then
        sText = sText & " ... " & TypeName(oObject)
    End If
    
    GetObjectName = sText
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
            
End Function

Public Function GetPlatePort(ByVal oPlate As IJPlate) As Object
Const METHOD = "GetPlatePortGeometry"

On Error GoTo ErrorHandler
     
    Dim oPlateGeometry As IJStructGeometry
    Set oPlateGeometry = oPlate
            
    Dim oTopoLocate As IJTopologyLocate
    Set oTopoLocate = New TopologyLocate
    
    Dim oEnumPorts As IEnumUnknown
    ' Filter
    ' JS_TOPOLOGY_FILTER_ALL_LFACES
    oTopoLocate.GetNamedPorts oPlateGeometry, JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, oEnumPorts
     
    If Not (oEnumPorts Is Nothing) Then
    
        'Convert the IEnumUnknown to a VB collection that we can use in VB
        Dim oCollectionOfPorts As Collection
        Dim oConvertUtils As New CCollectionConversions
        oConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPorts, oCollectionOfPorts
        
        'Get first port on List
        Dim oPort As IJPort
        Set oPort = oCollectionOfPorts.Item(1)
        Set GetPlatePort = oPort '.Geometry
    Else
        MsgBox " no ports returned from GetNamedPorts" 
    End If
                
Cleanup:
    Set oTopoLocate = Nothing
    Set oEnumPorts = Nothing
    Set oConvertUtils = Nothing
    Set oCollectionOfPorts = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    Set GetPlatePort = Nothing
    GoTo Cleanup
     
End Function



'This function will find aft, fore butts and upper, lower seams in the shell plate.
'And then return the type of port which are corresponding to given index.

Public Function FindOutEdgePositions(oElemsPort As IJElements, orderedIdx As Long) As String
    On Error GoTo ErrorHandler
    Const METHOD = "FindOutEdgePositions"
    
    Dim i               As Integer
    Dim edgeLists       As IEnumUnknown
    Dim aaa             As IEnumUnknown
''    Dim oPortSeam   As CPortSeam
    Dim oCollPorts      As New Collection
    Dim oConvert        As New CCollectionConversions
    Dim oGeomHelper     As New MfgGeomHelper
    
    Dim oAftButt        As IJWireBody
    Dim oForeButt       As IJWireBody
    Dim oUpperSeam1     As IJWireBody
    Dim oUpperSeam2     As IJWireBody   'In case of pentagon plate
    Dim oLowerSeam1     As IJWireBody
    Dim oLowerSeam2     As IJWireBody   'In case of pentagon plate
 
    'Collect all ports geometry
    For i = 1 To oElemsPort.Count
        
        oCollPorts.Add oElemsPort.Item(i).Geometry
    Next
    
    'Convert vb collection to edgeLists in order to use that in the GetButtLinesAndSeamLines method
    oConvert.CreateIEnumUnknownFromVBCollection oCollPorts, edgeLists
    
    Select Case oElemsPort.Count
        Case 3 'In case of triangle shpae shell plate.
        
        Case 4 'In case of rectangle shape shell plate.
        
'            oGeomHelper.GetButtLinesAndSeamLines edgeLists, oAftButt, oForeButt, oLowerSeam1, oUpperSeam1
            
'            MsgBox TypeName(oAftButt)
'            MsgBox TypeName(oForeButt)
'
'            MsgBox TypeName(oLowerSeam1)
'            MsgBox TypeName(oUpperSeam1)
          
            If oAftButt Is oElemsPort.Item(orderedIdx).Geometry Then
                FindOutEdgePositions = "AftButt"
            
            ElseIf oForeButt Is oElemsPort.Item(orderedIdx).Geometry Then
                FindOutEdgePositions = "ForeButt"
                
            ElseIf oLowerSeam1 Is oElemsPort.Item(orderedIdx).Geometry Then
                FindOutEdgePositions = "LowerSeam"
            
            ElseIf oUpperSeam1 Is oElemsPort.Item(orderedIdx).Geometry Then
                FindOutEdgePositions = "UpperSeam"
                
            End If
                        
        Case 5 'In case of pentagon shape shell plate.
            
    
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

Public Function CheckLongiProfile(oProfilePart As IJStiffenerPart) As Boolean
    On Error GoTo ErrorHandler
    Const METHOD = "CheckLongiProfile"
    
    Dim oProfile    As IJStiffener
    
    Set oProfile = oProfilePart
    
   
    If oProfile.pType = sptLongitudinal Then
        CheckLongiProfile = True
    Else
        CheckLongiProfile = False
    End If

    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

Public Function CheckPlateType(oPlatePart As IJPlatePart) As StructPlateType
    On Error GoTo ErrorHandler
    Const METHOD = "CheckLongiPlate"
    
    Dim oPlate      As IJPlate
    
    Set oPlate = oPlatePart
    
    If oPlate.plateType = LBulkheadPlate Then
        CheckPlateType = LBulkheadPlate
    
    ElseIf oPlate.plateType = DeckPlate Then
        CheckPlateType = DeckPlate
        
    ElseIf oPlate.plateType = TBulkheadPlate Then
        CheckPlateType = TBulkheadPlate
    
    Else
        oPlate.plateType = UntypedPlate
        CheckPlateType = UntypedPlate
    
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number

End Function

'Get Physical connection
Public Function GetPhysicalConnection(oPlatePart As IJPlatePart) As IJStructPhysicalConnection
    On Error GoTo ErrorHandler
    Const METHOD = "GetPhysicalConnection"
    
    Dim oPlate          As IJPlate
    Dim oSystem         As IJSystem
    Dim oSystemChild    As IJSystemChild
    Dim oChildren       As IJDTargetObjectCol
    Dim iCnt            As Integer
    
    Set oSystemChild = oPlatePart
    Set oPlate = oSystemChild.GetParent
    
    Set oSystem = oPlate
    
    'Get IJStructConnection
    Set oChildren = oSystem.GetChildren()
    Set oSystem = Nothing
    For iCnt = 1 To oChildren.Count
        If TypeOf oChildren.Item(iCnt) Is IJStructConnection Then
            Set oSystem = oChildren.Item(iCnt) 'This will be IJStructConnection
        End If
    Next
    Set oChildren = Nothing
    
    'Get IJAssemblyConnection
    Set oChildren = oSystem.GetChildren()
    Set oSystem = Nothing
    For iCnt = 1 To oChildren.Count
        If TypeOf oChildren.Item(iCnt) Is IJAssemblyConnection Then
            Set oSystem = oChildren.Item(iCnt) 'This will be IJAssemblyConnection
            
        End If
    Next
    
    'Get IJStructPhysicalConnection
    Set oChildren = oSystem.GetChildren()
    Set oSystem = Nothing
    For iCnt = 1 To oChildren.Count
        If TypeOf oChildren.Item(iCnt) Is IJStructPhysicalConnection Then
            Set GetPhysicalConnection = oChildren.Item(iCnt)
            Exit For
        End If
    Next
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
    
End Function


 
