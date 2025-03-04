Attribute VB_Name = "MarginHelpers"
Public Sub GetEligiblePortCollectionForCan(oPlatePart As IJPlatePart, oPortColl As IJElements)
    Const METHOD = "GetEligiblePortCollectionForCan"
    On Error GoTo ErrorHandler
    
    'get the root and leaf plate systems
    Dim oRootPlateSystem    As IJPlateSystem
    Dim oLeafPlateSystem    As IJPlateSystem
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
        
    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oRootPlateSystem, True
    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oLeafPlateSystem
    
    Dim oPlateUtil As IJPlateAttributes
    Set oPlateUtil = New PlateUtils
    
    'Check if the Can was split
    Dim oLeafPlateSysColl As Collection
    Set oLeafPlateSysColl = oPlateUtil.GetSplitResults(oRootPlateSystem)
    
    If oLeafPlateSysColl.Count > 1 Then 'if can is split into two parts
         
        'get the splitters
        Dim oParent   As IUnknown
        Dim oSplitters As IEnumUnknown
        oStructDetailHelper.IsResultOfSplitWithOpr oLeafPlateSystem, oParent, oSplitters
        
        Dim oCollectionOfSplitters  As Collection
        Dim ConvertUtils            As CCollectionConversions
        Dim SplitterColl            As IJElements
        
        Set ConvertUtils = New CCollectionConversions
        ConvertUtils.CreateVBCollectionFromIEnumUnknown oSplitters, oCollectionOfSplitters
        Set SplitterColl = ConvertUtils.CreateIJElementsCollectionFromVBCollection(oCollectionOfSplitters)
        
        'Get the lateral face ports which are the result of split done with above splitters
        Set oPortColl = GetLateralFacePortsFromSplitters(oPlatePart, SplitterColl)
        
        Set oParent = Nothing
        Set oSplitters = Nothing
    End If
    
CleanUp:
    Set oRootPlateSystem = Nothing
    Set oLeafPlateSystem = Nothing
    Set oStructDetailHelper = Nothing
    Set oPlateUtil = Nothing
    Set oLeafPlateSysColl = Nothing
    Set SplitterColl = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp

End Sub

Private Function GetLateralFacePortsFromSplitters(oPlatPart As IJPlatePart, SplitterColl As IJElements) As IJElements
    Const METHOD = "GetLateralFacePortsFromSplitters"
    On Error GoTo ErrorHandler
    
    Set GetLateralFacePortsFromSplitters = New JObjectCollection
    
    Dim oStructConnectable  As IJStructConnectable
    Set oStructConnectable = oPlatPart
    
    Dim oFacePortsColl      As IJElements
    
    'Get the all lateral face ports. OperationProgID is NULL. So, it gives the latest geometry
    oStructConnectable.enumConnectableTransientPorts oFacePortsColl, vbNullString, False, PortFace, JS_TOPOLOGY_FILTER_SOLID_LATERAL_LFACES, False
    
    Dim oStructDetailHelper As IJStructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
     
    Dim iCount As Long
    For iCount = 1 To oFacePortsColl.Count
    
       Dim oPort As IJStructPort
       Set oPort = oFacePortsColl.Item(iCount)
       
       Dim oOperator As Object
       Dim oOperation As IJStructOperation
       
       oStructDetailHelper.FindOperatorForOperationInGraphByID oPlatPart, oPort.OperationID, oPort.OperatorID, oOperation, oOperator
               
       If SplitterColl.Contains(oOperator) Then 'It is the port at the split
                
            'Donot include ports that are curved.
            If CheckIfPortIsLinear(oPort) Then
                Dim oTransientPort As IJTransientPort
                Set oTransientPort = oPort

                Dim oStructPort As IJStructPortEx
                oTransientPort.GetPersistentPort oStructPort, True
               
                GetLateralFacePortsFromSplitters.Add oStructPort.PrimaryPort
           
            End If
        
       End If
        
       Set oPort = Nothing
       Set oOperator = Nothing
       Set oOperation = Nothing

NextItem:
    Next
    
    
CleanUp:
    Set oStructConnectable = Nothing
    Set oFacePortsColl = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Function CheckIfPortIsLinear(oFacePort As IJStructPort) As Boolean
     Const METHOD = "CheckIfPortIsLinear"
     On Error GoTo ErrorHandler
    
     CheckIfPortIsLinear = False
     
     Dim oEntityHelper As New MfgEntityHelper
     Dim oEdgePort     As IJPort
     Set oEdgePort = oEntityHelper.GetEdgePortGivenFacePort(oFacePort, CTX_BASE) 'This gives the latest edge port geometry
    
     'Querying for IJLine on the port object was giving incorrect results for certain test cases(conical cross section)
     'Also, querying for IJLine on port geometry did not work.
     'So,using the MinBox routine to check if the port is linear.

     Dim oPortCurve As IJWireBody
     Set oPortCurve = oEdgePort.Geometry
           
     'get the MinBox to determine if the wirebody is a line or not
     Dim oMfgGeomHelper As MfgGeomHelper
     Set oMfgGeomHelper = New MfgGeomHelper
     Dim oMfgMGHelper As MfgMGHelper
     Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
     
     Dim oCurveElems As IJElements
     
     oMfgMGHelper.WireBodyToComplexStrings oPortCurve, oCurveElems
               
     Dim oBoxPoints As IJElements
     Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBox(oCurveElems)

     Dim length1 As Double, length2 As Double, length3 As Double
     Dim Points(1 To 4) As IJDPosition
     
     Set Points(1) = oBoxPoints.Item(1)
     Set Points(2) = oBoxPoints.Item(2)
     Set Points(3) = oBoxPoints.Item(3)
     Set Points(4) = oBoxPoints.Item(4)

     length1 = Points(1).DistPt(Points(2))
     length2 = Points(2).DistPt(Points(3))
     length3 = Points(1).DistPt(Points(4))
   
     'Two of the dimensions should be below the tolerance
     If length1 < 0.001 Then
        If length2 < 0.001 Or length3 < 0.001 Then
          CheckIfPortIsLinear = True
        End If
     ElseIf length2 < 0.01 Then
        If length1 < 0.01 Or length3 < 0.01 Then
          CheckIfPortIsLinear = True
        End If
     Else 'If length3 < 0.01 Then
        If length1 < 0.01 Or length2 < 0.01 Then
            CheckIfPortIsLinear = True
        End If
     End If
     
CleanUp:
     Set oEntityHelper = Nothing
     Set oMfgMGHelper = Nothing
     Set oMfgGeomHelper = Nothing
     Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
Public Function GetConnectionsBetweenAssemblies(obj1 As Object, obj2 As Object) As IJElements
Const METHOD As String = "GetConnectionsBetweenAssemblies"

    If (Not obj1 Is Nothing) Or (Not obj2 Is Nothing) Then
        Dim oMfgUtilAssyMargin As IJMfgUtilAssyMargin
        Set oMfgUtilAssyMargin = New MfgUtilAssyMargin
        
        Dim oConnections As IJElements
        Set oConnections = New JObjectCollection
        
        oMfgUtilAssyMargin.GetConnectedChildren obj1, obj2, oConnections
        
        Set GetConnectionsBetweenAssemblies = oConnections
    End If
        
    Exit Function
     
ErrorHandler:
 Err.Raise Err.Number, Err.Source, Err.Description

End Function

'----------------------------------------------------------------------------
'
' Method
'   GetSubAssemblies()
'
' Description
'   Returns a list of subassemblies to a selected block or assembly
'
'----------------------------------------------------------------------------

Public Function GetSubAssemblies(ByVal oElem As IJAssembly) As IJElements
Const METHOD As String = "GetSubAssemblies"
   
    On Error GoTo ErrorHandler
    
    'Get children of assembly
    Dim oChildren As IJDTargetObjectCol
    Dim oAssembly As IJAssembly
    Set oAssembly = oElem
    Set oChildren = oAssembly.GetChildren
                
    'Only highlight assemblies with assembly children
    If oChildren.Count > 0 Then
        If TypeOf oAssembly Is IJAssembly Then
            Dim oChild As IJAssemblyChild
            
            Dim oElements As IJElements
            Set oElements = New JObjectCollection
            
            'Add the block or assembly itself to the list
            oElements.Add oElem
                        
            'Get each assembly child and add child to elements list
            Dim index As Long
            For index = 1 To oChildren.Count
                'Get child
                Set oChild = oChildren.Item(index)
                                               
                'Add child to list if type of element is right
                If TypeOf oChild Is IJAssembly Then
                    oElements.Add oChild
                    
                    Dim oSubAssys As IJElements
                    Set oSubAssys = GetSubAssemblies(oChild)
                    
                    If Not oSubAssys Is Nothing Then
                        oElements.AddElements oSubAssys
                    End If
                    
                    Set oSubAssys = Nothing
                End If
            Next
            Set GetSubAssemblies = oElements
         End If
    End If
    
    'Clean up
    Set oAssembly = Nothing
    Set oChildren = Nothing
    Set oElements = Nothing
    Set oChild = Nothing
    
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

