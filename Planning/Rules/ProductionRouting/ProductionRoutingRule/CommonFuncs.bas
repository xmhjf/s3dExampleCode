Attribute VB_Name = "CommonFuncs"
Option Explicit



Public Sub GetDefBendingMachineAndCodeByType(oXMLNode As IXMLDOMNode, oPart As Object, strMcName As String, strMcCode As String)
Const METHOD = "GetDefBendingMachineAndCodeByType"
On Error GoTo ErrorHandler

    Dim strType                     As String
    Dim oChildXMLNode               As IXMLDOMNode
    Dim oSAAttribNode               As IXMLDOMAttribute
    Dim bResultFound                As Boolean
    Dim bPlanarPart                 As Boolean
    
    bResultFound = False
    
    If TypeOf oPart Is IJPlatePart Then
        strType = "Plate"
    ElseIf TypeOf oPart Is IJProfilePart Then
        strType = "Profile"
    End If
    
    Dim oplatepart As IJPlatePart
    Dim oPRofilepart As IJProfilePart
    
    Dim oSDPartSupport          As IJPartSupport
    Dim oSDPlatePartSupport     As IJPlatePartSupport
    
    If TypeOf oPart Is IJPlatePart Then
        Set oplatepart = oPart
        
        Set oSDPlatePartSupport = New PlatePartSupport
        Set oSDPartSupport = oSDPlatePartSupport
        Set oSDPartSupport.Part = oplatepart
        
        Dim bIsPlanar As Boolean
        bIsPlanar = False
        If oSDPlatePartSupport.IsShellPlate = False Then
            oSDPlatePartSupport.IsPlanar 0.001, bIsPlanar
            
            If bIsPlanar = True Then ' do not add routing actions to planar plates
                bPlanarPart = True
            End If
        End If
    
    End If
    
    If TypeOf oPart Is IJProfilePart Then
        
        Dim oTopologyLocate As TopologyLocate
        Dim oLandingCrv     As IJWireBody
        Dim oUtils As SGOWireBodyUtilities

        Set oTopologyLocate = New TopologyLocate
        Set oUtils = New SGOWireBodyUtilities

        Set oLandingCrv = oTopologyLocate.GetProfileParentWireBody(oPart)

        'If non-linear , add to excluded parts
        If oUtils.IsLinear(oLandingCrv) <> 0 Then
            bPlanarPart = True
        End If

        Set oLandingCrv = Nothing

    End If
    
    For Each oChildXMLNode In oXMLNode.childNodes
    
        For Each oSAAttribNode In oChildXMLNode.Attributes
            If (UCase(oSAAttribNode.nodeName) = UCase("Type")) And (UCase(oSAAttribNode.nodeValue) = UCase(strType)) Then
                bResultFound = True
                Exit For
            End If
        Next
        
        If bResultFound Then
            For Each oSAAttribNode In oChildXMLNode.Attributes
                If (UCase(oSAAttribNode.nodeName) = UCase("Name")) Then
                    If bPlanarPart = False Then
                        strMcName = oSAAttribNode.nodeValue
                    Else
                        strMcName = "None"
                    End If
                ElseIf (UCase(oSAAttribNode.nodeName) = UCase("Default")) Then
                    If bPlanarPart = False Then
                        strMcCode = oSAAttribNode.nodeValue
                    Else
                        strMcCode = "None"
                    End If
                End If
            Next
            
            Exit For
        End If
    Next
    

Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Public Sub GetPrimingMachineAndCode(oXMLNode As IXMLDOMNode, oPart As Object, strMcName As String, strMcCode As String)
Const METHOD = "GetPrimingMachineAndCode"
On Error GoTo ErrorHandler

    Dim oIJStructureMaterial        As IJStructureMaterial
    Dim strMaterialGrade            As String
    Dim oChildXMLNode               As IXMLDOMNode
    Dim oSAAttribNode               As IXMLDOMAttribute
    Dim bResultFound                As Boolean
    
    bResultFound = False
    
    If TypeOf oPart Is IJStructureMaterial Then
        Set oIJStructureMaterial = oPart
        
        strMaterialGrade = oIJStructureMaterial.MaterialGrade
        
        For Each oChildXMLNode In oXMLNode.childNodes
        
            For Each oSAAttribNode In oChildXMLNode.Attributes
                If (UCase(oSAAttribNode.nodeName) = UCase("Grade")) And (UCase(oSAAttribNode.nodeValue) = UCase(strMaterialGrade)) Then
                    bResultFound = True
                    Exit For
                End If
            Next
            
            If bResultFound Then
                For Each oSAAttribNode In oChildXMLNode.Attributes
                    If (UCase(oSAAttribNode.nodeName) = UCase("Name")) Then
                        strMcName = oSAAttribNode.nodeValue
                    ElseIf (UCase(oSAAttribNode.nodeName) = UCase("Default")) Then
                        strMcCode = oSAAttribNode.nodeValue
                    End If
                Next
                
                Exit For
            End If
        Next
    End If
    

Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Public Function GetStageCodeColl(strXMLFile As String, oPart As Object) As Collection
Const METHOD = "GetStageCodeColl"
On Error GoTo ErrorHandler

    
    If strXMLFile <> vbNullString Then
        Dim oNodeList As IXMLDOMNodeList
        Dim oXMLDOC As New MSXML2.DOMDocument60
        Dim oXMLNode As IXMLDOMNode
        Dim oXMLNodeList As IXMLDOMNodeList
        Dim oChildXMLNode As IXMLDOMNode
        Dim oChildXMLNodes As IXMLDOMNodeList
        Dim oSubChildXMLNode As IXMLDOMNode
        Dim oSubChildXMLNodes As IXMLDOMNodeList
        Dim oSCAttribNode As IXMLDOMAttribute
        Dim oSAAttribNode As IXMLDOMAttribute
        Dim strXMLCSName As String
        Dim strAssemblyType As String
        
        Dim oStageCodeColl As Collection
        
        Set oStageCodeColl = New Collection
        
        Dim lPartAssemblyType As Long
        Dim lAssemblyType As Long
        
        Dim oAsseBase As IJAssemblyBase
        Dim oAsmChild As IJAssemblyChild
        
        Dim oCodeListMetaData As IJDCodeListMetaData
        Dim oTempcodeListColl As IJDInfosCol
        Dim oCodelist As IJDCodeListValue
        Dim iIndex As Long
        
        Set oAsmChild = oPart
        
        On Error Resume Next
        Set oAsseBase = oAsmChild.Parent
        On Error GoTo ErrorHandler
        
        'get the part's assembly's assembly type
        If Not oAsseBase Is Nothing Then
            lPartAssemblyType = oAsseBase.Type
            Set oCodeListMetaData = oAsmChild.Parent
        End If
           
    
        'Load the file into XMLDOM
        oXMLDOC.Load (strXMLFile)
        Set oXMLNodeList = oXMLDOC.getElementsByTagName("StageCodes")

        For Each oXMLNode In oXMLNodeList
            Dim strAction As String
            For Each oSAAttribNode In oXMLNode.Attributes
                If (UCase(oSAAttribNode.nodeName) = UCase("AssemblyType")) Then
                    strAssemblyType = oSAAttribNode.nodeValue
                    Exit For
                End If
            Next oSAAttribNode
            
           'get the assembly type long value read through xml
            If Not oCodeListMetaData Is Nothing Then
            
                Set oTempcodeListColl = oCodeListMetaData.CodelistValueCollection("AssemblyType")
                If Not oTempcodeListColl Is Nothing Then
                    If oTempcodeListColl.Count > 0 Then
                        For iIndex = 1 To oTempcodeListColl.Count - 1
                            Set oCodelist = oTempcodeListColl.Item(iIndex)
                            If UCase(oCodelist.LongValue) = UCase(strAssemblyType) Then
                                lAssemblyType = oCodelist.ValueID
                                Set oCodelist = Nothing
                                Exit For
                            End If
                            Set oCodelist = Nothing
                        Next iIndex
                    End If
                End If
                
            End If
            
           
            If (lAssemblyType = lPartAssemblyType) Then

                For Each oChildXMLNode In oXMLNode.childNodes
                    If UCase(oChildXMLNode.nodeName) = UCase("StageCode") Then
                        
                        For Each oSAAttribNode In oChildXMLNode.Attributes
                            Dim strMachine As String
                            Dim strCode As String
                            If (UCase(oSAAttribNode.nodeName) = UCase("UserName")) Then
                                oStageCodeColl.Add oSAAttribNode.nodeValue
                            End If
                        Next oSAAttribNode
                        
                    End If
    
                Next oChildXMLNode
                 Exit For
            End If

        Next oXMLNode
    End If
    
    Set GetStageCodeColl = oStageCodeColl

Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Public Function GetStageCode(strXmlFilename As String, pPart As Object) As String
Const METHOD = "GetStageCode"
On Error GoTo ErrorHandler


    GetStageCode = ""
    
    Dim oStageCodeColl As Collection
    
    Set oStageCodeColl = New Collection
    
    Set oStageCodeColl = GetStageCodeColl(strXmlFilename, pPart)
    
    If oStageCodeColl.Count > 0 Then
        GetStageCode = oStageCodeColl.Item(1)
    End If
    
 
Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Public Sub GetBendingMachineAndCode(oPart As Object, strMcName As String, strMcCode As String)
Const METHOD = "GetBendingMachineAndCode"
On Error GoTo ErrorHandler
    
    Dim oRootPlateSys           As IJSystem
    Dim oDesignChild            As IJDesignChild
    Dim oChildren               As IJDTargetObjectCol
    Dim oKnuckle                As IJKnuckle
    Dim oModelBody              As IJDModelBody
    Dim oKnuckleWireBody        As IJWireBody
    Dim oEdgeWireBody           As IJWireBody
    Dim dArea                   As Double
    Dim dLength                 As Double
    Dim dVolume                 As Double
    Dim dAchAccuracy            As Double
    Dim i                       As Long
    Dim oAsChild                As IJAssemblyChild
    Dim oAsBase                 As IJAssemblyBase
    Dim oplatepart              As IJPlatePart
    Dim oPlnGeomHelper          As PlnGeometryHelper.PlnGeometryHelper
    Dim oEdges                  As IJDObjectCollection
    Dim oBasePort               As IJPort
    Dim oSDOPlatePart           As StructDetailObjects.PlatePart
    Dim oEndPt1                 As IJDPosition
    Dim oEndPt2                 As IJDPosition
    Dim oKEndPt1                As IJDPosition
    Dim oKEndPt2                As IJDPosition
    Dim dFirstDistance          As Double
    Dim dSecondDistance         As Double
    Dim oKPointBeingUsed        As IJDPosition
    Dim oSDOProfilePart         As StructDetailObjects.ProfilePart
    Dim oStiffenedPlate         As IJPlate
    Dim bIsSystem               As Boolean
    Dim oProfKnuckle            As ProfileKnuckle
    Dim oKnuckleCollection      As IJElements
    Dim pIJProfileAttributes    As IJProfileAttributes
    Dim oProfileUtils           As IJProfileDefinition
    Dim oProfileSys             As IJProfile
    
    Set oAsChild = oPart
    On Error Resume Next
    Set oAsBase = oAsChild.Parent
    On Error GoTo ErrorHandler
    
    Dim bPlanarPart                 As Boolean
    Dim oSDPartSupport          As IJPartSupport
    Dim oSDPlatePartSupport     As IJPlatePartSupport
    
    If TypeOf oPart Is IJPlatePart Then
       
        Set oSDPlatePartSupport = New PlatePartSupport
        Set oSDPartSupport = oSDPlatePartSupport
        Set oSDPartSupport.Part = oPart
        
        Dim bIsPlanar As Boolean
        bIsPlanar = False
        If oSDPlatePartSupport.IsShellPlate = False Then
            oSDPlatePartSupport.IsPlanar 0.001, bIsPlanar
            
            If bIsPlanar = True Then ' do not add routing actions to planar plates
                bPlanarPart = True
            End If
        End If
    
    End If
    
    If TypeOf oPart Is IJProfilePart Then
        
        Dim oTopologyLocate As TopologyLocate
        Dim oLandingCrv     As IJWireBody
        Dim oUtils As SGOWireBodyUtilities

        Set oTopologyLocate = New TopologyLocate
        Set oUtils = New SGOWireBodyUtilities

        Set oLandingCrv = oTopologyLocate.GetProfileParentWireBody(oPart)

        'If non-linear , add to excluded parts
        If oUtils.IsLinear(oLandingCrv) <> 0 Then
            bPlanarPart = True
        End If

        Set oLandingCrv = Nothing

    End If
    
    If bPlanarPart Then
        strMcName = "None"
        strMcCode = "None"
        Exit Sub
    End If
    
    
    
    
    If TypeOf oPart Is IJPlatePart Then
    
        Set oplatepart = oPart
    
        Set oDesignChild = oPart
        Set oRootPlateSys = oDesignChild.GetParent
        
        Set oDesignChild = oRootPlateSys
        Set oRootPlateSys = oDesignChild.GetParent
        
        Set oChildren = oRootPlateSys.GetChildren
        
        For i = 1 To oChildren.Count
            If TypeOf oChildren.Item(i) Is IJKnuckle Then
                Set oKnuckle = oChildren.Item(i)
                Exit For
            End If
        Next
        
        If oKnuckle Is Nothing Then Exit Sub
        Set oModelBody = oKnuckle.GetKnuckleCurve
        oModelBody.GetDimMetrics 0.001, dAchAccuracy, dLength, dArea, dVolume
                
        If dLength > 1.4 And dLength <= 2.2 Then
            strMcName = "500T Press"
            
            If Not oAsBase Is Nothing Then
                If oAsBase.Type = 7 Or oAsBase.Type = 15 Then
                    strMcCode = "5BJ"
                Else
                    strMcCode = "5BS"
                End If
            End If
        Else
            If oplatepart.PlateWidth <= 4.5 And oplatepart.PlateLength <= 4.5 Then
                'Distance between knuckle and edges
                
                Set oKnuckleWireBody = oKnuckle
                oKnuckleWireBody.GetEndPoints oKEndPt1, oKEndPt2
                
                Set oSDOPlatePart = New StructDetailObjects.PlatePart
                Set oSDOPlatePart.object = oPart
                Set oBasePort = oSDOPlatePart.BasePort(BPT_Base)
                
                Set oPlnGeomHelper = New PlnGeometryHelper.PlnGeometryHelper
                Set oEdges = oPlnGeomHelper.GetExternalEdgesOfSheet(oBasePort.Geometry)
                
                For Each oEdgeWireBody In oEdges
                    oEdgeWireBody.GetEndPoints oEndPt1, oEndPt2
                    
                    If oKPointBeingUsed Is Nothing Then
                        If oKEndPt1.DistPt(oEndPt1) < 0.001 Then
                            Set oKPointBeingUsed = oKEndPt1
                            dFirstDistance = oKEndPt1.DistPt(oEndPt2)
    
                        ElseIf oKEndPt1.DistPt(oEndPt2) < 0.001 Then
                            Set oKPointBeingUsed = oKEndPt1
                            dFirstDistance = oKEndPt1.DistPt(oEndPt1)
                        
                        ElseIf oKEndPt2.DistPt(oEndPt1) < 0.001 Then
                            Set oKPointBeingUsed = oKEndPt2
                            dFirstDistance = oKEndPt2.DistPt(oEndPt2)
                        
                        ElseIf oKEndPt2.DistPt(oEndPt2) < 0.001 Then
                            Set oKPointBeingUsed = oKEndPt2
                            dFirstDistance = oKEndPt2.DistPt(oEndPt1)
                        
                        End If
                    Else
                        If oKPointBeingUsed.DistPt(oEndPt1) < 0.001 Then
                            dSecondDistance = oKEndPt1.DistPt(oEndPt2)
    
                        ElseIf oKPointBeingUsed.DistPt(oEndPt2) < 0.001 Then
                            dSecondDistance = oKEndPt1.DistPt(oEndPt1)
                        End If
                    End If
                    
                    If dFirstDistance > 0# And dSecondDistance > 0# Then
                        Exit For
                    End If
                Next
                
                If dFirstDistance <= 3.5 And dSecondDistance <= 3.5 Then
                    If dLength > 2.2 Then
                        strMcName = "1000T Press"
                        If Not oAsBase Is Nothing Then
                            If oAsBase.Type = 7 Or oAsBase.Type = 15 Then
                                strMcCode = "PBA"
                            Else
                                strMcCode = "PBJS"
                            End If
                        End If
                    Else
                    '1750 press
                        
                    End If
                Else
                '1750 press
                
                End If
            Else
            '1750 press
            
            
            End If
        End If
    ElseIf TypeOf oPart Is IJProfileER Then
        strMcName = "Profile Curved Bender Machine"
        strMcCode = "OBRS"

    ElseIf TypeOf oPart Is IJProfilePart Then
        Set oSDOProfilePart = New StructDetailObjects.ProfilePart
        Set oSDOProfilePart.object = oPart
        
        If oSDOProfilePart.SectionType = "R" Then
            strMcName = "Profile Curved Bender Machine"
            If Not oAsBase Is Nothing Then
                If oAsBase.Type = 7 Or oAsBase.Type = 15 Then
                    strMcCode = "OBRS"
                Else
                    strMcCode = "OBRA"
                End If
            End If
        ElseIf oSDOProfilePart.SectionType = "FB" Then
            If oSDOProfilePart.WebThickness <= 0.8 Then
                oSDOProfilePart.GetStiffenedPlate oStiffenedPlate, bIsSystem
                
                If oStiffenedPlate.thickness <= 0.03 Then
                    Set oProfileUtils = New ProfileUtils
                    Set pIJProfileAttributes = oProfileUtils
                    
                    Set oDesignChild = oPart
                    Set oProfileSys = oDesignChild.GetParent
                    
                    Set oDesignChild = oProfileSys
                    Set oProfileSys = oDesignChild.GetParent
                    
                    Set oKnuckleCollection = pIJProfileAttributes.GetKnucklesOnProfile(oProfileSys)
                    
                    For i = 1 To oKnuckleCollection.Count
                        If TypeOf oKnuckleCollection.Item(i) Is ProfileKnuckle Then
                            Set oProfKnuckle = oKnuckleCollection.Item(i)
                            Exit For
                        End If
                    Next
                    
                    If Not oProfKnuckle Is Nothing Then
                        If oProfKnuckle.Angle > 2.279 And oProfKnuckle.Angle < 3.142 Then
                            strMcName = "Face Bender Machine"
                            If Not oAsBase Is Nothing Then
                                If oAsBase.Type = 7 Or oAsBase.Type = 15 Then
                                    strMcCode = "FBRS"
                                Else
                                    strMcCode = "FBRA"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    

    
    
Exit Sub
ErrorHandler:
     MsgBox Err.Description
End Sub

'Fills the report with the Compartment Coating Type and Coating Level attributes
Public Sub ComputeproductionData(oSatgeProduction As IJDProductionRouting, oPart As Object)
On Error GoTo ErrorHandler
Const METHOD = "ComputeproductionData"
    
    Dim oAttributeMetaData          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim oPropColl                   As Collection
    Dim lIndex                      As Long
    Dim strXmlFilename              As String
    Dim oContext                    As IJContext
    Dim oSRDQuery                   As IJSRDQuery
    Dim oRuleQuery                  As IJSRDRuleQuery
    Dim oRule                       As IJSRDRule
    Dim oRuleObj                    As Object
    Dim strRuleNames()              As String
    Dim strRuleProgId               As String
    Dim ii                          As Integer
    
    strRuleProgId = ""
    
    Set oSRDQuery = New SRDQuery
    Set oRuleQuery = oSRDQuery.GetRulesQuery
    
    Set oContext = GetJContext()
    
    'Getting the Production Routing xml path from the catalog database
    If Not oRuleQuery Is Nothing Then
    On Error Resume Next
            strRuleNames = oRuleQuery.GetRuleNamesByType(200, 212, True)
            If Not UBound(strRuleNames) >= 0 Then
            strRuleNames = oRuleQuery.GetRuleNamesByType(400, 420, True)
            End If
     Err.Clear
     On Error GoTo ErrorHandler
            If UBound(strRuleNames) >= LBound(strRuleNames) Then
                For ii = LBound(strRuleNames) To UBound(strRuleNames)
                    If strRuleNames(ii) = "ProductionRoutingXML Path" Then
                        Set oRuleObj = oRuleQuery.GetRule(strRuleNames(ii))
                        Exit For
                    End If
                Next ii
            End If
        
        If Not oRuleObj Is Nothing Then
           
            Set oRule = oRuleObj
            
            'Getting the Production Routing xml path
            strRuleProgId = oRule.ProgId
          
        End If
    End If
    'If there is no Production Routing xml path in the database then use the Production Routing xml path from default location
    If strRuleProgId = "" Then
        strXmlFilename = oContext.GetVariable("OLE_SERVER") & "\Production\ProductionRouting\ProductionRoutingData.xml"
    Else
        strXmlFilename = oContext.GetVariable("OLE_SERVER") & strRuleProgId
    End If
   
    lIndex = 0
    
    Set oPropColl = GetProductionRoutingProps(strXmlFilename, oPart)

    Set oAttributeMetaData = oSatgeProduction
    Set oAttrHelper = oSatgeProduction
    Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(oAttributeMetaData.IID("IJPlnProductionRouting"))

     If Not oInterfaceInfo Is Nothing Then
         On Error Resume Next
         Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
         Err.Clear
         On Error GoTo ErrorHandler
         If Not oAttributesCollection Is Nothing Then
             For Each oAttribute In oAttributesCollection
                 lIndex = lIndex + 1
                 If (lIndex <= oPropColl.Count) Then
                     oAttribute.Value = oPropColl.Item(lIndex)
                 End If
             Next
         End If
     End If
   
    Set oAttributeMetaData = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing

Exit Sub
ErrorHandler:
     MsgBox Err.Description
End Sub
Public Sub ComputeActionData(oSatgeProduction As IJDProductionRouting, oPart As Object, oAction As Object)
On Error GoTo ErrorHandler
Const METHOD = "ComputeActionData"
    
    Dim oAttributeMetaData          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim oActionColl                 As Collection
    Dim lIndex                      As Long
    Dim oRoutingActionColl          As IJElements

    Dim strXmlFilename              As String
    Dim oContext                    As IJContext
    Dim strAction                   As String
     Dim oSRDQuery                   As IJSRDQuery
    Dim oRuleQuery                  As IJSRDRuleQuery
    Dim oRule                       As IJSRDRule
    Dim oRuleObj                    As Object
    Dim strRuleNames()              As String
    Dim strRuleProgId               As String
    Dim ii                          As Integer
    
    Set oContext = GetJContext()
    
    strRuleProgId = ""
    
    Set oSRDQuery = New SRDQuery
    Set oRuleQuery = oSRDQuery.GetRulesQuery
    
    'Getting the Production Routing xml path from the catalog database
    If Not oRuleQuery Is Nothing Then
    On Error Resume Next
            strRuleNames = oRuleQuery.GetRuleNamesByType(200, 212, True)
            If Not UBound(strRuleNames) >= 0 Then
            strRuleNames = oRuleQuery.GetRuleNamesByType(400, 420, True)
            End If
     Err.Clear
     On Error GoTo ErrorHandler
            If UBound(strRuleNames) >= LBound(strRuleNames) Then
                For ii = LBound(strRuleNames) To UBound(strRuleNames)
                    If strRuleNames(ii) = "ProductionRoutingXML Path" Then
                        Set oRuleObj = oRuleQuery.GetRule(strRuleNames(ii))
                        Exit For
                    End If
                Next ii
            End If
        
        If Not oRuleObj Is Nothing Then
           
            Set oRule = oRuleObj
            
            'Getting the Production Routing xml path
            strRuleProgId = oRule.ProgId
          
        End If
    End If
    
    'If there is no Production Routing xml path in the database then use the Production Routing xml path from default location
    If strRuleProgId = "" Then
        strXmlFilename = oContext.GetVariable("OLE_SERVER") & "\Production\ProductionRouting\ProductionRoutingData.xml"
    Else
        strXmlFilename = oContext.GetVariable("OLE_SERVER") & strRuleProgId
    End If
    
    Set oAttributeMetaData = oAction
    Set oAttrHelper = oAction
    Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(oAttributeMetaData.IID("IJPlnRoutingAction"))
        
    If Not oInterfaceInfo Is Nothing Then
        On Error Resume Next
        Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
        Err.Clear
        On Error GoTo ErrorHandler
        If Not oAttributesCollection Is Nothing Then
            Set oActionColl = GetRoutingActions(strXmlFilename, oPart, oAttributesCollection.Item(1).Value)

            For Each oAttribute In oAttributesCollection
                lIndex = lIndex + 1
                If (lIndex <= oActionColl.Count) Then
                    If (oActionColl.Item(lIndex) <> vbNullString) Then
                        oAttribute.Value = oActionColl.Item(lIndex)
                    End If
                End If
            Next
        End If
    End If

   
    Set oAttributeMetaData = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing

Exit Sub
ErrorHandler:
     MsgBox Err.Description
End Sub

Public Function GetRoutingActions(strXMLFile As String, oPart As Object, strActionName As String) As Collection
Const METHOD = "GetRoutingActions"
On Error GoTo ErrorHandler

    Set GetRoutingActions = New Collection
    
    If strXMLFile <> vbNullString Then
        Dim oNodeList As IXMLDOMNodeList
        Dim oXMLDOC As New MSXML2.DOMDocument60
        Dim oXMLNode As IXMLDOMNode
        Dim oXMLNodeList As IXMLDOMNodeList
        Dim oChildXMLNode As IXMLDOMNode
        Dim oChildXMLNodes As IXMLDOMNodeList
        Dim oSubChildXMLNode As IXMLDOMNode
        Dim oSubChildXMLNodes As IXMLDOMNodeList
        Dim oSCAttribNode As IXMLDOMAttribute
        Dim oSAAttribNode As IXMLDOMAttribute
        Dim strXMLCSName As String
        

        'Load the file into XMLDOM
        oXMLDOC.Load (strXMLFile)
        Set oXMLNodeList = oXMLDOC.getElementsByTagName("RoutingAction")

        For Each oXMLNode In oXMLNodeList
            Dim strAction As String
            For Each oSAAttribNode In oXMLNode.Attributes
                 If ((UCase(oSAAttribNode.nodeName) = UCase("Name")) And (UCase(oSAAttribNode.nodeValue) = UCase(strActionName))) Then
                    GetRoutingActions.Add oSAAttribNode.nodeValue
                    strAction = oSAAttribNode.nodeValue
                    Exit For
                End If
            Next oSAAttribNode
            If UCase(strAction) = UCase(strActionName) Then

                For Each oChildXMLNode In oXMLNode.childNodes
                    If UCase(oChildXMLNode.nodeName) = UCase("RoutingMachine") Then
                        
                        For Each oSAAttribNode In oChildXMLNode.Attributes
                            Dim strMachine As String
                            Dim strCode As String
                            If (UCase(oSAAttribNode.nodeName) = UCase("Name")) Or (UCase(oSAAttribNode.nodeName) = UCase("Default")) Then
                                If UCase(strAction) = UCase("Bending") Then
                                    Call GetBendingMachineAndCode(oPart, strMachine, strCode)
                                    
                                        If strMachine = vbNullString Then
                                            Call GetDefBendingMachineAndCodeByType(oXMLNode, oPart, strMachine, strCode)
                                        End If
                                        
                                    If (UCase(oSAAttribNode.nodeName) = UCase("Name")) Then
                                        If strMachine <> vbNullString Then
                                            GetRoutingActions.Add strMachine
                                        Else
                                            GetRoutingActions.Add oSAAttribNode.nodeValue
                                        End If
                                    End If
                                    
                                    If (UCase(oSAAttribNode.nodeName) = UCase("Default")) Then
                                        If strCode <> vbNullString Then
                                            GetRoutingActions.Add strCode
                                        Else
                                            GetRoutingActions.Add oSAAttribNode.nodeValue
                                        End If
                                    End If
                                ElseIf UCase(strAction) = UCase("Priming") Then
                                    Call GetPrimingMachineAndCode(oXMLNode, oPart, strMachine, strCode)
                                    If (UCase(oSAAttribNode.nodeName) = UCase("Name")) Then
                                        If strMachine <> vbNullString Then
                                            GetRoutingActions.Add strMachine
                                        Else
                                            GetRoutingActions.Add oSAAttribNode.nodeValue
                                        End If
                                    End If
                                    
                                    If (UCase(oSAAttribNode.nodeName) = UCase("Default")) Then
                                        If strCode <> vbNullString Then
                                            GetRoutingActions.Add strCode
                                        Else
                                            GetRoutingActions.Add oSAAttribNode.nodeValue
                                        End If
                                    End If
                                Else
                                    GetRoutingActions.Add oSAAttribNode.nodeValue
                                End If
                            End If
                        Next oSAAttribNode
                        
                        Exit For
                        
                    End If
    
                Next oChildXMLNode
            End If
            
            If UCase(strActionName) = UCase(strAction) Then Exit For
            
        Next oXMLNode
    End If
    
    

Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Public Function GetProductionRoutingProps(strXmlFilename As String, oPart As Object) As Collection
Const METHOD = "GetProductionRoutingProps"
On Error GoTo ErrorHandler
        
        
    Set GetProductionRoutingProps = New Collection
    
    Dim oWorkCenterColl As Collection
    
    Set oWorkCenterColl = New Collection
    
    Set oWorkCenterColl = GetWorkCenterColl()
                
    GetProductionRoutingProps.Add ""
    GetProductionRoutingProps.Add oWorkCenterColl.Item(1)
    GetProductionRoutingProps.Add GetStageCode(strXmlFilename, oPart)
    
Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function
'Gets the list of workcenters available in catalog
Public Function GetWorkCenterColl() As Collection
Const METHOD = "GetWorkCenterColl"
On Error GoTo ErrorHandler

    Dim objSRDQuery                 As IJSRDQuery
    Dim objWorkcenterQuery          As IJSRDWorkcenterQuery
    Dim oWorkCenterColl             As New Collection
    Dim i                           As Long
    Dim strListOfWorkcenters()              As String
    ' Create the non-persistent SRDServices.SRDQuery to get the Query Object for Workcenter
    Set objSRDQuery = New SRDQuery
    
    ' Fetch the SRDWorkcenterQuery (persisted in the catalog db) to perform SQL queries on the catalog database
    Set objWorkcenterQuery = objSRDQuery.GetWorkcenterQuery()
    
    If Not objWorkcenterQuery Is Nothing Then
        objWorkcenterQuery.GetAllWorkcenters strListOfWorkcenters
    
        For i = LBound(strListOfWorkcenters) To UBound(strListOfWorkcenters)
            oWorkCenterColl.Add strListOfWorkcenters(i)
        Next i
        
        oWorkCenterColl.Add "<None>"

    End If
    
    Set GetWorkCenterColl = oWorkCenterColl
    
    Set objSRDQuery = Nothing
    Set objWorkcenterQuery = Nothing
    Set oWorkCenterColl = Nothing
    
Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Function IsStageCodeSet(oProductionRouting As Object) As Boolean
Const METHOD = "IsStageCodeSet"
On Error GoTo ErrorHandler
     Dim oAttributeMetaData          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim lIndex As Long
    
    IsStageCodeSet = False
    
    Set oAttributeMetaData = oProductionRouting
    Set oAttrHelper = oProductionRouting
    Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(oAttributeMetaData.IID("IJPlnProductionRouting"))

     If Not oInterfaceInfo Is Nothing Then
         On Error Resume Next
         Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
         Err.Clear
         On Error GoTo ErrorHandler
         If Not oAttributesCollection Is Nothing Then
             For Each oAttribute In oAttributesCollection
                 If oAttribute.AttributeInfo.Name = "StageCode" Then
                    
                    If Not IsEmpty(oAttribute.Value) Then
                        IsStageCodeSet = True
                    End If
                    
                 End If
             Next
         End If
     End If
        
    
    Set oAttributeMetaData = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing


Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function
Private Function GetNoOfActions(strXMLFile As String) As Long
Const METHOD = "GetNoOfActions"
On Error GoTo ErrorHandler
  
    Dim oXMLDOC As New MSXML2.DOMDocument60
    Dim oXMLNodeList As IXMLDOMNodeList

    'Load the file into XMLDOM
    oXMLDOC.Load (strXMLFile)
    
    Set oXMLNodeList = oXMLDOC.getElementsByTagName("RoutingAction")
    GetNoOfActions = oXMLNodeList.length
    
Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Sub GetRuleControl(strXMLFile As String, strRuleName As String)
 Const METHOD = "GetRuleControl"
    On Error GoTo ErrorHandler

    
    If strXMLFile <> vbNullString Then
        Dim oNodeList As IXMLDOMNodeList
        Dim oXMLDOC As New MSXML2.DOMDocument60
        Dim oXMLNode As IXMLDOMNode
        Dim oXMLNodeList As IXMLDOMNodeList
        Dim oChildXMLNode As IXMLDOMNode
        Dim oSAAttribNode As IXMLDOMAttribute

        'Load the file into XMLDOM
        oXMLDOC.Load (strXMLFile)
        Set oXMLNodeList = oXMLDOC.getElementsByTagName("RuleControl")

        Dim bNodefound As Boolean
        
        For Each oXMLNode In oXMLNodeList
            For Each oSAAttribNode In oXMLNode.Attributes
                If UCase(oSAAttribNode.nodeName) = UCase("Name") Then
                    If UCase(strRuleName) = UCase(oSAAttribNode.nodeValue) Then
                        bNodefound = True
                        Exit For
                    End If
                End If
            Next oSAAttribNode
            
            If bNodefound Then
                Exit For
            End If
        Next oXMLNode

        ReDim m_oRuleControl(0 To oXMLNode.childNodes.length - 1)
        
        For Each oChildXMLNode In oXMLNode.childNodes

                m_oRuleControl(CLng(oChildXMLNode.Attributes(2).nodeValue)).Type = oChildXMLNode.Attributes(0).nodeValue
                m_oRuleControl(CLng(oChildXMLNode.Attributes(2).nodeValue)).Action = oChildXMLNode.Attributes(1).nodeValue
                m_oRuleControl(CLng(oChildXMLNode.Attributes(2).nodeValue)).Order = CLng(oChildXMLNode.Attributes(2).nodeValue)
                
        Next oChildXMLNode

       
    End If

              
    Exit Sub
    
ErrorHandler:
     MsgBox Err.Description
End Sub
