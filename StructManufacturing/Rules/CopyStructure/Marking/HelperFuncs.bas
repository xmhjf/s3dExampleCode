Attribute VB_Name = "HelperFuncs"
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project: MfgCSOptions
' Module: HelperFuncs
'
' Description:  Helper functions to create different APS marks
'
' Author: Manasa.J
'
' Comments:
' User         Date        Comments
' Manasa.J     9.12.09     Initial creation
'*******************************************************************************
Option Explicit
Private MODULE As String
Private Const IID_IJMfgMarkingLines_AE = "{666DFFE4-015C-4F3F-8087-A0DA1C60A64D}"

Public Function CreateMarkingLine(oPart As Object, oGeom As Object, oRelatedObj As Object, oPhyConn As Object, lMarkingSide As Long, lMarkingType As Long) As Object
    Const METHOD = "CreateMarkingLines"
    On Error GoTo ErrorHandler
    
    Dim oMfgMarkingLinesAE          As IJMfgMarkingLines_AE
    Dim oMfgMarkingLinesAEFactory   As MfgMarkingLines_AEFactory
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Set oMfgMarkingLinesAEFactory = New MfgMarkingLines_AEFactory
    Set oMfgMarkingLinesAE = oMfgMarkingLinesAEFactory.Create(oResourceManager)
    
    Dim oIJDObjectCollection As IJDObjectCollection
    Set oIJDObjectCollection = New JObjectCollection
    
    oMfgMarkingLinesAE.GenerateMfgMarkingLine oPart, oGeom, oIJDObjectCollection, Nothing, oRelatedObj
    
    oMfgMarkingLinesAE.MfgMarkingSide = lMarkingSide
    oMfgMarkingLinesAE.MfgMarkingType = lMarkingType
    
    Dim oRuleAttributes As Dictionary
    Set oRuleAttributes = New Dictionary
    Dim oMarkingBOColl As IJElements
    Set oMarkingBOColl = New JObjectCollection
    oMarkingBOColl.Add oMfgMarkingLinesAE
    
    Dim oCustomAttributes As IJElements
    Set oCustomAttributes = FillMarkingUserAttributes(oRuleAttributes, oMarkingBOColl, oPhyConn, oRelatedObj, lMarkingType)
    
    Set CreateMarkingLine = oMfgMarkingLinesAE
    
    Set oResourceManager = Nothing
    Set oMfgMarkingLinesAEFactory = Nothing
    
    Set oIJDObjectCollection = Nothing
    Set oRuleAttributes = Nothing
    Set oMarkingBOColl = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function FillMarkingUserAttributes(ByVal oColl As Dictionary, _
                                            ByVal pMarkingLines As IJElements, _
                                            ByVal pDispReferencePart As Object, _
                                            ByVal pDispRelatedPart As Object, _
                                            ByVal lMarkingType As Long) As IJElements
    Const METHOD = "FillMarkingUserAttributes"
    On Error GoTo ErrorHandler
    
    Dim oPOM   As IJDPOM
    Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oMoniker   As IUnknown
    Set oMoniker = oPOM.GetObjectMoniker(pMarkingLines.Item(1))
       
    If oPOM.SupportsInterface(oMoniker, "IJUAMfgSketchLocation") Then
        
        Dim i             As Integer
        Dim strAttrName   As String
        Dim varAttrValue  As Variant
    
        Dim oAttribute    As IJDAttribute
        Dim oReturnColl   As IJElements
        Set oReturnColl = New JObjectCollection
        
        Dim strColl As New Collection
        Set strColl = GetUserAttributes(pMarkingLines, lMarkingType)
        
        Dim j As Integer
        
        For j = 1 To strColl.Count
            strAttrName = strColl(j)
            
            'For i = 1 To oColl.Count
                 
                varAttrValue = vbNull
                
                On Error Resume Next
                varAttrValue = GetUserAttribute(pMarkingLines, strAttrName).Value
                On Error GoTo ErrorHandler
                
                If varAttrValue = vbNullString Then
                           
                    If oColl.Exists(strAttrName) Then
                        On Error Resume Next
                        varAttrValue = oColl.Item(strAttrName)
                        On Error GoTo ErrorHandler
                    Else
         
                        Select Case strAttrName
                
                            Case "RelatedPartName"
                                If Not pDispRelatedPart Is Nothing Then
                                    varAttrValue = GetRelatedPartName(pDispRelatedPart)
                                End If
                
                            Case "ConnectionGUID"
                                
                                If Not pDispReferencePart Is Nothing Then
                                    varAttrValue = GetPhysicalConnectionGuid(pDispReferencePart)
                                End If
                
                            Case "Direction"
                                If Not pDispRelatedPart Is Nothing Then
                                    varAttrValue = GetRelatedObjectDirection(pDispRelatedPart, pDispReferencePart)
                                End If
                
                            Case "FittingAngle"
                                If Not pDispReferencePart Is Nothing Then
                                    varAttrValue = GetPCFittingAngle(pDispReferencePart)
                                End If
                
                            Case "PartGUID"
                                If Not pDispRelatedPart Is Nothing Then
                                    varAttrValue = GetConnectedPartGuid(pDispRelatedPart)
                                End If
                
                            Case "MaxAssyMarginValue"
                                varAttrValue = 0
                
                            Case "MaxFabMarginValue"
                                'To Do: Implementation for System Default
                                varAttrValue = 0
                
                            Case "MaxCustomMarginValue"
                                'To Do: Implementation for System Default
                                varAttrValue = 0
                
                            Case "FlangeDirection"
                                'To Do: Implementation for System Default
                                'varAttrValue = "opposite"
                                 varAttrValue = GetFlangeDirection(pDispRelatedPart)
                
                        End Select
                    
                    End If
                End If
            
                Set oAttribute = SetUserAttribute(pMarkingLines, varAttrValue, strAttrName)
                
                If Not oAttribute Is Nothing Then
                    oReturnColl.Add oAttribute
                End If
                
                'Exit For
            
            'Next i
        
        Next j
         
        Set FillMarkingUserAttributes = oReturnColl
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetUserAttribute(ByVal pMarkingLines As IJElements, strPropertyName As String) As IJDAttribute
    Const METHOD = "GetUserAttribute"
    On Error GoTo ErrorHandler
    
    Dim oAttributes        As IJDAttributes
    Dim oAttrIID           As Variant
    Dim oAttribute         As IJDAttribute
    Dim oIJDAttributesCol  As IJDAttributesCol
    Dim oIJDInterfaceInfo  As IJDInterfaceInfo
    Dim oMetaDataHelp      As IJDAttributeMetaData
    
    Set oAttributes = pMarkingLines.Item(1)
    Set oMetaDataHelp = pMarkingLines.Item(1)
        
    For Each oAttrIID In oAttributes
        If oAttrIID = oMetaDataHelp.IID("IJUAMfgSketchLocation") Then
            Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(oAttrIID)
            Set oIJDAttributesCol = oAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)
            
            For Each oAttribute In oIJDAttributesCol
                If oAttribute.AttributeInfo.Name = strPropertyName Then
                    Set GetUserAttribute = oAttribute
                    Exit Function
                End If
            Next
            
        End If
    Next
     
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function SetUserAttribute(ByVal pMarkingLines As IJElements, varValue As Variant, strPropertyName As String) As IJDAttribute
    Const METHOD = "SetUserAttribute"
    On Error GoTo ErrorHandler

    Dim oAttribute   As IJDAttribute
    Set oAttribute = GetUserAttribute(pMarkingLines, strPropertyName)
    
    If Not oAttribute Is Nothing Then
        oAttribute.Value = varValue
    
        Set SetUserAttribute = oAttribute
        
        Dim oAssocHelper         As New AssocHelper
        Dim oMfgMarkingLinesAE   As IJMfgMarkingLines_AE
        
        Set oMfgMarkingLinesAE = pMarkingLines.Item(1)
        oAssocHelper.UpdateObject oMfgMarkingLinesAE, IID_IJMfgMarkingLines_AE
        
        Set oAssocHelper = Nothing
        Set oMfgMarkingLinesAE = Nothing
    End If

    Exit Function
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetUserAttributes(ByVal pMarkingLines As IJElements, ByVal lMarkingType As Long) As Collection
    Const METHOD = "GetUserAttributes"
    On Error GoTo ErrorHandler

    Set GetUserAttributes = New Collection

    Dim oAttributes        As IJDAttributes
    Dim oAttrIID           As Variant
    Dim oAttribute         As IJDAttribute
    Dim oIJDAttributesCol  As IJDAttributesCol
    Dim oIJDInterfaceInfo  As IJDInterfaceInfo
    Dim oMetaDataHelp      As IJDAttributeMetaData

    Set oAttributes = pMarkingLines.Item(1)
    Set oMetaDataHelp = pMarkingLines.Item(1)

    Dim strPropertyName As String

    For Each oAttrIID In oAttributes
        If oAttrIID = oMetaDataHelp.IID("IJUAMfgSketchLocation") Then
            Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(oAttrIID)
            Set oIJDAttributesCol = oAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)

            For Each oAttribute In oIJDAttributesCol
                strPropertyName = oAttribute.AttributeInfo.Name

                Select Case lMarkingType
                    Case STRMFG_PLATELOCATION_MARK, STRMFG_BRACKETLOCATION_MARK:
                        If strPropertyName = "RelatedPartName" Or _
                            strPropertyName = "Direction" Or _
                            strPropertyName = "FittingAngle" Or _
                            strPropertyName = "PartGUID" Or _
                            strPropertyName = "ConnectionGUID" Then

                                GetUserAttributes.Add strPropertyName

                        End If
                    Case STRMFG_COLLARPLATELOCATION_MARK
                        If strPropertyName = "RelatedPartName" Or _
                            strPropertyName = "Direction" Or _
                            strPropertyName = "PartGUID" Or _
                            strPropertyName = "ConnectionGUID" Then

                                GetUserAttributes.Add strPropertyName

                        End If
                    Case STRMFG_PROFILELOCATION_MARK
                        If strPropertyName = "RelatedPartName" Or _
                            strPropertyName = "Direction" Or _
                            strPropertyName = "FittingAngle" Or _
                            strPropertyName = "FlangeDirection" Or _
                            strPropertyName = "PartGUID" Or _
                            strPropertyName = "ConnectionGUID" Then
                        

                                GetUserAttributes.Add strPropertyName

                        End If
                    Case STRMFG_KNUCKLE_MARK
                        If strPropertyName = "RelatedPartName" Or _
                            strPropertyName = "Direction" Or _
                            strPropertyName = "FittingAngle" Or _
                            strPropertyName = "PartGUID" Or _
                            strPropertyName = "ConnectionGUID" Then

                                GetUserAttributes.Add strPropertyName

                        End If
                    Case STRMFG_DIRECTION
                        If strPropertyName = "Direction" Then
                        
                            GetUserAttributes.Add strPropertyName

                        End If

                    Case STRMFG_MARGIN_MARK
                        If strPropertyName = "RelatedPartName" Or _
                            strPropertyName = "PartGUID" Or _
                            strPropertyName = "MaxAssyMarginValue" Or _
                            strPropertyName = "MaxFabMarginValue" Or _
                            strPropertyName = "MaxCustomMarginValue" Then

                                GetUserAttributes.Add strPropertyName

                        End If
                        
                    Case STRMFG_END_MARK, STRMFG_LAP_MARK, STRMFG_SEAM_MARK:
                        If strPropertyName = "RelatedPartName" Or _
                            strPropertyName = "PartGUID" Or _
                            strPropertyName = "ConnectionGUID" Then
                        
                                GetUserAttributes.Add strPropertyName

                        End If
                    Case STRMFG_NAVALARCHLINE
                         If strPropertyName = "RelatedPartName" Then
                             GetUserAttributes.Add strPropertyName
                         End If
                End Select
            Next
        End If
    Next

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetRelatedPartName(oRelatedPart As Object) As Variant
    Const METHOD = "GetRelatedPartName"
    On Error GoTo ErrorHandler

    GetRelatedPartName = vbNullString
    
    If oRelatedPart Is Nothing Then Exit Function
    
    Dim oNamedItem   As IJNamedItem
    Set oNamedItem = oRelatedPart

    GetRelatedPartName = oNamedItem.Name
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetPhysicalConnectionGuid(oReferenceCurve As Object) As Variant
    Const METHOD = "GetPhysicalConnectionGuid"
    On Error GoTo ErrorHandler

    GetPhysicalConnectionGuid = vbNullString

    Dim oAppConnType          As IJAppConnectionType
    Dim oPhysicalConnection   As Object

    On Error Resume Next
    Set oAppConnType = oReferenceCurve

    If oAppConnType Is Nothing Then Exit Function

    If oAppConnType.Type = ConnectionLogical Then
        Dim oDesignparent   As IJDesignParent
        Set oDesignparent = oAppConnType
        
        Dim pcCollection    As New Collection

        GetPhysicalConnections oDesignparent, pcCollection
        
        If pcCollection.Count > 0 Then
            Set oPhysicalConnection = pcCollection.Item(1)
        End If
        
    ElseIf oAppConnType.Type = ConnectionPhysical Then
    
        Set oPhysicalConnection = oReferenceCurve
        
    Else
        Exit Function
    End If

    If Not oPhysicalConnection Is Nothing Then
        
        Dim oPOM As IJDPOM
        Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
        
        Dim sDbIdentifier    As String
        sDbIdentifier = oPOM.DbIdentifierFromMoniker(oPOM.GetObjectMoniker(oPhysicalConnection))

        GetPhysicalConnectionGuid = sDbIdentifier
    End If
 
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetRelatedObjectDirection(oRelatedObj As Object, Optional oRefInput As Object) As Variant
    Const METHOD = "GetRelatedObjectDirection"
    On Error GoTo ErrorHandler

    GetRelatedObjectDirection = "unknown"
    
    If (TypeOf oRelatedObj Is IJPlatePart) Then
        If (TypeOf oRelatedObj Is IJCollarPart) Then
            GetRelatedObjectDirection = ""
            Exit Function
        End If
        
        If (TypeOf oRelatedObj Is IJSmartPlate) Then
           GetRelatedObjectDirection = GetBracketPlateThicknessDir(oRelatedObj, oRefInput)
           Exit Function
        End If
        
        Dim oPlate As IJPlate
        Set oPlate = GetParentPlateSystem(oRelatedObj)
        GetRelatedObjectDirection = GetPlateThicknessDirection(oPlate)
        
    ElseIf (TypeOf oRelatedObj Is IJPlateSystem) Then
        Set oPlate = oRelatedObj
        GetRelatedObjectDirection = GetPlateThicknessDirection(oPlate)
        
    ElseIf (TypeOf oRelatedObj Is IJStiffenerPart) Then
        GetRelatedObjectDirection = GetStiffenerDirection(oRelatedObj, oRefInput)
    Else
        'ToDo: Implementation for other object types
        Exit Function
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetPCFittingAngle(oReferenceCurve As Object) As Variant
    Const METHOD = "GetPCFittingAngle"
    On Error GoTo ErrorHandler

    Dim oAppConnType         As IJAppConnectionType
    Dim oPhysicalConnection  As StructDetailObjects.PhysicalConn
    Set oPhysicalConnection = New StructDetailObjects.PhysicalConn

    GetPCFittingAngle = 0

    On Error Resume Next
    Set oAppConnType = oReferenceCurve

    If oAppConnType Is Nothing Then Exit Function

    If oAppConnType.Type = ConnectionLogical Then
        Dim oDesignparent As IJDesignParent
        Set oDesignparent = oAppConnType
        Dim pcCollection As New Collection

        GetPhysicalConnections oDesignparent, pcCollection
        
        If pcCollection.Count > 0 Then
            Set oPhysicalConnection.object = pcCollection.Item(1)
            GetPCFittingAngle = oPhysicalConnection.TeeMountingAngle
        End If
        
    ElseIf oAppConnType.Type = ConnectionPhysical Then
        Set oPhysicalConnection.object = oReferenceCurve
        GetPCFittingAngle = oPhysicalConnection.TeeMountingAngle
    Else
        Exit Function
    End If

    

    Exit Function
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetConnectedPartGuid(oRelatedObject As Object) As Variant
    Const METHOD = "GetConnectedPartGuid"
    On Error GoTo ErrorHandler

    GetConnectedPartGuid = vbNullString

    If oRelatedObject Is Nothing Then Exit Function
    
    Dim oPOM   As IJDPOM
    Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim sDbIdentifier As String
    sDbIdentifier = oPOM.DbIdentifierFromMoniker(oPOM.GetObjectMoniker(oRelatedObject))

    GetConnectedPartGuid = sDbIdentifier

    Exit Function
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetPhysicalConnections(oLCObject As IJDesignParent, ByRef oPCCollection As Collection)
    Const METHOD = "GetPhysicalConnections"
    On Error GoTo ErrorHandler
    
    Dim lCChildren As IJDObjectCollection
    oLCObject.GetChildren lCChildren
    
    Dim lcChildObject As Object
    
    For Each lcChildObject In lCChildren
        If TypeName(lcChildObject) = "IJStructPhysicalConnection" Then
            oPCCollection.Add lcChildObject
        Else
            Dim oDesignparent As IJDesignParent
            
            On Error Resume Next
            Set oDesignparent = lcChildObject
            
            If Not oDesignparent Is Nothing Then
                GetPhysicalConnections lcChildObject, oPCCollection
            End If
        End If
    Next
    
    Set lcChildObject = Nothing
    Set lCChildren = Nothing

    Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetFlangeDirection(pDispRelatedPart As Object) As String
    Const METHOD = "GetFlangeDirection"
    On Error GoTo ErrorHandler
    
    GetFlangeDirection = "opposite"
    
    If Not TypeOf pDispRelatedPart Is IJStiffenerPart Then Exit Function
    
    Dim oSDProfileWrapper   As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = pDispRelatedPart
    
    Dim oPartSupport        As GSCADSDPartSupport.IJPartSupport
    Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
    Dim eTSide              As GSCADSDPartSupport.ThicknessSide
    
    Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set oPartSupport.Part = oSDProfileWrapper.object
    Set oProfilePartSupport = oPartSupport
    
    eTSide = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint
    
    Dim bIsOpposite As Boolean
    bIsOpposite = False
    
    If eTSide = SideB Then
        bIsOpposite = True
    ElseIf eTSide = SideA Then
        bIsOpposite = False
    End If
        
    If bIsOpposite Then
        GetFlangeDirection = "opposite"
    Else
        GetFlangeDirection = "same"
    End If
    
    Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetBracketPlateThicknessDir(oBracketPlate As Object, oRefInput As Object) As String
    Const METHOD = "GetBracketPlateThicknessDir"
    
    GetBracketPlateThicknessDir = "unknown"
    
    Dim oSDPhysicalConn   As StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn.object = oRefInput
    
    Dim oPlatePart        As IJPlatePart
    If oSDPhysicalConn.ConnectedObject1 Is oBracketPlate Then
        Set oPlatePart = oSDPhysicalConn.ConnectedObject2
    Else
        Set oPlatePart = oSDPhysicalConn.ConnectedObject1
    End If
    
    Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oPlatePart
    
    Dim oPartInfo   As IJDPartInfo
    Set oPartInfo = New PartInfo
    
    Dim eSideOfConnectedObjectToBeMarked As ThicknessSide
    Dim eMoldedDir                       As StructMoldedDirection
    eMoldedDir = oPartInfo.GetPlatePartThicknessDirection(oBracketPlate)
    
    Dim oSDConPlateWrapper   As New StructDetailObjects.PlatePart
    Set oSDConPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDConPlateWrapper.object = oBracketPlate
    
    Dim sMoldedSide As String
    sMoldedSide = oSDConPlateWrapper.MoldedSide
    
    If eMoldedDir = Centered Then
        eSideOfConnectedObjectToBeMarked = SideUnspecified
    Else
        If sMoldedSide = "Base" Then
            eSideOfConnectedObjectToBeMarked = SideA
        ElseIf sMoldedSide = "Offset" Then
            eSideOfConnectedObjectToBeMarked = SideB
        End If
    End If
        
    Dim bContourTee  As Boolean
    Dim oVector      As IJDVector
    Dim oWB          As IJWireBody, oTeeWire As IJWireBody
    
    bContourTee = oSDPlateWrapper.Connection_ContourTee(oRefInput, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)
    
    Dim oMfgRuleHelper   As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    GetBracketPlateThicknessDir = oMfgRuleHelper.GetDirection(oVector)
    
    On Error GoTo ErrorHandler
    Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetPlateThicknessDirection(oPlate As Object) As String
    Const METHOD = "GetPlateThicknessDirection"
    On Error GoTo ErrorHandler
    
    Dim oPlateMC   As IJDPlateMoldedConventions
    Set oPlateMC = oPlate
    
    Dim Dir        As StructMoldedDirection
    Dir = oPlateMC.plateThicknessDirection

    Dim ePlateType As StructPlateType
    Dim strDir     As String
    
    ePlateType = oPlate.plateType
    
    Select Case ePlateType
        Case DeckPlate:
            Select Case Dir
                Case AboveDir:
                    strDir = "upper"
                Case BelowDir:
                    strDir = "lower"
            End Select
        
        Case LBulkheadPlate
            Select Case Dir
                Case StarDir:
                    strDir = "starboard"
                Case PortDir:
                    strDir = "port"
                Case InboardDir:
                    strDir = "in"
                Case OutboardDir:
                    strDir = "out"
                Case Centered:
                    strDir = "centered"
            End Select
        Case TBulkheadPlate
            Select Case Dir
                Case ForeDir:
                    strDir = "fore"
                Case AftDir:
                    strDir = "aft"
            End Select
        Case TransverseTube, VerticalTube, LongitudinalTube
            Select Case Dir
                Case InDir:
                    strDir = "in"
                Case OutDir:
                    strDir = "out"
                Case Centered:
                    strDir = "centered"
            End Select
    End Select
    
    GetPlateThicknessDirection = strDir
  Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetStiffenerDirection(oRelatedObject As Object, oRefInput As Object) As String
    Const METHOD = "GetStiffenerDirection"
    On Error GoTo ErrorHandler
    
    GetStiffenerDirection = "unknown"
    
    Dim oSDPhysicalConn   As StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn.object = oRefInput
    
    Dim oPlatePart     As IJPlatePart
    If oSDPhysicalConn.ConnectedObject1Type = SDOBJECT_PLATE Then
        Set oPlatePart = oSDPhysicalConn.ConnectedObject1
    Else
        Set oPlatePart = oSDPhysicalConn.ConnectedObject2
    End If
    
    Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oPlatePart
       
    Dim oMfgRuleHelper   As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    Dim oSDProfileWrapper   As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = oRelatedObject
    
    Dim oPartSupport        As GSCADSDPartSupport.IJPartSupport
    Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
    
    Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set oPartSupport.Part = oSDProfileWrapper.object
    Set oProfilePartSupport = oPartSupport
    
    Dim eSideOfConnectedObjectToBeMarked As ThicknessSide
    eSideOfConnectedObjectToBeMarked = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint
        
    Dim bContourTee  As Boolean
    Dim oVector      As IJDVector
    Dim oWB          As IJWireBody, oTeeWire As IJWireBody
    
    bContourTee = oSDPlateWrapper.Connection_ContourTee(oRefInput, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)
        
    Dim oCS          As IJComplexString
    Dim oVector1     As IJDVector
    Dim oStart       As IJDPosition, oEnd As IJDPosition
    Dim oPlateNormal As IJDVector
    Dim oSurfaceBody As IJSurfaceBody
    
    If bContourTee = True Then
        'Bound the wire based on split points, if there are any.
        Set oWB = TrimTeeWireForSplitPCs(oRefInput, oTeeWire)
        Set oCS = oMfgRuleHelper.WireBodyToComplexString(oWB)
        
        oWB.GetEndPoints oStart, oEnd
        Set oSurfaceBody = oSDPlateWrapper.BasePort(BPT_Base).Geometry
        Set oStart = oMfgRuleHelper.ProjectPointOnSurface(oStart, oSurfaceBody, oVector1)
        oSurfaceBody.GetNormalFromPosition oStart, oPlateNormal
        Set oSurfaceBody = Nothing
        Set oVector1 = Nothing
        Set oStart = Nothing
        Set oEnd = Nothing
            
        If (Abs(oPlateNormal.x) > Abs(oPlateNormal.y) And Abs(oPlateNormal.x) > Abs(oPlateNormal.z)) Then
            oVector.x = 0
        ElseIf Abs(oPlateNormal.y) > Abs(oPlateNormal.z) Then
            oVector.y = 0
        Else
            oVector.z = 0
        End If
                
        'Determine whether the Plate part is crossing the centre line
        Dim bCenterPart     As Boolean
        Dim oHelperSupport  As IJMfgRuleHelpersSupport
        
        Set oHelperSupport = New MfgRuleHelpersSupport
        bCenterPart = False
        
        If Not oHelperSupport Is Nothing Then
            Dim dMinX As Double, dMinY As Double, dMinZ As Double
            Dim dMaxX As Double, dMaxY As Double, dMaxZ As Double
            
            'Get the Part Range
            oHelperSupport.GetRange oPlatePart, dMinX, dMinY, dMinZ, dMaxX, dMaxY, dMaxZ
            'Check for Y value to determine whether part is crossing the Center line (y=0)
            If dMinY < 0 And dMaxY > 0 Then
                bCenterPart = True
            End If
        End If
                
        GetStiffenerDirection = oMfgRuleHelper.GetThicknessDirection(oCS, oVector, bCenterPart)
        
    End If
          
Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

' This method gets the parent plate system recursively
Private Function GetParentPlateSystem(oPlatePart As Object) As Object

    Const METHOD = "GetParentPlateSystem"
    On Error GoTo ErrHandler

    Dim oChild   As IJSystemChild
    Dim oPlate   As IJPlate

    Set oChild = oPlatePart
    
    Set oPlate = oChild.GetParent()
    If TypeOf oPlate Is IJPlate Then
        If TypeOf oPlate Is IJPlateSystem Then
            Set GetParentPlateSystem = oPlate
        Else
            Set GetParentPlateSystem = GetParentPlateSystem(oPlate)
        End If
    Else
        GoTo ErrHandler
    End If

CleanUp:
    Set oPlate = Nothing
    Set oChild = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp

End Function
Public Function CreateMarksOfGivenType(ByVal oInputPlates As IJElements, ByVal oNewPlates As IJElements, eMarkingType As StrMfgGeometryType) As IJElements
    Const METHOD = "CreateMarksOfGivenType"
    On Error GoTo ErrorHandler
    
    Set CreateMarksOfGivenType = New JObjectCollection
    
    Dim lOldPlatesCount As Long
    Dim lNewPlatesCount As Long
    
    Dim oSeamsGeom      As IJElements
    Set oSeamsGeom = New JObjectCollection
    
    Dim oSeamsColl      As IJDObjectCollection
    Set oSeamsColl = New JObjectCollection
    
    Dim oMfgMGHelper        As IJMfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    For lOldPlatesCount = 1 To oInputPlates.Count
            Dim oPartSupport   As IJPartSupport
            Set oPartSupport = New PartSupport
            Set oPartSupport.Part = oInputPlates.Item(lOldPlatesCount)
        
            Dim ConnectedObjColl  As Collection
            Dim ConnectionsColl   As Collection
        
            Dim ThisPortColl  As Collection
            Dim OtherPortColl As Collection
            oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                             ConnectedObjColl, ConnectionsColl, _
                                             ThisPortColl, OtherPortColl
            
            Dim oConnObjColl      As Collection
            Dim oConnections      As Collection
            Dim oSidesColl        As Collection
            Dim oPhyConnsForMarks As Collection
            
            Set oConnObjColl = New Collection
            Set oConnections = New Collection
            Set oSidesColl = New Collection
            Set oPhyConnsForMarks = New Collection
            
            Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
            Set oSDPlateWrapper = New StructDetailObjects.PlatePart
            Set oSDPlateWrapper.object = oInputPlates.Item(lOldPlatesCount)
        
            Dim i As Integer
            
            Select Case eMarkingType
                Case STRMFG_PLATELOCATION_MARK
                   
                   For i = 1 To ConnectedObjColl.Count
                        'Need to filter out bracket marks, collar marks and lap connection marks
                        ' As both Bracket and Collar are implementing IJPlatePart we need to check if the
                        ' connected item is not either of those
                        If (TypeOf ConnectedObjColl.Item(i) Is IJSmartPlate) Or (TypeOf ConnectedObjColl.Item(i) Is IJCollarPart) Then
                            GoTo NextItemForPlateMark
                        End If
                        
                        ' Check if the type of connection is PARTSUPPORT_CONNTYPE_TEE
                        Dim bIsCrossOfTee As Boolean
                        Dim oConnType As ContourConnectionType
                        
                        oPartSupport.GetConnectionTypeForContour ConnectionsColl.Item(i), _
                                                                 oConnType, _
                                                                 bIsCrossOfTee
                                                       
                        If TypeOf ConnectedObjColl.Item(i) Is IJPlatePart And oConnType = PARTSUPPORT_CONNTYPE_TEE And bIsCrossOfTee Then
                            Dim bContourLap   As Boolean
                            Dim oWBColl       As Collection
                            bContourLap = oSDPlateWrapper.Connection_ContourLap(ConnectionsColl.Item(i), oWBColl)
                    
                            If bContourLap = True Then
                                GoTo NextItemForPlateMark
                            End If
                            
                            If IsPartOfSplitOperation(ConnectionsColl.Item(i)) Then
                                GoTo NextItemForPlateMark
                            End If
                            
                            Dim oGeom   As IJComplexString
                            Set oGeom = CreatePersistentComplexString(GetConnectionContourTee(ConnectionsColl.Item(i), ConnectedObjColl.Item(i), oInputPlates.Item(lOldPlatesCount)))
                            If Not oGeom Is Nothing Then
                                oSidesColl.Add ThisPortColl.Item(i)
                                oConnObjColl.Add ConnectedObjColl.Item(i)
                                oConnections.Add oGeom
                                oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                            End If
                            
                        End If
NextItemForPlateMark:
                    Next i
                    
               Case STRMFG_BRACKETLOCATION_MARK
                    
                     For i = 1 To ConnectedObjColl.Count
                        If TypeOf ConnectedObjColl.Item(i) Is IJSmartPlate Then
                            
                            Dim oSmartPlate        As IJSmartPlate
                            Dim eSmartPlateTypes   As SmartPlateTypes
                            
                            Set oSmartPlate = ConnectedObjColl.Item(i)
                            eSmartPlateTypes = oSmartPlate.SmartPlateType
                            
                            If Not eSmartPlateTypes = spType_BRACKET Then
                                 GoTo NextItemForBracket
                            End If
                                                         
                            If IsPartOfSplitOperation(ConnectionsColl.Item(i)) Then
                                GoTo NextItemForPlateMark
                            End If
                            
                            Dim ComplexString   As IJComplexString
                            Set ComplexString = CreatePersistentComplexString(GetConnectionContourTee(ConnectionsColl.Item(i), ConnectedObjColl.Item(i), oInputPlates.Item(lOldPlatesCount)))
                            
                            If Not ComplexString Is Nothing Then
                                oSidesColl.Add ThisPortColl.Item(i)
                                oConnObjColl.Add ConnectedObjColl.Item(i)
                                oConnections.Add ComplexString
                                oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                            End If
                            
                        End If
NextItemForBracket:
                     Next i
                
                Case STRMFG_COLLARPLATELOCATION_MARK
                     For i = 1 To ConnectedObjColl.Count
                        If TypeOf ConnectedObjColl.Item(i) Is IJCollarPart Then
                            Dim oComplexStrings   As IJElements
                            Set oComplexStrings = GetConnectionContourLap(ConnectionsColl.Item(i), oInputPlates.Item(lOldPlatesCount))

                            Dim oComplexString   As IJComplexString
                            
                            For Each oComplexString In oComplexStrings
                                If Not oComplexString Is Nothing Then
                                    oConnections.Add oComplexString
                                    oSidesColl.Add ThisPortColl.Item(i)
                                    oConnObjColl.Add ConnectedObjColl.Item(i)
                                    oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                                End If
                            Next
                        End If
                     Next i
                     
                Case STRMFG_LAP_MARK
                     For i = 1 To ConnectedObjColl.Count
                        If (TypeOf ConnectedObjColl.Item(i) Is IJSmartPlate) Or (TypeOf ConnectedObjColl.Item(i) Is IJCollarPart) Then
                             GoTo NextItemForLap
                        End If
                        
                        If TypeOf ConnectedObjColl.Item(i) Is IJPlatePart Then
                            Dim oCmplxStrings   As IJElements
                            Set oCmplxStrings = GetConnectionContourLap(ConnectionsColl.Item(i), oInputPlates.Item(lOldPlatesCount))

                            Dim oCS   As IJComplexString
                            
                            For Each oCS In oCmplxStrings
                                If Not oCS Is Nothing Then
                                    oConnections.Add oCS
                                    oSidesColl.Add ThisPortColl.Item(i)
                                    oConnObjColl.Add ConnectedObjColl.Item(i)
                                    oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                                End If
                            Next
                        End If
NextItemForLap:
                     Next i
                     
                Case STRMFG_PROFILELOCATION_MARK
                    For i = 1 To ConnectedObjColl.Count
                        If TypeOf ConnectedObjColl.Item(i) Is IJProfilePart Then
                            'Filter out the end connection marks
                             Dim bContourEnd    As Boolean
                             Dim oWireBodyColl  As Collection
                             bContourEnd = oSDPlateWrapper.Connection_ContourProfileEnd(ConnectionsColl.Item(i), oWireBodyColl)
                             
                             If bContourEnd = True Then
                                 GoTo NextItemForProfie
                             End If
                                                           
                            If IsPartOfSplitOperation(ConnectionsColl.Item(i)) Then
                                GoTo NextItemForPlateMark
                            End If
                             
                            Dim oMarkGeom   As IJComplexString
                            Set oMarkGeom = CreatePersistentComplexString(GetConnectionContourTee(ConnectionsColl.Item(i), ConnectedObjColl.Item(i), oInputPlates.Item(lOldPlatesCount)))
                            If Not oMarkGeom Is Nothing Then
                                oSidesColl.Add ThisPortColl.Item(i)
                                oConnObjColl.Add ConnectedObjColl.Item(i)
                                oConnections.Add oMarkGeom
                                oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                            End If
                             
                        End If
NextItemForProfie:
                     Next i
                
                Case STRMFG_END_MARK
                    For i = 1 To ConnectedObjColl.Count
                        If TypeOf ConnectedObjColl.Item(i) Is IJProfilePart Then
                            
                            Dim bEndConn             As Boolean
                            Dim oWireBodyCollection  As Collection
                            bEndConn = oSDPlateWrapper.Connection_ContourProfileEnd(ConnectionsColl.Item(i), oWireBodyCollection)
                             
                            If (bEndConn = True) And Not (oWireBodyCollection Is Nothing) Then
                                'Merge the outer contour wirebody collection into one wirebody
                                Dim oProfileEndContour  As IJWireBody
                                
                                oMfgMGHelper.MergeWireBodyCollection oWireBodyCollection, oProfileEndContour
                                
                                Dim oMfgRuleHelper   As New MfgRuleHelpers.Helper
                                Dim oComplexStr      As IJComplexString
                                Set oComplexStr = oMfgRuleHelper.WireBodyToComplexString(oProfileEndContour)
                                 
                                If Not oComplexStr Is Nothing Then
                                    oSidesColl.Add ThisPortColl.Item(i)
                                    oConnObjColl.Add ConnectedObjColl.Item(i)
                                    oConnections.Add CreatePersistentComplexString(oComplexStr)
                                    oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                                End If
                                
                            End If
                             
                        End If
                    Next i
                    
                Case STRMFG_SIGHTLINE_MARK, STRMFG_SEAM_MARK:
                             
                     For i = 1 To ConnectedObjColl.Count
                        If TypeOf ConnectedObjColl.Item(i) Is IJPlatePart Then
                                                    
                            Dim oThisStructPort  As IJStructPort, oOtherStructPort As IJStructPort
                            Dim ThisPortFlag     As eUSER_CTX_FLAGS
                            Dim OtherPortFlag    As eUSER_CTX_FLAGS
                            
                            Set oThisStructPort = ThisPortColl.Item(i)
                            Set oOtherStructPort = OtherPortColl.Item(i)
                            
                            ThisPortFlag = oThisStructPort.ContextID
                            OtherPortFlag = oOtherStructPort.ContextID
                            
                            Dim iThisPort As Integer
                            Dim iOtherPort As Integer
                            
                            iThisPort = ThisPortFlag And CTX_LATERAL
                            iOtherPort = OtherPortFlag And CTX_LATERAL
                            
                            If iThisPort > 0 And iOtherPort > 0 Then
                                Dim oStructDetailHelper   As IJStructDetailHelper
                                Set oStructDetailHelper = New StructDetailHelper
                                
                                Dim oOperator   As Object
                                Dim oOperation  As IJStructOperation
            
                                'The operator has to be a seam
                                oStructDetailHelper.FindOperatorForOperationInGraphByID oInputPlates.Item(lOldPlatesCount), oThisStructPort.OperationID, oThisStructPort.OperatorID, oOperation, oOperator
                                
                                If TypeOf oOperator Is IJSeam Then
                                    
                                    If oInputPlates.Contains(ConnectedObjColl.Item(i)) Then
                                        If Not oSeamsColl.Contains(oOperator) Then
                                            oSeamsColl.Add oOperator
                                            'It has to be the inner seam
                                            If eMarkingType = STRMFG_SIGHTLINE_MARK Then
                                                    
                                                Dim oProjCSColl As Collection
                                                'There will be only one part in case copy by seams
                                                Set oProjCSColl = ProjectCurveOnPlatePart(oNewPlates.Item(1), oOperator)
                                                
                                                Dim iIndex As Long
                                                For iIndex = 1 To oProjCSColl.Count
                                                    oConnections.Add CreatePersistentComplexString(oProjCSColl.Item(iIndex))
                                                Next
                                                
                                                oSeamsGeom.Add ConnectionsColl.Item(i)
                                                oConnObjColl.Add oOperator 'Seam
                                                oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                                            End If
                                            
                                        End If
                                    Else
                                        'It is an outer seam
                                        If eMarkingType = STRMFG_SEAM_MARK Then
                                            If Not oSeamsGeom.Contains(ConnectionsColl.Item(i)) Then
                                            
                                                Dim oSeamMarkGeom   As IJComplexString
                                                Set oSeamMarkGeom = GetSeamMarkGeometry(oInputPlates.Item(lOldPlatesCount), ConnectionsColl.Item(i), ThisPortColl.Item(i))
                                    
                                                oSeamsGeom.Add ConnectionsColl.Item(i)
                                                oConnObjColl.Add oOperator 'Seam
                                                OffsetOuterSeamCurve oSeamMarkGeom, oInputPlates.Item(lOldPlatesCount), ThisPortColl.Item(i), ConnectionsColl.Item(i)
                                                oConnections.Add CreatePersistentComplexString(oSeamMarkGeom)
                                                oPhyConnsForMarks.Add ConnectionsColl.Item(i)
                                            End If
                                        End If
                                    End If
                                    
                                End If
                                
                            End If
                                
                        End If 'If TypeOf ConnectedObjColl.Item(i) Is IJPlatePart Then
                    
                    Next i
                    
            End Select
            
        For lNewPlatesCount = 1 To oNewPlates.Count
            CreateMarksOfGivenType.AddElements CreateAPSMarks(oConnObjColl, oConnections, oSidesColl, oPhyConnsForMarks, oInputPlates.Item(lOldPlatesCount), oNewPlates.Item(lNewPlatesCount), eMarkingType)
        Next lNewPlatesCount
            
        Set ConnectedObjColl = Nothing
        Set ConnectionsColl = Nothing
        Set oConnObjColl = Nothing
        Set oConnections = Nothing
        Set ThisPortColl = Nothing
        Set OtherPortColl = Nothing
            
    Next lOldPlatesCount
    
    
Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ProjectCurveOnPlatePart(oPlate As IJPlatePart, oWireBody As IJWireBody) As Collection
    Const METHOD = "ProjectCurveOnPlatePart"
    On Error GoTo ErrorHandler
     
    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = oPlate
    
    Dim oUpsidePort As IJPort
    Set oUpsidePort = oPlateWrapper.GetSurfacePort(SideA)  'Side_A is plate base side
    
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oUpsidePort.Geometry
    
    Dim oCSColl As IJElements
    Dim oOutputColl As Collection
    Set oOutputColl = New Collection
    
    Dim oMfgMGHelper        As IJMfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    oMfgMGHelper.WireBodyToComplexStrings oWireBody, oCSColl
    
    Dim iCount As Integer
    For iCount = 1 To oCSColl.Count
        Dim oInnerCS As IJComplexString
        Set oInnerCS = oCSColl.Item(iCount)

        On Error Resume Next
        Dim oProjCS As IJComplexString
        oMfgMGHelper.ProjectComplexStringToSurface oInnerCS, oSurfaceBody, Nothing, oProjCS
        On Error GoTo ErrorHandler
        
        If Not oProjCS Is Nothing Then
            oOutputColl.Add oProjCS
        End If
        
    Next
    
    Set ProjectCurveOnPlatePart = oOutputColl

Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function CreateAPSMarks(oConnObjColl As Collection, oConnections As Collection, oSidesColl As Collection, oPhyConns As Collection, oOldPlatePart As Object, oNewPlatePart As Object, ByVal eMarkingType As StrMfgGeometryType) As IJElements
     Const METHOD = "CreateAPSMarks"
     On Error GoTo ErrorHandler
     
     Set CreateAPSMarks = New JObjectCollection
     
     Dim i As Integer
     
     For i = 1 To oConnObjColl.Count
        
        Dim lMarkingSide As Long
        
        Dim lMarkingType As Long
        lMarkingType = eMarkingType
        
        If eMarkingType = STRMFG_SEAM_MARK Or eMarkingType = STRMFG_SIGHTLINE_MARK Then
            lMarkingSide = 1111 'Molded side
            eMarkingType = STRMFG_SEAM_MARK
            lMarkingType = eMarkingType
        Else
            Dim oThisStructPort   As IJStructPort
            Dim eThisPortContext  As eUSER_CTX_FLAGS
            Dim strMoldedSide     As String
    
            Set oThisStructPort = oSidesColl.Item(i)
            eThisPortContext = oThisStructPort.ContextID
            
            Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
            Set oSDPlateWrapper = New StructDetailObjects.PlatePart
            Set oSDPlateWrapper.object = oOldPlatePart
                        
            strMoldedSide = oSDPlateWrapper.AlternateMoldedSide
            
            If (eThisPortContext And CTX_BASE) <> 0 Then 'Base
                If strMoldedSide = "Base" Then
                    lMarkingSide = 1111 'molded side
                Else
                    lMarkingSide = 1113 'Anti-molded side
                End If
            Else 'Offset
                If strMoldedSide = "Offset" Then
                    lMarkingSide = 1111
                Else
                    lMarkingSide = 1113
                End If
            End If
        End If
                
        Dim oMarkingLineObj As Object
        Set oMarkingLineObj = CreateMarkingLine(oNewPlatePart, oConnections.Item(i), oConnObjColl.Item(i), oPhyConns.Item(i), lMarkingSide, lMarkingType)
        
        If Not oMarkingLineObj Is Nothing Then
            CreateAPSMarks.Add oMarkingLineObj
        End If
        
        Set oSDPlateWrapper = Nothing
               
    Next i


Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetConnectionContourTee(oAppConnection As Object, oConnectedPart As Object, oInputPlatePart As Object) As Object
    Const METHOD = "GetConnectionContourTee"
    On Error GoTo ErrorHandler
        
    Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oInputPlatePart
   
    Dim eSideOfConnectedObjectToBeMarked As ThicknessSide
    
    If TypeOf oConnectedPart Is IJPlatePart Then
        
        Dim oPartInfo   As IJDPartInfo
        Set oPartInfo = New PartInfo
        
        Dim eMoldedDir  As StructMoldedDirection
        eMoldedDir = oPartInfo.GetPlatePartThicknessDirection(oConnectedPart)
        
        Dim oSDConPlateWrapper   As New StructDetailObjects.PlatePart
        Set oSDConPlateWrapper = New StructDetailObjects.PlatePart
        Set oSDConPlateWrapper.object = oConnectedPart
        
        Dim sMoldedSide As String
        sMoldedSide = oSDConPlateWrapper.MoldedSide
        
        If eMoldedDir = Centered Then
            eSideOfConnectedObjectToBeMarked = SideUnspecified
        Else
            If sMoldedSide = "Base" Then
                eSideOfConnectedObjectToBeMarked = SideA
            ElseIf sMoldedSide = "Offset" Then
                eSideOfConnectedObjectToBeMarked = SideB
            End If
        End If
        
    ElseIf TypeOf oConnectedPart Is IJProfilePart Then
        
        Dim oSDProfileWrapper   As StructDetailObjects.ProfilePart
        Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
        Set oSDProfileWrapper.object = oConnectedPart
        
        Dim oPartSupport        As GSCADSDPartSupport.IJPartSupport
        Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
        
        Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
        Set oPartSupport.Part = oSDProfileWrapper.object
        Set oProfilePartSupport = oPartSupport
               
        eSideOfConnectedObjectToBeMarked = oProfilePartSupport.ThicknessSideAdjacentToLoadPoint
        
    End If
    
    Dim bContourTee  As Boolean
    Dim oVector      As IJDVector
    Dim oWB          As IJWireBody, oTeeWire As IJWireBody
    
    bContourTee = oSDPlateWrapper.Connection_ContourTee(oAppConnection, eSideOfConnectedObjectToBeMarked, oTeeWire, oVector)
    
    Dim oMfgRuleHelper   As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    
    If bContourTee = True Then
        'Bound the wire based on split points, if there are any.
        Set oWB = TrimTeeWireForSplitPCs(oAppConnection, oTeeWire)
        
        'Convert the IJWireBody to ComplexString
        Set GetConnectionContourTee = oMfgRuleHelper.WireBodyToComplexString(oWB)
          
    End If
    
    
Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetConnectionContourLap(oAppConnection As Object, oInputPlatePart As Object) As IJElements
    Const METHOD = "GetConnectionContourLap"
    On Error GoTo ErrorHandler
    
    Dim bContourLap   As Boolean
    Dim oWBColl       As Collection
    Dim oCSElements           As IJElements
    Set oCSElements = New JObjectCollection
    Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oInputPlatePart
       
    bContourLap = oSDPlateWrapper.Connection_ContourLap(oAppConnection, oWBColl)

    If (bContourLap = True) And Not (oWBColl Is Nothing) Then
        Dim lWBIndex  As Long
        Dim oWB       As IJWireBody
        
        Dim oWireBodyColl   As IJElements
        Set oWireBodyColl = New JObjectCollection
        
        For lWBIndex = 1 To oWBColl.Count
            Set oWB = oWBColl.Item(lWBIndex)
            oWireBodyColl.Add oWB
        Next lWBIndex
        
        Dim oMfgGeomHelper As New MfgGeomHelper
        Dim oComplexStrings As IJElements
        Set oComplexStrings = oMfgGeomHelper.MergeCollectionToComplexStrings(oWireBodyColl)
        
        Dim oCS As IJComplexString
        Dim oTempCS As IJComplexString
        For Each oTempCS In oComplexStrings
            Set oCS = CreatePersistentComplexString(oTempCS)
            oCSElements.Add oCS
        Next
    End If
    
    Set GetConnectionContourLap = oCSElements
    Set oWBColl = Nothing

Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description

End Function
Private Function IsPartOfSplitOperation(oAppConnection As Object) As Boolean
        Const METHOD = "IsPartOfSplitOperation"
        On Error GoTo ErrorHandler
    
        IsPartOfSplitOperation = False
        
        ' Check if this physical connection is a Root PC, which particpates in Split operation
        ' If so, Skip the marking line line creation for this Root PC.
        Dim oStructEntityOperation As IJDStructEntityOperation
        Dim opeartionProgID        As String
        Dim opeartionID            As StructOperation
        Dim oOperColl              As New Collection
    
        Set oStructEntityOperation = oAppConnection
        oStructEntityOperation.GetEntityOperation opeartionProgID, opeartionID, oOperColl
        
        Set oStructEntityOperation = Nothing
        Set oOperColl = Nothing
        
        ' If the RootPC has Split operation in its graph, just goto the next pc.
        If opeartionID = ConnectionSplitOperation Then
            ' No need to create marking line.
            IsPartOfSplitOperation = True
        End If
        
        Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function CreatePersistentComplexString(oComplexString As IJComplexString) As IJComplexString
    Const METHOD = "CreatePersistentComplexString"
    On Error GoTo ErrorHandler
        
    If oComplexString Is Nothing Then Exit Function
        
    ' create persistent complex string
    Dim oGeometryFactory   As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Dim oComplexStrings3d As IComplexStrings3d
    Set oComplexStrings3d = oGeometryFactory.ComplexStrings3d
    
    Dim oResourceManager As Object
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oCrvElemets As IJElements
    oComplexString.GetCurves oCrvElemets
    
    Dim oCS As IJComplexString
    Set oCS = oComplexStrings3d.CreateByCurves(oResourceManager, oCrvElemets)
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = oCS
    oControlFlags.ControlFlags(CTL_FLAG_CACHE) = CTL_FLAG_CACHE ' Hide the complex string
    Set oControlFlags = Nothing
    
    Set CreatePersistentComplexString = oCS
 
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


 

Private Function GetSeamMarkGeometry(oInputPlatePart As Object, oAppConnection As Object, oThisPort As IJPort) As IJComplexString
    Const METHOD = "GetSeamMarkGeometry"
    On Error GoTo ErrorHandler
    
    Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oInputPlatePart

    Dim sMoldedSide As String
    sMoldedSide = oSDPlateWrapper.MoldedSide
    
    Dim oEntityHelper As New MfgEntityHelper
    Dim oEdgePort     As IJPort
    
    If sMoldedSide = "Base" Then
        Set oEdgePort = oEntityHelper.GetEdgePortGivenFacePort(oThisPort, CTX_BASE)
    Else
        Set oEdgePort = oEntityHelper.GetEdgePortGivenFacePort(oThisPort, CTX_OFFSET)
    End If
    
    Dim oWireBody As IJWireBody
    Set oWireBody = oEdgePort.Geometry
  
    Dim oMfgMGHelper   As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
        
    Dim oComplexString As IJComplexString
    oMfgMGHelper.WireBodyToComplexString oWireBody, oComplexString
    
    Set GetSeamMarkGeometry = oComplexString
    
        
'    Dim oCommonGeometry             As IUnknown
'    Dim oPhyConGeom                 As Object
'    Dim oPhyConGeomColl             As IJDObjectCollection
'    Dim oPhyConCurve                As IJCurve
'    Dim dStartX                     As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
'    Dim oPhyConPos                  As New DPosition
'    Dim oProjectedPhyConPos         As IJDPosition
'    Dim dDist                       As Double
'    Dim pNormal                     As IJDVector
'    Dim oMfgGeomHelper              As New MfgGeomHelper
'    Dim oTouchingComplexString      As IJComplexString
'    Dim oTouchingComplexStringColl  As IJElements
'
'    Set oTouchingComplexStringColl = New JObjectCollection
'
'    Dim oSDPhysicalConn   As StructDetailObjects.PhysicalConn
'    Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
'    Set oSDPhysicalConn.object = oAppConnection
'
'    Dim oPlateWrapper        As New MfgRuleHelpers.PlatePart
'    Set oPlateWrapper.object = oInputPlatePart
'
'    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
'    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
'    Set oSDPlateWrapper.object = oInputPlatePart
'
'    Dim sMoldedSide As String
'    sMoldedSide = oSDPlateWrapper.MoldedSide
'
'    Dim oUpsidePort As IJPort
'    If sMoldedSide = "Base" Then
'        Set oUpsidePort = oPlateWrapper.GetSurfacePort(Side_A)
'    ElseIf sMoldedSide = "Offset" Then
'        Set oUpsidePort = oPlateWrapper.GetSurfacePort(Side_B)
'    End If
'
'    Dim oMfgRuleHelper   As MfgRuleHelpers.Helper
'    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
'
'    Dim oGeomOpsToolBox   As IJDTopologyToolBox
'    Set oGeomOpsToolBox = New DGeomOpsToolBox
'
'    Set oPhyConGeomColl = oSDPhysicalConn.GetConnectionGeometries
'
'    For Each oPhyConGeom In oPhyConGeomColl
'        Set oPhyConCurve = oPhyConGeom
'        oPhyConCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
'
'        oPhyConPos.Set dStartX, dStartY, dStartZ
'
'        oGeomOpsToolBox.ProjectPointOnSurfaceBody oUpsidePort.Geometry, _
'                    oPhyConPos, oProjectedPhyConPos, pNormal
'
'        dDist = 1000
'        dDist = oProjectedPhyConPos.DistPt(oPhyConPos)
'
'        If dDist < 0.001 Then
'            oPhyConPos.Set dEndX, dEndY, dEndZ
'            oGeomOpsToolBox.ProjectPointOnSurfaceBody oUpsidePort.Geometry, _
'                    oPhyConPos, oProjectedPhyConPos, pNormal
'            dDist = 1000
'            dDist = oProjectedPhyConPos.DistPt(oPhyConPos)
'
'            If dDist < 0.001 Then
'                oTouchingComplexStringColl.Add oMfgRuleHelper.ComplexStringToWireBody(oPhyConGeom)
'            End If
'        End If
'    Next
'
'    If oTouchingComplexStringColl Is Nothing Then Exit Function
'
'    If oTouchingComplexStringColl.Count = 0 Then Exit Function
'
'    oMfgGeomHelper.MergeCollectionToComplexString oTouchingComplexStringColl, oTouchingComplexString
'
'    Set GetSeamMarkGeometry = oTouchingComplexString
'
'    Set oMfgRuleHelper = Nothing
'    Set oGeomOpsToolBox = Nothing
'    Set oPhyConGeomColl = Nothing
'    Set oMfgGeomHelper = Nothing
    
 Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function OffsetOuterSeamCurve(oSeamMarkGeom As IJComplexString, oInputPlatePart As Object, oThisPort As IJPort, oAppConnection As Object)
        Const METHOD = "OffsetOuterSeamCurve"
        On Error GoTo ErrorHandler
           
        Dim oMfgRuleHelper   As MfgRuleHelpers.Helper
        Set oMfgRuleHelper = New MfgRuleHelpers.Helper
        
        Dim oEdgeWire   As IJWireBody
        Set oEdgeWire = oMfgRuleHelper.ComplexStringToWireBody(oSeamMarkGeom)
           
        'OffSet the wirebody from the edge into the platesurface as a fixed distance
        'Get input parameters
        Dim oSDPhysicalConn   As StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oAppConnection
        
        Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
        Set oSDPlateWrapper = New StructDetailObjects.PlatePart
        Set oSDPlateWrapper.object = oInputPlatePart
        
        Dim oPlateWrapper        As New MfgRuleHelpers.PlatePartHlpr
        Set oPlateWrapper.object = oInputPlatePart
         
        Dim sMoldedSide As String
        sMoldedSide = oSDPlateWrapper.MoldedSide
            
        Dim oUpsidePort As IJPort
        If sMoldedSide = "Base" Then
            Set oUpsidePort = oPlateWrapper.GetSurfacePort(SideA)
        ElseIf sMoldedSide = "Offset" Then
            Set oUpsidePort = oPlateWrapper.GetSurfacePort(SideB)
        End If
        
        Dim oOffSetDirection           As IJDVector
        Dim oUnkUpSideSurface          As IUnknown
        Dim oUnkConnectingSideSurface  As IUnknown
        Dim oConnectingSideSurface     As IJSurfaceBody
        Dim oUpsideSurface             As IJSurfaceBody

        Set oUnkUpSideSurface = oUpsidePort.Geometry
        Set oUpsideSurface = oUnkUpSideSurface
        Set oUnkConnectingSideSurface = oThisPort.Geometry
        Set oConnectingSideSurface = oUnkConnectingSideSurface

        ' get a direction from somewhere on the edge
        Dim oPosition1 As IJDPosition, oPosition2 As IJDPosition
        Dim oTangent1 As DVector, oNormalDirection As DVector, oCrossVec As DVector
        
        oEdgeWire.GetEndPoints oPosition1, oPosition2, oTangent1
        oUpsideSurface.GetNormalFromPosition oPosition1, oNormalDirection
        
        Set oCrossVec = oTangent1.Cross(oNormalDirection)
        
        Dim oVector As IJDVector
        oUpsideSurface.GetNormalFromPosition oPosition1, oVector
        oMfgRuleHelper.ScaleVector oVector, -1
                
        ' this gets the direction pointing outwards
        oConnectingSideSurface.GetNormalFromPosition oPosition1, oOffSetDirection
       
       'Get the offset distance by which the Wire Body has to be moved.
       Dim dOffSet As Double
       dOffSet = GetSeamDistance
       
       Dim rootgapvalue As Double
       oSDPhysicalConn.GetBevelParameterValue "RootGap", rootgapvalue, DoubleType
       
       ' Note that often the rootgapvalue is a negative value and therefore
       'I need to add the value to the offset in order to substract
       If rootgapvalue > 0 Then
          dOffSet = dOffSet + rootgapvalue
       Else
          If rootgapvalue < 0 Then
             dOffSet = dOffSet - rootgapvalue
          End If
       End If
        
       ' In case the offset value is less or equal to zero then reset it
       ' to the intial offset values.
       If dOffSet <= 0 Then
           dOffSet = GetSeamDistance
       End If
       
       If oCrossVec.Dot(oOffSetDirection) < 0# Then
            dOffSet = -1# * dOffSet
       End If
       
        dOffSet = -1# * dOffSet
        
       Dim oUnkOffsetCurve As IUnknown
       Set oUnkOffsetCurve = oMfgRuleHelper.OffsetCurve(oUpsideSurface, oEdgeWire, Nothing, dOffSet, False)
        
       Dim oWireBodyUnk  As IUnknown
       Dim oProject      As New IMSModelGeomOps.Project
       
       oProject.CurveAlongVectorOnToSurface Nothing, oUpsideSurface, oUnkOffsetCurve, oVector, Nothing, oWireBodyUnk
        
       'Get the Complex string from the Wire Body.
       Set oSeamMarkGeom = oMfgRuleHelper.WireBodyToComplexString(oWireBodyUnk)
       
Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'Public Function GetOuterSeams(oInputPlates As IJElements) As IJElements
'    Const METHOD = "GetOuterSeams"
'    On Error GoTo ErrorHandler
'
'    Set GetOuterSeams = New JObjectCollection
'
'    Dim oMfgUtilSurface As IJMfgGeomUtil
'    Set oMfgUtilSurface = New MfgUtilSurface
'
'    Dim oOuterSeamColl As Collection
'    Dim oOuterGeomColl As Collection
'    Dim oOuterEdgeNames() As String
'    oMfgUtilSurface.GetEntitiesCreatingOuterEdgesInPlateCollection oInputPlates, Nothing, oOuterSeamColl, oOuterEdgeNames, oOuterGeomColl
'
'    If oOuterSeamColl Is Nothing Then Exit Function
'
'    If oOuterSeamColl.Count = 0 Then Exit Function
'
'    Dim oAllSeamsColl As IJElements
'    Set oAllSeamsColl = New JObjectCollection
'
'    Dim i As Integer
'    For i = 1 To oOuterSeamColl.Count
'        If TypeOf oOuterSeamColl.Item(i) Is IJSeam Then
'            oAllSeamsColl.Add oOuterSeamColl.Item(i)
'        End If
'    Next
'
'    Dim oInnerSeams As IJElements
'    Set oInnerSeams = GetInnerSeams(oInputPlates)
'
'    'Remove the inner seams from all seams collection to get the outer seams
'    oAllSeamsColl.RemoveElements oInnerSeams
'
'    GetOuterSeams.AddElements oAllSeamsColl
'
'Exit Function
'
'ErrorHandler:
'    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
'End Function
'Public Sub GetInnerAndOuterSeams(oInputPlates As IJElements, ByRef oInnerSeams As IJElements, ByRef oOuterSeams As IJElements)
'    Const METHOD = "GetInnerAndOuterSeams"
'    On Error GoTo ErrorHandler
'
'    Dim lOldPlatesCount As Long
'    Dim i As Long
'
'    For lOldPlatesCount = 1 To oInputPlates.Count
'
'        Dim oPartSupport As IJPartSupport
'        Set oPartSupport = New PartSupport
'        Set oPartSupport.Part = oInputPlates.Item(lOldPlatesCount)
'
'        Dim ConnectedObjColl As Collection
'        Dim ConnectionsColl As Collection
'
'        Dim ThisPortColl As Collection
'        Dim OtherPortColl As Collection
'        oPartSupport.GetConnectedObjects ConnectionPhysical, _
'                                         ConnectedObjColl, ConnectionsColl, _
'                                         ThisPortColl, OtherPortColl
'
'        For i = 1 To ConnectedObjColl.Count
'            If TypeOf ConnectedObjColl.Item(i) Is IJPlatePart Then
'
'                Dim oThisStructPort As IJStructPort, oOtherStructPort As IJStructPort
'                Dim ThisPortFlag As eUSER_CTX_FLAGS
'                Dim OtherPortFlag As eUSER_CTX_FLAGS
'
'                Set oThisStructPort = ThisPortColl.Item(i)
'                Set oOtherStructPort = OtherPortColl.Item(i)
'
'                ThisPortFlag = oThisStructPort.ContextID
'                OtherPortFlag = oOtherStructPort.ContextID
'
'                Dim iThisPort As Integer
'                Dim iOtherPort As Integer
'
'                iThisPort = ThisPortFlag And CTX_LATERAL
'                iOtherPort = OtherPortFlag And CTX_LATERAL
'
'                If iThisPort > 0 And iOtherPort > 0 Then
'
'                    Dim oStructDetailHelper As IJStructDetailHelper
'                    Set oStructDetailHelper = New StructDetailHelper
'
'                    Dim oOperator As Object
'                    Dim oOperation As IJStructOperation
'
'                    'The operator has to be a seam
'                    oStructDetailHelper.FindOperatorForOperationInGraphByID oInputPlates.Item(lOldPlatesCount), oThisStructPort.OperationID, oThisStructPort.OperatorID, oOperation, oOperator
'
'                    If TypeOf oOperator Is IJSeam Then
'                        Dim oSeamMarkGeom As IJComplexString
'                        Set oSeamMarkGeom = GetSeamMarkGeometry(oInputPlates.Item(lOldPlatesCount), ConnectionsColl.Item(i))
'
'                        If oInputPlates.Contains(ConnectedObjColl.Item(i)) Then
'                            'It has to be the inner seam
'                            oInnerSeams.Add oOperator
'                        Else
'                            'It is an outer seam
'                            oOuterSeams.Add oOperator
'                        End If
'                    End If
'
'                End If
'
'           End If
'
'      Next i
'
' Next lOldPlatesCount
'
'Exit Sub
'
'ErrorHandler:
'    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
'End Sub
