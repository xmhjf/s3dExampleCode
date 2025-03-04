VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberSketchedSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'
'
'History
'    SS         10/29/08      Creation
'---------------------------------------------------------------------------------------
Option Explicit

Private Const MODULE = "MemberSketchedSym"
Private Const MYSYMBOL_PROGID = "SPSSketchedShapes.MemberSketchedSym"
Private Const MYSYMBOL_LIBNAME = "SPSProgrammedShapes_MemberSketchedSym"
Private Const RADSYMBOL_REPRESENTATIONNAME = "SimplePhysical"

Implements IJDUserSymbolServices
Implements IStructInsulationGraphicInputFilter

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler

    Dim insertTypeVal As IMSInsertionType
    Dim strFilePath As String, strFileName As String

    insertTypeVal = igLinked
    strFileName = definitionParameters
     
    strFilePath = "%OLE_SERVER%\" & strFileName
                            
    IJDUserSymbolServices_GetDefinitionName = MYSYMBOL_PROGID + "|" + _
                                              Str(insertTypeVal) + "|" + strFilePath
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    DefineEncasementSymbolInputs pSymbolDefinition
   
    DefineEncasementGraphicSymbolOutput MYSYMBOL_LIBNAME, MYSYMBOL_PROGID, pSymbolDefinition
    
    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler

    Dim pSymbolDefinition As IJDSymbolDefinition
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set pSymbolDefinition = oSymbolFactory.CreateEntity(definition, pResourceMgr)
    Set oSymbolFactory = Nothing
    
    pSymbolDefinition.ProgId = MYSYMBOL_PROGID
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(definitionParameters)
    pSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    pSymbolDefinition.SupportOnlyOption = igSYMBOL_SUPPORT_ONLY    'go away when no occurences
    
    pSymbolDefinition.IJDServerIdentification.ServerName = "%OLE_SERVER%\" & definitionParameters
    pSymbolDefinition.IJDServerIdentification.InsertionType = igLinked
    
    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub CMEvaluateEncasementCrossSection(pRepSCM As IJDRepresentationStdCustomMethod)
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
 
    Dim oRepDuringGame As IMSSymbolEntities.IJDRepresentationDuringGame
    Dim oInput As IJDInput
    
    Dim iIJProxy As IJDProxy
    Dim iInsulationSpec As IStructInsulationSpec
    Dim iMemberPart As ISPSMemberPartPrismatic
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    
    'Dim pOutputs As IMSSymbolEntities.IJDOutputs
    Dim outputColl As IJDOutputCollection
    
    Dim repID As Long
    Dim bCalculateIsOff As Boolean
    Dim ServerDoc As RAD2D.Document
    Dim ActiveSheet As RAD2D.Sheet
    Dim RAD2DSymbolDef As RAD2D.SymbolProperties
    Dim RAD2DParams As SymbolDefinitionParameters
    Dim RAD2DParam As SymbolDefinitionParameter
    Dim RAD2DGroup As RAD2D.group, RAD2DGroup_Outer As RAD2D.group
    Dim oStructConverter As IMSTools2D.GeometryConverter

    Dim ii As Long, countParameters As Long
    Dim strAttrName As String, strParamName As String
    Dim oDefinition As Object
    
    Dim XServices As New CrossSectionServices
    Dim iSblDefHelper As IJDSymbolDefHelper
    Dim ParamAndProp() As Variant
    Dim vValue As Variant

    Dim bDebug As Boolean
    bDebug = False
    bCalculateIsOff = False

    Set oRepDuringGame = pRepSCM
    Set outputColl = oRepDuringGame.outputCollection
    
    Set oInput = oRepDuringGame.definition.IJDInputs.GetInputByIndex(1)
    Set iIJProxy = oInput.IJDInputDuringGame.Argument
    Set iInsulationSpec = iIJProxy.Source

    Set oInput = oRepDuringGame.definition.IJDInputs.GetInputByIndex(3)
    Set iMemberPart = oInput.IJDInputDuringGame.Argument
    Set oDefinition = iMemberPart.CrossSection.definition

    Set iSblDefHelper = iMemberPart.CrossSection.definition
    iSblDefHelper.GetParametersAndProperties ParamAndProp()
    
    On Error Resume Next
    Set ServerDoc = oRepDuringGame.definition.IJDServerIdentification.object
    If ServerDoc Is Nothing Then
        MsgBox "serverdoc is nothing"
    End If
    Set ActiveSheet = ServerDoc.ActiveSheet

    'Build IMSTools2D converter
    Set oStructConverter = New IMSTools2D.GeometryConverter
    oStructConverter.Context3D = outputColl.ResourceManager
    oStructConverter.Context2D = ServerDoc
    
    Set RAD2DSymbolDef = ServerDoc.SymbolProperties
    'Test if not a .sha file, if not it's a .sym file, load parameters and dimensions from attributes
    'of the member part's cross-section or he specification thickness.

    ' Read the actual document, and populate the parameter values from the member's cross-section or spec thickness
    If Not RAD2DSymbolDef Is Nothing Then

        ServerDoc.CalculateOff
        bCalculateIsOff = True
        If bDebug = True Then
            SaveCopyAs "Symbols\" & ServerDoc.Name & "_BeforeEvaluate", ServerDoc
        End If
        
        Set RAD2DParams = RAD2DSymbolDef.SymbolDefinitionParameters
        countParameters = RAD2DParams.Count
        '' MsgBox "countParameters = " & countParameters

        For ii = 1 To countParameters
            Set RAD2DParam = RAD2DParams.Item(ii)
            If Not RAD2DParam.Parameter Is Nothing Then     ' if it is a zombie parameter, ignore
                Dim oRADObject As Object
                Dim strRADKey As String

                strParamName = RAD2DParam.Name
                If strParamName = INSULATION_THICKNESS_SYMBOL_PARAMETER Then
                    vValue = iInsulationSpec.Thickness
                Else
                    GetAttributeNameFromSymbolParameter ParamAndProp, strParamName, strAttrName
                    If strAttrName <> "" Then
                        XServices.GetCrossSectionAttributeValue oDefinition, strAttrName, vValue
                    Else
                        Err.Raise E_FAIL, METHOD, "Cannot find attribute for symbol parameter " & strParamName
                    End If
                End If

                Set oRADObject = RAD2DParam.Parameter
                strRADKey = ""
                On Error Resume Next
                strRADKey = oRADObject.Key
                If strRADKey = "" Then
                    Set oRADObject = ServerDoc.Variables.Item(RAD2DParam.Name)
                Else
                    Set oRADObject = ActiveSheet.DrawingObjects(strRADKey)
                End If
                On Error GoTo ErrorHandler

                oRADObject.Value = vValue
            End If
        Next ii
        If bDebug = True Then
            SaveCopyAs "Symbols\" & ServerDoc.Name & "_AfterApplyingParameters", ServerDoc
        End If

    Else
        MsgBox "RAD2DSymbolDef is nothing"
    End If

    ' done reading inputs and setting the values.  Now calculate the RAD graphics.
    ServerDoc.CalculateOn
    bCalculateIsOff = False
    If bDebug = True Then
        SaveCopyAs "Symbols\" & ServerDoc.Name & "_AfterEvaluate", ServerDoc
    End If

    Set RAD2DGroup = Nothing

    ' We may need to move the insulation graphics to match the member.
    ' Do that by reading cardinal point from each, and if it exists and the delta is not zero,
    ' then transform output on the x-y plane from the rad position to the member's position
    
    Dim radX As Double, radY As Double, memX As Double, memY As Double
    Dim tol As Double, vecX As Double, vecY As Double
    Dim i4x4 As IJDT4x4
    
    tol = oGeomFactory.GeometryServices.DistTolerance

    On Error Resume Next
    RAD2DSymbolDef.GetRepresentation "LoadPoints", repID, RAD2DGroup
    Err.Clear
    On Error GoTo ErrorHandler
    
    If GetRADLoadPointCoordinates(RAD2DGroup, 1, radX, radY) Then
    
        XServices.GetCardinalPoint iMemberPart.CrossSection.symbol, 1, memX, memY
        
        vecX = memX - radX
        vecY = memY - radY
        ''  MsgBox "vecx=" & vecX & ", vecy=" & vecY
        
        If Abs(vecX) > tol Or Abs(vecY) > tol Then
            Set i4x4 = New DT4x4
            i4x4.LoadIdentity
            i4x4.IndexValue(12) = vecX
            i4x4.IndexValue(13) = vecY
        End If
    End If

    Set RAD2DGroup = Nothing

    ' Now read the output groups IN THE SAME MANNER as used during Initialize_Definition
    ' ( it is during initialize that holes are identified using output.Description )
    ' read the group, setup an array of the radObjects in end-point order
    ' convert to 3D, and make a complex string, transform it if necessary, and set it as the output object
    repID = 0
    RAD2DSymbolDef.GetRepresentation RADSYMBOL_REPRESENTATIONNAME, repID, RAD2DGroup

    ' check whether a group in the rep is identified as an "outer" group
    ' if so, that implies that holes exist.  The holes could be nested inside the outer or just be siblings.
    
    Dim ele As Object, ele2 As Object
    Dim radSubGroup As RAD2D.group
    Dim strOuter As String
    Dim outputIndex As Long
    Dim oWireCollection As IJElements
    Set oWireCollection = New JObjectCollection
    Set RAD2DGroup_Outer = GetOuterGroup(RAD2DGroup)
    If RAD2DGroup_Outer Is RAD2DGroup Then      ' the rep group is outer, but has no holes
        Set RAD2DGroup_Outer = Nothing
    End If
    
    Dim oWireToAdd As Object
    If RAD2DGroup_Outer Is Nothing Then
        If TypeOf RAD2DGroup.Item(1) Is RAD2D.group Then
            outputIndex = 1
            For Each ele In RAD2DGroup
                ' to make each wire body persistent, pass in
                ' outputColl.ResourceManager below.  The merge
                ' will persist the entire milti-lump body though
                Set oWireToAdd = GetAttributedWireBodyFromRADGroup(Nothing, _
                                                                  ele, _
                                                                  oStructConverter, _
                                                                  oGeomFactory, _
                                                                  i4x4)
                
               
                oWireCollection.Add oWireToAdd
                If bDebug = True Then
                    SaveSatFile "Symbols\" & ServerDoc.Name & "_" & RAD2DGroup.Name, oWireToAdd
                End If
                
                outputIndex = outputIndex + 1

            Next

        Else
            Set oWireToAdd = GetAttributedWireBodyFromRADGroup(Nothing, _
                                                              RAD2DGroup, _
                                                              oStructConverter, _
                                                              oGeomFactory, _
                                                              i4x4)
            oWireCollection.Add oWireToAdd
            If bDebug = True Then
                SaveSatFile "Symbols\" & ServerDoc.Name & "_" & RAD2DGroup.Name, oWireToAdd
            End If
        End If

    Else
        outputIndex = 1
        ' add outer group first.
        Set oWireToAdd = GetAttributedWireBodyFromRADGroup(Nothing, _
                                                   RAD2DGroup_Outer, _
                                                   oStructConverter, _
                                                   oGeomFactory, _
                                                   i4x4)
                                                                            
                If bDebug = True Then
                    SaveSatFile "Symbols\" & ServerDoc.Name & "_" & RAD2DGroup.Name, oWireToAdd
                End If
        oWireCollection.Add oWireToAdd
        Set oWireToAdd = Nothing
        outputIndex = outputIndex + 1

        For Each ele In RAD2DGroup

            If Not ele Is RAD2DGroup_Outer Then
        
                strOuter = GetRADAttributeValueOrKey("Concentric", ele, "Position")
                If LCase(strOuter) = "outer" Then
                
                    Set oWireToAdd = GetAttributedWireBodyFromRADGroup(Nothing, _
                                                                      ele, _
                                                                      oStructConverter, _
                                                                      oGeomFactory, _
                                                                      i4x4)
                    oWireCollection.Add oWireToAdd
                    If bDebug = True Then
                        SaveSatFile "Symbols\" & ServerDoc.Name & "_" & RAD2DGroup.Name, oWireToAdd
                    End If
                    outputIndex = outputIndex + 1
                    
                    ' check if user has nested holes within the outer group
                    Set radSubGroup = ele
                    If TypeOf radSubGroup.Item(1) Is RAD2D.group Then
                        For Each ele2 In RAD2DGroup_Outer
                           Set oWireToAdd = GetAttributedWireBodyFromRADGroup(Nothing, _
                                                                      ele2, _
                                                                      oStructConverter, _
                                                                      oGeomFactory, _
                                                                      i4x4)
                                                                                               
                            If bDebug = True Then
                                SaveSatFile "Symbols\" & ServerDoc.Name & "_" & RAD2DGroup.Name, oWireToAdd
                            End If
                            oWireCollection.Add oWireToAdd
                            Set oWireToAdd = Nothing
                            outputIndex = outputIndex + 1
                        Next
                    End If
                Else
                    Set oWireToAdd = GetAttributedWireBodyFromRADGroup(Nothing, _
                                                                      ele, _
                                                                      oStructConverter, _
                                                                      oGeomFactory, _
                                                                      i4x4)
                    oWireCollection.Add oWireToAdd
                    If bDebug = True Then
                        SaveSatFile "Symbols\" & ServerDoc.Name & "_" & RAD2DGroup.Name, oWireToAdd
                    End If
                    outputIndex = outputIndex + 1
                    Set oWireToAdd = Nothing
                End If
            End If
        Next
    End If
    
    
    Dim pIJDProfile2dHelper As IMSModelGeomOps.IJDProfile2dHelper
    Set pIJDProfile2dHelper = New DProfile2dHelper
    
    Dim oMergedWire As Object
    Set oMergedWire = pIJDProfile2dHelper.MergeWireBodies(oWireCollection, outputColl.ResourceManager)
    
    If bDebug = True Then
        SaveSatFile "Symbols\" & ServerDoc.Name & "_" & ENCASEMENTSYMBOL_OUTPUTNAME, oMergedWire
    End If

    
    outputColl.AddOutput ENCASEMENTSYMBOL_OUTPUTNAME, _
                         oMergedWire
    
    
    
    Exit Sub

ErrorHandler:
    If Not ServerDoc Is Nothing Then
        If bCalculateIsOff Then
            ServerDoc.CalculateOn
        End If
    End If
    HandleError MODULE, METHOD
End Sub

Private Function IStructInsulationGraphicInputFilter_FilterGraphicInput(ByVal oGraphicObject As Object, ByVal oDefinition As Object) As Long
Const METHOD = "IStructInsulationGraphicInputFilter_FilterGraphicInput"
On Error GoTo ErrorHandler

    IStructInsulationGraphicInputFilter_FilterGraphicInput = FilterMembers(oGraphicObject, oDefinition)

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Function IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria() As String
    IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria = "[" & MYSYMBOL_PROGID & ",FilterMembers]"
End Function

Private Function IStructInsulationGraphicInputFilter_SingleGraphicInput() As Boolean
    IStructInsulationGraphicInputFilter_SingleGraphicInput = True
End Function

Public Function FilterMembers(oGraphicObject As Object, oDefinition As Object) As Integer
    Const METHOD = "FilterMembers"
    On Error GoTo ErrorHandler

    FilterMembers = ValidateMemberFPCriteria(oGraphicObject, oDefinition)
    If FilterMembers = StructInsulationInputHelper_Ok Then
        FilterMembers = 1
    Else
        FilterMembers = 0
    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub GetAttributeNameFromSymbolParameter(ParamAndProp() As Variant, strParameterName As String, _
                        ByRef strAttributeName As String)
    Const METHOD = "GetAttributeNameFromSymbolParameter"
    On Error GoTo ErrorHandler
    
    Dim lCount1 As Long
    Dim lCount2 As Long
    Dim lCount3 As Long
    Dim lCount4 As Long
    Dim ii As Integer, jj As Integer
    Dim strName As String
    Dim strLCaseParameterName As String

    lCount1 = LBound(ParamAndProp, 1)
    lCount2 = UBound(ParamAndProp, 1)
    lCount3 = LBound(ParamAndProp, 2)
               
    strAttributeName = ""
    strLCaseParameterName = LCase(strParameterName)

    If lCount2 > 0 Then
        For ii = lCount1 To lCount2
            jj = lCount3 + 1
            strName = LCase(ParamAndProp(ii, jj))
            If strName = strLCaseParameterName Then
                jj = lCount3 + 3
                strAttributeName = LCase(ParamAndProp(ii, jj))
                Exit For
            End If
        Next ii
    End If
    
    '' MsgBox "paramName=" & strParameterName & ", attrName=" & strAttributeName
    
    Exit Sub
  
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'' ============================================================================================
''        from m:\CatalogData\BulkLoad\CatalogProcessor\modPartClassHelperFunctions.bas
''        If bIsMapExists = False Then
''            lUboundForPropertyMap = lUboundForPropertyMap + 1
''            ReDim Preserve avarSblParameterInfo(1 To 14, 1 To lUboundForPropertyMap) As Variant
''            avarSblParameterInfo(1, lUboundForPropertyMap) = SYMBOL_PARAMETER
''            avarSblParameterInfo(2, lUboundForPropertyMap) = strSblMapName
''            avarSblParameterInfo(3, lUboundForPropertyMap) = Property
''            If Len(strUserPropName) > 0 Then
''                avarSblParameterInfo(4, lUboundForPropertyMap) = strUserPropName
''            End If
''            avarSblParameterInfo(5, lUboundForPropertyMap) = "OccProperty"
''            If Len(strOccPropName) > 0 Then
''                avarSblParameterInfo(6, lUboundForPropertyMap) = strOccPropName
''            End If
''            avarSblParameterInfo(7, lUboundForPropertyMap) = INTERFACETYPE
''            avarSblParameterInfo(8, lUboundForPropertyMap) = lInterfaceType
''            avarSblParameterInfo(9, lUboundForPropertyMap) = "IID"
''            avarSblParameterInfo(10, lUboundForPropertyMap) = varUserIID
''            avarSblParameterInfo(11, lUboundForPropertyMap) = "OccIID"
''            avarSblParameterInfo(12, lUboundForPropertyMap) = varOccIID
''            avarSblParameterInfo(13, lUboundForPropertyMap) = "NozzleId"
''        End If
'' ============================================================================================

Public Function GetOuterGroup(RAD2DGroup As RAD2D.group) As RAD2D.group

    Dim strOuter As String

    Set GetOuterGroup = Nothing
    Dim ele As Object

    strOuter = GetRADAttributeValueOrKey("Concentric", RAD2DGroup, "Position")
    If LCase(strOuter) = "outer" Then
        Set GetOuterGroup = RAD2DGroup
    Else
        For Each ele In RAD2DGroup
            If ele.Type = igGroup Then
                Set GetOuterGroup = GetOuterGroup(ele)
                If Not GetOuterGroup Is Nothing Then
                    Exit Function
                End If
            End If
        Next
    End If
    
End Function


Public Function GetRADAttributeValueOrKey(strAttrSetName As String, _
                                        oRADObject As Object, strAttName As String) As String
    Const METHOD = "GetRADAttributeValueOrKey"

    Dim oAttribute As RAD2D.Attribute
    Dim bAttrFound As Boolean
    Dim oAttributeSet As RAD2D.AttributeSet
    Dim strLCaseAttName As String
    
    On Error GoTo ErrorHandler
    
    GetRADAttributeValueOrKey = ""
    strLCaseAttName = LCase(strAttName)
    
    If oRADObject Is Nothing Then
        Exit Function
    End If
    
    GetAttributeSetOnRADObj oRADObject, strAttrSetName, oAttributeSet
    
    If Not oAttributeSet Is Nothing Then
        For Each oAttribute In oAttributeSet
        
            If LCase(oAttribute.Name) = strLCaseAttName Then
                GetRADAttributeValueOrKey = oAttribute.Value
                bAttrFound = True
                Exit For
            End If
        Next oAttribute
    End If
    
    If Not bAttrFound Then
        GetRADAttributeValueOrKey = oRADObject.Key
    End If
    
 Exit Function
 
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub GetAttributeSetOnRADObj(oGraphicObj As Object, _
                                    strAttrSetName As String, _
                                    oNamedAttributeSet As Object)
    Const METHOD = "GetAttributeSetOnRADObj"
    On Error GoTo ErrorHandler
  
    Dim oAttributeSet As RAD2D.AttributeSet
  
    Set oNamedAttributeSet = Nothing
    
    If Not oGraphicObj.AttributeSets Is Nothing Then

        For Each oAttributeSet In oGraphicObj.AttributeSets

            If LCase(oAttributeSet.SetName) = LCase(strAttrSetName) Then
                Set oNamedAttributeSet = oAttributeSet
                Exit For
            End If

        Next oAttributeSet
        
    End If
    
    Exit Sub
    
ErrorHandler:
   HandleError MODULE, METHOD
End Sub

Public Function GetRADLoadPointCoordinates(RAD2DGroup_LoadPoint As Object, cpIndex As Long, ByRef radX As Double, ByRef radY As Double) As Boolean
    Const METHOD = "GetRADLoadPointCoordinates"
    On Error GoTo ErrorHandler

    Dim strLoadPoints As String, strLoadPointIndex As String
    Dim bFoundIt As Boolean
    Dim ele As Object
    Dim point2d As RAD2D.point2d

    Dim lpIndex As Long
    Dim lPos1 As Long, lPos2 As Long, lPrevPos1 As Long
    Dim lLenStrLoadPoints As Long

    Dim RAD2DGroup As RAD2D.group
    
    bFoundIt = False
    
    GetRADLoadPointCoordinates = False
    If RAD2DGroup_LoadPoint Is Nothing Then
        GetRADLoadPointCoordinates = False
        Exit Function

    ElseIf TypeOf RAD2DGroup_LoadPoint Is RAD2D.group Then
        Set RAD2DGroup = RAD2DGroup_LoadPoint
        For Each ele In RAD2DGroup
            If GetRADLoadPointCoordinates(ele, cpIndex, radX, radY) Then
                GetRADLoadPointCoordinates = True
                Exit Function
            End If
        Next
    
    ElseIf Not TypeOf RAD2DGroup_LoadPoint Is RAD2D.point2d Then
        GetRADLoadPointCoordinates = False
        Exit Function
    End If

    strLoadPoints = GetRADAttributeValueOrKey("PointNames", RAD2DGroup_LoadPoint, "Load Point")
    lLenStrLoadPoints = Len(strLoadPoints)
    If lLenStrLoadPoints = 0 Then
        GetRADLoadPointCoordinates = False
        Exit Function
    End If

    Set point2d = RAD2DGroup_LoadPoint

    lPos1 = 0
    lPos2 = 0
    lPrevPos1 = 0

    Do
        lPos1 = InStr(lPrevPos1 + 1, strLoadPoints, ",", vbTextCompare)
        If lPos1 > 0 Then
            lPos2 = InStr(lPos1 + 1, strLoadPoints, ",", vbTextCompare)
            If lPos2 > 0 Then
                strLoadPointIndex = Mid$(strLoadPoints, lPrevPos1 + 1, lPos2 - lPos1 - 1)
            Else
                strLoadPointIndex = Mid$(strLoadPoints, lPrevPos1 + 1)
            End If
            If Len(strLoadPointIndex) > 0 Then
                lpIndex = CInt(strLoadPointIndex)
                If lpIndex = cpIndex Then
                    radX = point2d.x
                    radY = point2d.y
                    bFoundIt = True
                End If
            End If
        Else
            If lPrevPos1 > 0 Then
                strLoadPointIndex = Mid$(strLoadPoints, lPrevPos1 + 1)
            Else
                strLoadPointIndex = strLoadPoints
            End If
            If Len(strLoadPointIndex) > 0 Then
                lpIndex = CInt(strLoadPointIndex)
                If lpIndex = cpIndex Then
                    radX = point2d.x
                    radY = point2d.y
                    bFoundIt = True
                End If
            End If
        End If
        lPrevPos1 = lPos1
    Loop While (bFoundIt = False And lPos1 > 0)

    GetRADLoadPointCoordinates = bFoundIt

    Exit Function

ErrorHandler:
   HandleError MODULE, METHOD
End Function

Public Sub DefineEncasementGraphicSymbolOutput(strLibName As String, strLibProgId As String, _
                                pSymbolDefinition As IJDSymbolDefinition)
                                
    Const METHOD = "DefineEncasementGraphicSymbolOutput"
    On Error GoTo ErrorHandler
    
    Dim lLibCookie As Long
    Dim lMethodCookie As Long
    Dim LibDesc As DLibraryDescription
    Dim LibDescCustomMethods As IJDLibraryDescription
    Dim repID As Long, outputIndex As Long
    
    'Defines the output of this symbol
    Dim oRep As IJDRepresentation
    Dim oOutput As IJDOutput

    Dim ServerDoc As RAD2D.Document
    Dim RAD2DSymbolDef As RAD2D.SymbolProperties
    'Dim RAD2DParams As SymbolDefinitionParameters
    'Dim RAD2DParam As SymbolDefinitionParameter
    Dim strOuter As String
    Dim ele As Object, ele2 As Object
    Dim RAD2DGroup As RAD2D.group, RAD2DGroup_Outer As RAD2D.group

    ' Set the reference of myself as library
    Set LibDesc = pSymbolDefinition.IJDUserMethods.GetLibrary(strLibName)
    If LibDesc Is Nothing Then
        Set LibDescCustomMethods = New DLibraryDescription
        
        LibDescCustomMethods.Name = strLibName
        LibDescCustomMethods.Type = imsLIBRARY_IS_ACTIVEX
        LibDescCustomMethods.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
        LibDescCustomMethods.Source = strLibProgId
        pSymbolDefinition.IJDUserMethods.SetLibrary LibDescCustomMethods
        Set LibDesc = pSymbolDefinition.IJDUserMethods.GetLibrary(strLibName)
    End If

    If LibDesc Is Nothing Then
        WriteError METHOD, "IJUserMethods.GetLibrary failed for " & strLibName
        Exit Sub
    End If

    lLibCookie = LibDesc.Cookie
    lMethodCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMEvaluateEncasementCrossSection", lLibCookie)
    If lMethodCookie = 0 Then
        WriteError METHOD, "IJUserMethods.GetMethodCookie failed for " & strLibName & ", CMEvaluateEncasementCrossSection"
        Exit Sub
    End If

    Set oRep = New DRepresentation
    oRep.Name = ENCASEMENT_CROSSSECTION_REPNAME
    oRep.RepresentationId = 1
    oRep.Description = "Member Encasement"
    oRep.Properties = igCOLLECTION_VARIABLE     ' may not need to be variable..
    oRep.IJDRepresentationStdCustomMethod.SetCMEvaluate lLibCookie, lMethodCookie
    
    Set oOutput = New DOutput
    oOutput.Name = ENCASEMENTSYMBOL_OUTPUTNAME
    oOutput.Description = "Cross Section Of Encasement"
    oRep.IJDOutputs.Add oOutput
    pSymbolDefinition.IJDRepresentations.Add oRep
    
    
    

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub ConvertAndAddGroupToOutput(outputColl As IJDOutputCollection, RAD2DGroup As Object, _
                pConverter As IMSTools2D.GeometryConverter, oGeomFactory As IngrGeom3D.GeometryFactory, _
                                                Name As String, i4x4 As IJDT4x4)
    
    Const METHOD = "ConvertAndAddGroupToOutput"
    On Error GoTo ErrorHandler
    
    Dim objOrderedArray() As Object, ii As Long, lNumReturned As Long
    Dim oSP3DCurve As Object
    Dim elesSP3DCurves As IJElements
    Dim oComplexString As ComplexString3d
    Dim iIJDObject As IJDObject

    ReturnGroupAsOrderedArray RAD2DGroup, objOrderedArray, lNumReturned

    Set elesSP3DCurves = New JObjectCollection

    For ii = 0 To lNumReturned - 1
        Set oSP3DCurve = Convert2DTo3D(objOrderedArray(ii), pConverter)
        If Not oSP3DCurve Is Nothing Then
            elesSP3DCurves.Add oSP3DCurve
        End If
    Next ii

    Set oComplexString = oGeomFactory.ComplexStrings3d.CreateByCurves(outputColl.ResourceManager, elesSP3DCurves)

    If Not i4x4 Is Nothing Then
        oComplexString.Transform i4x4
    End If

    outputColl.AddOutput Name, oComplexString
    Set oComplexString = Nothing

    For ii = 1 To elesSP3DCurves.Count
        Set iIJDObject = elesSP3DCurves.Item(ii)
        iIJDObject.Remove
    Next ii
    Set elesSP3DCurves = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Function GetAttributedWireBodyFromRADGroup(pResourceManager As IUnknown, _
                                             RAD2DGroup As Object, _
                                             pConverter As IMSTools2D.GeometryConverter, _
                                             oGeomFactory As IngrGeom3D.GeometryFactory, _
                                             i4x4 As IJDT4x4) As Object
    
    Const METHOD = "GetAttributedWireBodyFromRADGroup"
    On Error GoTo ErrorHandler
    
    Dim objOrderedArray() As Object, ii As Long, lNumReturned As Long
    Dim oSP3DCurve As Object
    Dim elesSP3DCurves As IJElements
    Dim oComplexString As ComplexString3d
    Dim iIJDObject As IJDObject

    ReturnGroupAsOrderedArray RAD2DGroup, objOrderedArray, lNumReturned

    Set elesSP3DCurves = New JObjectCollection
    Dim lXIDArray() As Long
    Dim lOPTArray() As Long
    Dim lOPRArray() As Long
    ReDim lXIDArray(lNumReturned)
    ReDim lOPTArray(lNumReturned)
    ReDim lOPRArray(lNumReturned)
    Dim bDebug As Boolean
    bDebug = False
    For ii = 0 To lNumReturned - 1
        Set oSP3DCurve = Convert2DTo3D(objOrderedArray(ii), pConverter)
        If Not oSP3DCurve Is Nothing Then
            elesSP3DCurves.Add oSP3DCurve
            lOPTArray(ii) = 2000
            lXIDArray(ii) = CLng(GetRADAttributeValueOrKey("EdgeNames", objOrderedArray(ii), "EdgeName"))
            lOPRArray(ii) = CLng(objOrderedArray(ii).Key)
        End If
    Next ii

    ' don't want to persist the cpx so pass in Nothing for the resource manager
    Set oComplexString = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, _
                                                                     elesSP3DCurves)

    If Not i4x4 Is Nothing Then
        oComplexString.Transform i4x4
    End If
    Dim XServices As New CrossSectionServices
    Dim oNormVec As IJDVector
    Set oNormVec = New DVector
    Dim strOuter As String
    strOuter = GetRADAttributeValueOrKey("Concentric", RAD2DGroup, "Position")
    Dim lInOrOut As Long
    lInOrOut = 0
    If LCase(strOuter) = "outer" Then
        lInOrOut = 1
        oNormVec.Set 0#, 0#, 1#
    Else
         oNormVec.Set 0#, 0#, -1#
    End If
    
    
    
    ' convert the complex string into a wire body
    Dim oIJGeometryMisc As IMSModelGeomOps.IJGeometryMisc
    Set oIJGeometryMisc = New IMSModelGeomOps.DGeomOpsMisc
    Dim oWireBody As Object
    oIJGeometryMisc.CreateModelGeometryFromGType pResourceManager, _
                                                 oComplexString, _
                                                 Nothing, _
                                                 oWireBody
    Set oComplexString = Nothing

    'Attribute wire body in order to monikerize later
    Dim pIJDProfile2dHelper As IMSModelGeomOps.IJDProfile2dHelper
    Set pIJDProfile2dHelper = New DProfile2dHelper
   
    
    
    With pIJDProfile2dHelper
        .SetIntegerAttributesOnWireBodyEdges oWireBody, _
                                             "JSXid", _
                                             lNumReturned, _
                                             lXIDArray(0)
        .SetIntegerAttributesOnWireBodyEdges oWireBody, _
                                             "JSOpr", _
                                             lNumReturned, _
                                             lOPRArray(0)
       .SetIntegerAttributesOnWireBodyEdges oWireBody, _
                                             "JSOpt", _
                                             lNumReturned, _
                                             lOPTArray(0)
        .SetIntegerAttributeOnWireBodyLump oWireBody, _
                                             "JSIn_Out", _
                                             lInOrOut
    End With
    If bDebug = True Then
        SaveSatFile "Symbols\" & RAD2DGroup.Application.ActiveDocument.Name & "_" & RAD2DGroup.Name, oWireBody
    End If

   Set GetAttributedWireBodyFromRADGroup = oWireBody

    For ii = 1 To elesSP3DCurves.Count
        Set iIJDObject = elesSP3DCurves.Item(ii)
        iIJDObject.Remove
    Next ii
    Set elesSP3DCurves = Nothing
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function SaveCopyAs(sName As String, _
                            ServerDoc As Document) As String
Const MT = "SaveCopyAs"
On Error GoTo ErrorHandler
    
    Dim sFileName As String
    
    On Error Resume Next
    sFileName = Environ("TEMP") & "\" & sName & Format$(Time, "hh-mm-ss") & ".sym"
    ServerDoc.SaveCopyAs sFileName
    On Error GoTo ErrorHandler
    
    SaveCopyAs = sFileName
    
    Exit Function
    
ErrorHandler:
     MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
End Function

Public Function SaveSatFile(sName As String, _
                             pModelBody As Object) As String
Const MT = "SaveSatFile"
On Error GoTo ErrorHandler
    
    Dim lRandom As Long
    Dim sFileName As String
    Dim oModelBody As IJDModelBody
    
    On Error Resume Next

    sFileName = Environ("TEMP") & "\" & sName & "_" & _
                Format$(Time, "hh-mm-ss") & ".sat"
    
    Set oModelBody = pModelBody
    If Not oModelBody Is Nothing Then
        oModelBody.DebugToSATFile sFileName
        Set oModelBody = Nothing
    End If
    
    On Error GoTo ErrorHandler
    
    SaveSatFile = sFileName

    Exit Function
    
ErrorHandler:
     MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description

End Function

