Attribute VB_Name = "MemberEncasementCommon"
'---------------------------------------------------------------------------
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'
'
'History
'    SS         10/29/08      Creation
'---------------------------------------------------------------------------------------
Option Explicit

Public Const E_FAIL = -2147467259

Public Const IJWeightCG = "{DC284B37-5A00-11D2-BE2B-0800364AA003}"
Public Const IJGeometry = "{96eb9676-6530-11d1-977f-080036754203}"
Public Const AssemblyMembers1RelationshipCLSID = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}"
Public Const IJSurfaceArea = "{2A7D37F3-0440-4ACF-A731-90CE0613ABA0}"
Public Const IJGenericVolume = "{E897A3AF-E949-457C-968D-67E5DFDAD154}"
Public Const IJSmartOccurrence = "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}"
Public Const IStructInsulation = "{B75F82E2-CAEF-4C12-8BCB-281D004FDBDD}"
Public Const IStructHasInsulation = "{B18873BB-00E1-4c51-B3EA-023B70FC7AA5}"
Public Const ISPSMemberType = "{7B6EDBF9-E6B0-4016-AF41-0A0FB3C7AB6E}"
Public Const IJDPartClass = "{7B3E6F7F-93FB-11D1-BDDD-0060973D4805}"
Public Const IJLocalCoordinateSystem = "{704439B4-898E-450a-A45D-BBAEA04C2FA1}"
Public Const IJGraphicDataCache = "{3209E400-C632-4FB2-B6A4-0ADAEB53C848}"
Public Const IJDMemberObjects = "{9FC1AC01-9684-4e11-ABB8-6BDC3F636FE7}"
Public Const TOMEMBERS1 = "toMembers1"

Public Const INSULATIONSERVICESLIB_NAME = "InsulationSymbolServicesLib"
Public Const INSULATIONSERVICESLIB_PROGID = "SPSEncaseSymFuncs.ProfileCacheFuncs"
Public Const INSULATION_BO_PROGID = "StructInsulations.StructInsulation"

Public Const ENCASEMENTSYMBOL_OUTPUTNAME = "EncasementCrossSection"
Public Const ENCASEMENT_CROSSSECTION_REPNAME = "SimplePhysicalWireBody"     'output by the shapes
Public Const MEMBER_CROSSSECTION_REPNAME = "SimplePhysical"                 'input for programmed shape

Public Const RELATED_PARTS = "Part"
Public Const MEMBER_TYPECATGORY = "MemberTypeCategory"
Public Const CROSSSECTION_TYPE = "CrossSectionType"
Public Const INSULATION_FP_CRITERIA_IFACE = "IJUAStrMemberFPCriteria"

Public Const SETBACKREFERENCE_PARTSTART = 1
Public Const SETBACKREFERENCE_PARTEND = 2
Public Const SETBACKREFERENCE_AXISSTART = 3
Public Const SETBACKREFERENCE_AXISEND = 4

Public Const SETBACK_INTERFACE As String = "IJUAStructEncaseSetback"
Public Const SETBACK_DISTANCE1 As String = "SetbackDistance1"
Public Const SETBACK_DISTANCE2 As String = "SetbackDistance2"
Public Const SETBACK_REFERENCE1 As String = "SetbackReference1"
Public Const SETBACK_REFERENCE2 As String = "SetbackReference2"

Public Const INSULATIONERROR_TDLCODELISTNAME = "SPSTDLCodeListErrors"
Public Const INSULATIONERROR_UNEXPECTED = 130           'Insulation encountered an unexpected error.
Public Const INSULATIONERROR_MISSINGINPUTS = 131        'Insulation is missing required inputs.  Delete and replace.
Public Const INSULATIONERROR_BADATTRIBUTEVALUES = 132   'Insulation can not be computed with given attribute values.
Public Const INSULATIONERROR_ENCASEMENTSYMBOL = 133     'Insulation encountered a problem computing the encasement profile.
Public Const INSULATIONERROR_EXTRUSIONGEOMETRY = 134    'Insulation encountered a problem computing the extruded geometry.
Public Const INSULATIONERROR_CANNOTACCESSSYMBOL = 135   'Insulation could not access the encasement symbol.
Public Const INSULATIONERROR_NOSELECTION = 136          'No insulation encasement item found for given profile.  Select different encasement.
Public Const INSULATIONERROR_ENCASEMENTSELFINTERSECT = 137 'Insulation encasement profile geometry self-intersects.  Select different encasement.
Public Const INSULATIONERROR_BADSETBACKVALUES = 138     'Insulation setbacks exceed insulation length.  Reduce setback values.

Public Const INSULATION_THICKNESS_SYMBOL_PARAMETER = "thk_fp"

Public Const INSULATION_FP_CRITERIA = "IJUAStrMemberFPCriteria"
Public Const INSULATION_FP_CRITERIA_MEMBERTYPECATEGORY = "MemberTypeCategory"
Public Const INSULATION_FP_CRITERIA_CROSSSECTIONCLASS = "CrossSectionType"
Public Const INSULATION_FP_CRITERIA_CROSSSECTIONCLASSANY = "*"

Private Const MODULE = "MemberEncasementCommon"


Public Function GetRefCollection(pSO As Object) As IJDReferencesCollection
    Const METHOD = "GetRefCollection"
    On Error GoTo ErrorHandler
  
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Set pRelationHelper = pSO
    On Error Resume Next
    Set pCollectionHelper = pRelationHelper.CollectionRelations(IJSmartOccurrence, "toArgs_O")
    On Error Resume Next
    Set GetRefCollection = pCollectionHelper.Item("RC")
      
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function GetAttributeCollection(oBO As Object, attrInterface As String) As Object
    Const METHOD = "GetAttributeCollection"
    On Error GoTo ErrorHandler
    Dim pIJAttrbs As IJDAttributes
    
    If Not oBO Is Nothing Then
        Set pIJAttrbs = oBO
        On Error Resume Next
        Set GetAttributeCollection = pIJAttrbs.CollectionOfAttributes(attrInterface)
        Err.Clear
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function GetAttributeValue(oAttrCollection As CollectionProxy, attrName As String, ByRef outAttrValue As Variant) As Boolean
    Const METHOD = "GetAttributeValue"
    On Error GoTo ErrorHandler
    
    Dim attrValueEmpty As Variant

    Dim oAttr As IJDAttribute

    outAttrValue = attrValueEmpty        'set output to empty.
    GetAttributeValue = False

    If Not oAttrCollection Is Nothing Then
        On Error Resume Next
        Set oAttr = oAttrCollection.Item(attrName)
        If Err.Number = 0 Then
            outAttrValue = oAttr.Value
            GetAttributeValue = True
        Else
            Err.Clear
        End If
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub AddRelationship(oObj1 As Object, strInterface As String, oObj2 As Object, strCollection As String, strRelationName As String, bReplace As Boolean)
    Const METHOD = "AddRelationship"
    On Error GoTo ErrorHandler

   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = oObj1
    Set pCollectionHelper = pRelationHelper.CollectionRelations(strInterface, strCollection)

    If pCollectionHelper.count > 0 Then
        On Error Resume Next
        Set pRelationshipHelper = pCollectionHelper.GetRelationshipToTarget(oObj2)
        Err.Clear
        On Error GoTo ErrorHandler
        If bReplace = True Then
            If pRelationshipHelper Is Nothing Then     ' remove the relationship to a different object
                Dim i As Long, lCount As Long
                lCount = pCollectionHelper.count
                For i = lCount To 1 Step -1
                    pCollectionHelper.Remove (i)
                Next
            End If
        End If
    End If

    If pRelationshipHelper Is Nothing Then
        pCollectionHelper.Add oObj2, strRelationName, pRelationshipHelper
        Set pRevision = New JRevision
        pRevision.AddRelationship pRelationshipHelper
    End If
  
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub DefineEncasementSymbolInputs(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "SetSymbolInputs"
    On Error GoTo ErrorHandler
    'Define the inputs
    
    Dim lLibCookie As Long
    Dim lMethodCookie As Long
    Dim LibDesc As DLibraryDescription
    Dim LibDescCustomMethods As IJDLibraryDescription
    Dim methodDesc As IJDMethodDescription

    ' Set the reference of the InsulationServices library
    Set LibDesc = pSymbolDefinition.IJDUserMethods.GetLibrary(INSULATIONSERVICESLIB_NAME)
    If LibDesc Is Nothing Then
        Set LibDescCustomMethods = New DLibraryDescription
        
        LibDescCustomMethods.Name = INSULATIONSERVICESLIB_NAME
        LibDescCustomMethods.Type = imsLIBRARY_IS_ACTIVEX
        LibDescCustomMethods.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
        LibDescCustomMethods.Source = INSULATIONSERVICESLIB_PROGID
        pSymbolDefinition.IJDUserMethods.SetLibrary LibDescCustomMethods
        Set LibDescCustomMethods = Nothing
        Set LibDesc = pSymbolDefinition.IJDUserMethods.GetLibrary(INSULATIONSERVICESLIB_NAME)
    End If

    If LibDesc Is Nothing Then
        WriteError METHOD, "IJUserMethods.GetLibrary failed for " & INSULATIONSERVICESLIB_NAME
        Exit Sub
    End If

    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_SHARED
    
    lLibCookie = LibDesc.Cookie

    Dim oInput As IJDInput
    Dim pIJDInputs As IJDInputs
    
    Set pIJDInputs = pSymbolDefinition.IJDInputs

    'input 1 is the InsulationSpec object
    'CMCacheSpecThickness converts it to ParameterContent whose value is thickness
    Set oInput = New DInput
    oInput.Name = "InsulationSpecThickness"
    oInput.Description = "Insulation Specification Thickness"
    oInput.Properties = igINPUT_IS_CACHED
    oInput.index = 1

    lMethodCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheSpecThickness", lLibCookie)
    oInput.IJDInputStdCustomMethod.SetCMCache lLibCookie, lMethodCookie
    pIJDInputs.Add oInput

    'input 2 is the Material object
    'CMCacheMaterialDensity converts it to ParameterContent whose value is a fixed value
    'since the material does not affect the geometry of the profile
    oInput.Reset
    oInput.Name = "InsulationMaterial"
    oInput.Description = "Insulation Material"
    oInput.Properties = igINPUT_IS_CACHED
    oInput.index = 2

    lMethodCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheMaterial", lLibCookie)
    oInput.IJDInputStdCustomMethod.SetCMCache lLibCookie, lMethodCookie
    pIJDInputs.Add oInput
    
    'input 3 is the MemberPartPrismatic object
    'CMCacheMemberCrossSection converts it to ParameterContent whose value is the oid of the flavor of the symbol of the member
    oInput.Reset
    oInput.Name = "MemberPartCrossSection"
    oInput.Description = "Member Part CrossSection"
    oInput.Properties = igINPUT_IS_CACHED
    oInput.index = 3

    lMethodCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheMemberCrossSection", lLibCookie)
    oInput.IJDInputStdCustomMethod.SetCMCache lLibCookie, lMethodCookie
    pIJDInputs.Add oInput
    
    'input 4 is the MemberPartPrismatic object as ISPSMemberType interface
    'CMCacheMemberType converts it to ParameterContent whose value is a fixed value
    'since the member type does not affect the geometry of the profile, although it can affect the rule.
    oInput.Reset
    oInput.Name = "MemberPartType"
    oInput.Description = "Member Part Type"
    oInput.Properties = igINPUT_IS_CACHED
    oInput.index = 4

    lMethodCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheMemberType", lLibCookie)
    oInput.IJDInputStdCustomMethod.SetCMCache lLibCookie, lMethodCookie
    pIJDInputs.Add oInput

    pIJDInputs.Property = igCOLLECTION_VARIABLE

''''    Set methodDesc = New DMethodDescription
''''    methodDesc.Name = "GetReferences"
''''    methodDesc.Description = "GetReferences"
''''    methodDesc.Library = lLibCookie
''''    methodDesc.Cookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMGetReferences", lLibCookie)
''''    methodDesc.Properties = imsMETHOD_OVERRIDE
''''    pSymbolDefinition.IJDUserMethods.SetMethod methodDesc
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub WriteError(strSource As String, strMessage As String)
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.Add 0, strSource, strMessage
    End If
    Set oEditErrors = Nothing
End Sub

Public Sub CopyValuesFromItemToSO(soCol As CollectionProxy, itmCol As CollectionProxy)
Const METHOD = "CopyValuesFromItemToSO"
On Error GoTo ErrorHandler

    Dim i As Integer
   
    If soCol.count <> itmCol.count Then
        Exit Sub
    End If
    
    For i = 1 To soCol.count
        If soCol.Item(i).AttributeInfo.Name = itmCol.Item(i).AttributeInfo.Name Then
            soCol.Item(i).Value = itmCol.Item(i).Value
        End If
    Next i
     
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Function IsSOOverridden(attrCol As CollectionProxy) As Boolean
Const METHOD = "IsSOOverridden"
On Error GoTo ErrorHandler

    Dim i As Integer
    Dim oAttr As IJDAttribute
    
    IsSOOverridden = False
    For i = 1 To attrCol.count
      Set oAttr = attrCol.Item(i)
      If Not IsEmpty(oAttr.Value) Then
          IsSOOverridden = True
          Set oAttr = Nothing
          Exit For
      End If
      Set oAttr = Nothing
    Next i
 
 Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

' Declared as integer but will be returning StructInsulationGraphicInputHelperStatus
Public Function ValidateMemberFPCriteria(oGraphicObject As Object, oDefinition As Object) As Integer
Const METHOD = "ValidateMemberFPCriteria"
On Error GoTo ErrorHandler

    ValidateMemberFPCriteria = StructInsulationInputHelper_Unknown

    If oGraphicObject Is Nothing Then
        ValidateMemberFPCriteria = StructInsulationInputHelper_UnexpectedError

    ElseIf TypeOf oGraphicObject Is ISPSMemberPartPrismatic Then
        Dim oMemberPart As ISPSMemberPartPrismatic
        Dim oAttrColl As Object
        Dim vAttrValue As Variant
        Dim lDefTypeCategory As Long
        Dim strDefCrossSectionType As String

        Set oMemberPart = oGraphicObject

        Set oAttrColl = GetAttributeCollection(oDefinition, INSULATION_FP_CRITERIA)
        If oAttrColl Is Nothing Then
            ValidateMemberFPCriteria = StructInsulationInputHelper_Ok     'not using Insulation Criteria interface
        Else
            ' Validate member type
            If GetAttributeValue(oAttrColl, INSULATION_FP_CRITERIA_MEMBERTYPECATEGORY, vAttrValue) Then
                lDefTypeCategory = vAttrValue
                If lDefTypeCategory = oMemberPart.MemberType.TypeCategory Then
                    ValidateMemberFPCriteria = StructInsulationInputHelper_Ok
                Else
                ' wrong fireproofing chosen for the member type
                    ValidateMemberFPCriteria = StructInsulationInputHelper_IncompatibleEncasement
                End If
            Else
                'no MemberTypeCategory attribute
                ValidateMemberFPCriteria = StructInsulationInputHelper_BadAttributeValues
            End If
            
            ' validate cross section
            If ValidateMemberFPCriteria = StructInsulationInputHelper_Ok Then
                If GetAttributeValue(oAttrColl, INSULATION_FP_CRITERIA_CROSSSECTIONCLASS, vAttrValue) Then
                    If VarType(vAttrValue) = vbEmpty Then
                    ' not specified
                        ValidateMemberFPCriteria = StructInsulationInputHelper_Ok
                    Else
                        strDefCrossSectionType = vAttrValue
                        If strDefCrossSectionType = "" Then
                            ' not specified
                            ValidateMemberFPCriteria = StructInsulationInputHelper_Ok
                        ElseIf strDefCrossSectionType = INSULATION_FP_CRITERIA_CROSSSECTIONCLASSANY Then
                            ValidateMemberFPCriteria = StructInsulationInputHelper_Ok
                        ElseIf LCase(strDefCrossSectionType) = LCase(oMemberPart.CrossSection.SectionType) Then
                            ValidateMemberFPCriteria = StructInsulationInputHelper_Ok
                        Else
                            ' wrong insulation for the type of cross section
                            ValidateMemberFPCriteria = StructInsulationInputHelper_IncompatibleCrossSection
                        End If
                    End If
                Else
                    ' no CrossSectionType attribute found
                    ValidateMemberFPCriteria = StructInsulationInputHelper_BadAttributeValues
                End If
            End If
        End If
    Else
    ' Not a member
        ValidateMemberFPCriteria = StructInsulationInputHelper_InvalidTypeOfObject
    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

