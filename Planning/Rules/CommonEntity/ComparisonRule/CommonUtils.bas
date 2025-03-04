Attribute VB_Name = "CommonUtils"
Option Explicit


Public sSOURCEFILE As String
Public m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Public m_oError As IJEditError     ' Defined here for convenience
Public Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Public Const PI As Double = 3.14159265358979

Public m_dDistanceTolerance     As Double
Public m_dThicknessTolerance    As Double
Public m_dAreaTolerance         As Double
Public m_dVolumeTolerance       As Double
Public m_dWeightTolerance       As Double
Public m_bflippingpart          As Double
Public m_dAngleTolerance        As Double

Public m_bStandardPartRule      As Boolean
Public m_bGeometriesCompared    As Boolean

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Public Function AreConnectedFeaturesCommon(oMatchingPairs As IJElements, oTransMatrix As IJDT4x4, Optional bCompareAttr As Boolean = True) As Boolean
Const METHOD = "AreConnectedFeaturesCommon"
On Error GoTo ErrorHandler
    
    Dim oPlnMatchingPair        As IJPlnMatchingPair
    Dim oCndtFeatDefAttribs     As IJDAttributesCol
    Dim oTgtFeatDefAttribs      As IJDAttributesCol
    Dim bAreCommon              As Boolean
    
    bAreCommon = True

    If Not oMatchingPairs Is Nothing Then
        
        'Return False if atleast one match is partial or unique
        For Each oPlnMatchingPair In oMatchingPairs
            If oPlnMatchingPair.TypeOfMatch = PlnMatch_Unique Or oPlnMatchingPair.TypeOfMatch = PlnMatch_Partial Then
                AreConnectedFeaturesCommon = False
                Exit Function
            End If
        Next
    
        If bCompareAttr = True Then
            'Loop through the pairs
            For Each oPlnMatchingPair In oMatchingPairs
            
                'If the definition is not the same, return False
                If Not oPlnMatchingPair.CandidateDefinition Is oPlnMatchingPair.TargetDefinition Then
                    bAreCommon = False
                    Exit For
                Else
                    'If the number of definition attributes is not the same, return False
                    oPlnMatchingPair.GetDefinitionAttributes oCndtFeatDefAttribs, oTgtFeatDefAttribs
                    
                    bAreCommon = CompareAttributes(oCndtFeatDefAttribs, oTgtFeatDefAttribs)
                End If
                
                'Exit if we already found a mis-match
                If bAreCommon = False Then
                    Exit For
                End If
            Next
        End If
    End If

    AreConnectedFeaturesCommon = bAreCommon
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function ArePCMountingAnglesCommon(oMatchingPairs As IJElements, oTransMatrix As IJDT4x4) As Boolean
Const METHOD = "ArePCMountingAnglesCommon"
On Error GoTo ErrorHandler
    
    Dim oPhysConnPair       As IJPlnPhysConnPair
    Dim bAreCommon          As Boolean
    Dim dCndtMtAngle1       As Double
    Dim dTgtMtAngle1        As Double
    Dim dCndtMtAngle2       As Double
    Dim dTgtMtAngle2        As Double
    
    bAreCommon = True

    If Not oMatchingPairs Is Nothing Then
        'Exit if atleast one PC does not have a matching PC on the other part
        For Each oPhysConnPair In oMatchingPairs
            If oPhysConnPair.TypeOfMatch = PlnMatch_Partial Or oPhysConnPair.TypeOfMatch = PlnMatch_Unique Then
                bAreCommon = False
                Exit For
            End If
        Next
        
        'Loop through PCs, compare mounting angles
        For Each oPhysConnPair In oMatchingPairs
            oPhysConnPair.GetMountingAngles dCndtMtAngle1, dCndtMtAngle2, dTgtMtAngle1, dTgtMtAngle2

            If Abs(dCndtMtAngle1 - dTgtMtAngle1) > m_dAngleTolerance Or _
               Abs(dCndtMtAngle2 - dTgtMtAngle2) > m_dAngleTolerance Then
                
                bAreCommon = False
                Exit For
            End If
        Next
    End If

    ArePCMountingAnglesCommon = bAreCommon
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function GetPCMountingAngles(oMatchingPairs As IJElements, oTransMatrix As IJDT4x4) As String
Const METHOD = "GetPCMountingAngles"
On Error GoTo ErrorHandler

    Dim oPhysConnPair       As IJPlnPhysConnPair
    Dim bAreCommon          As Boolean
    Dim dCndtMtAngle1       As Double
    Dim dTgtMtAngle1        As Double
    Dim dCndtMtAngle2       As Double
    Dim dTgtMtAngle2        As Double

    bAreCommon = True
    Dim Cstrangel As String
    Dim Tstrangel As String
    If Not oMatchingPairs Is Nothing Then
        'Exit if atleast one PC does not have a matching PC on the other part

        'Loop through PCs, compare mounting angles
         For Each oPhysConnPair In oMatchingPairs
        
            If Not oPhysConnPair.TypeOfMatch = PlnMatch_Full Then
                Cstrangel = "PC Geometries are not common"
                Tstrangel = "PC Geometries are not common"
                
                GetPCMountingAngles = Cstrangel + "@" + Tstrangel
                Exit Function
            End If
        Next
        
        For Each oPhysConnPair In oMatchingPairs
            oPhysConnPair.GetMountingAngles dCndtMtAngle1, dCndtMtAngle2, dTgtMtAngle1, dTgtMtAngle2
                Dim arr() As String
                arr = Split(m_dDistanceTolerance, GetDecimalSeparator)
                dCndtMtAngle1 = FormatNumber(dCndtMtAngle1, Len(arr(1)))
                dCndtMtAngle2 = FormatNumber(dCndtMtAngle2, Len(arr(1)))
                dTgtMtAngle1 = FormatNumber(dTgtMtAngle1, Len(arr(1)))
                dTgtMtAngle2 = FormatNumber(dTgtMtAngle2, Len(arr(1)))
               
               Cstrangel = Cstrangel + CStr(dCndtMtAngle1) + " - " + CStr(dCndtMtAngle2) + " ; "
               Tstrangel = Tstrangel + CStr(dTgtMtAngle1) + " - " + CStr(dTgtMtAngle2) + " ; "
        Next
    
    End If
    GetPCMountingAngles = Cstrangel + "@" + Tstrangel
    
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function AreConnectedMfgObjectsCommon(oMatchingPairs As IJElements, oTransMatrix As IJDT4x4) As Boolean
Const METHOD = "AreConnectedObjectsCommon"
On Error GoTo ErrorHandler

    'Just call the method used to compare features
    AreConnectedMfgObjectsCommon = AreConnectedFeaturesCommon(oMatchingPairs, oTransMatrix)

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function GetRelatedObjects _
                 (oGivenObject, strGivenObjIfaceGUID, strInterestedRoleName) As IJElements
Const METHOD = "GetRelatedObjects"
On Error GoTo ErrorHandler

    Dim oAssocRel               As IJDAssocRelation
    Dim oTargetObjCol           As IJDTargetObjectCol
    Dim oResultColl             As IJElements
    Dim oProductionRouting      As IJDProductionRouting
    Dim oRoutingActionColl      As IJElements
    
    Set oResultColl = New JObjectCollection
    
    Set oAssocRel = oGivenObject
    Set oTargetObjCol = oAssocRel.CollectionRelations(strGivenObjIfaceGUID, strInterestedRoleName)
    
    If oTargetObjCol.Count > 0 Then
        Dim i As Long
        
        For i = 1 To oTargetObjCol.Count
            oResultColl.Add oTargetObjCol.Item(i)
        Next
    End If
    
    Set GetRelatedObjects = oResultColl

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function getAttrValue(strAttrName As String, oObj As Object) As String

    Dim oSmartOccurrence                As IJSmartOccurrence
    Dim oSmartItem                      As IJSmartItem
    Dim osmartclass                     As IJSmartClass
    Dim strWeldName                     As String
       
    Set oSmartOccurrence = oObj
    
    If Not oSmartOccurrence Is Nothing Then
        Set oSmartItem = oSmartOccurrence.ItemObject
        If Not oSmartItem Is Nothing Then
            strWeldName = oSmartItem.Name
            Set osmartclass = oSmartItem.Parent
        End If
    End If
    
    Dim oAttrHelper                     As IJDAttributes
    Dim oAttribute                      As Variant
    Dim oUserAttrColl                   As IJDAttributesCol
    Dim varInterfaceID                  As Variant
    
    Set oAttrHelper = oSmartOccurrence
    If oAttrHelper Is Nothing Then Exit Function
        
    For Each varInterfaceID In oAttrHelper
        
        Dim oparentsc As IJSmartClass
        
        On Error Resume Next
        Set oparentsc = osmartclass.Parent
        
        If Not oparentsc Is Nothing Then
            Set oUserAttrColl = oAttrHelper.CollectionOfAttributes(oparentsc.SelectionRuleInterface)
        End If
        
        If oUserAttrColl Is Nothing Then
            Set oUserAttrColl = oAttrHelper.CollectionOfAttributes(osmartclass.SelectionRuleInterface)
        End If

        For Each oAttribute In oUserAttrColl
            If Not oAttribute Is Nothing Then
                If oAttribute.AttributeInfo.Name = strAttrName Then
                    If Not IsEmpty(oAttribute.Value) Then
                        getAttrValue = oAttribute.Value
                        getAttrValue = GetStringFromId(oAttribute.AttributeInfo.CodeListTableName, CLng(getAttrValue))
                        GoTo CleanUp
                    End If
                End If
            End If
         Next
    Next
                           
CleanUp:
    Set oAttrHelper = Nothing
    Set oAttribute = Nothing
End Function

Public Function CompareAttributes(oCandAttribs As IJDAttributesCol, oTargAttribs As IJDAttributesCol) As Boolean
Const METHOD = "CompareAttributes"
On Error GoTo ErrorHandler

    CompareAttributes = True
    
    Dim oMatchedAttrbutes       As IJElements
    Dim oCndtAttrib             As IJDAttribute
    Dim oTgtAttrib              As IJDAttribute
    Dim bAttributeFound         As Boolean
    Dim i                       As Long
    Dim j                       As Long
        
    'Compare all the attributes.
    If Not oCandAttribs Is Nothing And Not oTargAttribs Is Nothing Then
        If oCandAttribs.Count <> oTargAttribs.Count Then
            CompareAttributes = False
        Else
            'Create a collection to hold matched attributes
            Set oMatchedAttrbutes = New JObjectCollection
            
            'Loop through the candidate definition attributes
             For i = 1 To oCandAttribs.Count
             
                'Initialize a variable to tell us if the same attribute is available on the target
                bAttributeFound = False
                Set oCndtAttrib = oCandAttribs.Item(i)
                
                If Not oCndtAttrib Is Nothing Then
                
                    'Lopp through the target attributes
                    For j = 1 To oTargAttribs.Count
                        Set oTgtAttrib = oTargAttribs.Item(j)
                        
                        If Not oTgtAttrib Is Nothing Then
                        
                            'Continue only if the attribute is not already matched
                            If oMatchedAttrbutes.Contains(oTgtAttrib) = False Then
                                
                                'If the attribute name is the same, we found a match
                                If oCndtAttrib.AttributeInfo.Name = oTgtAttrib.AttributeInfo.Name Then
                                    oMatchedAttrbutes.Add oTgtAttrib
                                    bAttributeFound = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                Else
                    bAttributeFound = True
                End If
                
                'If we did not find a matching attribute with the same name, treat the connected objects as uncommon
                If bAttributeFound = False Then
                    CompareAttributes = False
                    Exit For
                'Else, if the attribute values are not the same treat the connected objects as uncommon
                ElseIf Not oCndtAttrib Is Nothing And Not oTgtAttrib Is Nothing Then
                    If oCndtAttrib.Value <> oTgtAttrib.Value Then
                        CompareAttributes = False
                        Exit For
                    End If
                End If
            Next i
        End If
    End If

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function GetDryWeight(oCandidate As IJAssembly, oTarget As IJAssembly, ByRef vCndtPropValue As Variant, ByRef vTgtPropValue As Variant) As Boolean

    Dim oAssemblyBase As IJWeightCGgrouping
        
    Set oAssemblyBase = oCandidate
    oAssemblyBase.UpdateWCG
    vCndtPropValue = CVar(oAssemblyBase.DryWeight)
    
    Set oAssemblyBase = Nothing
    
    Set oAssemblyBase = oTarget
    oAssemblyBase.UpdateWCG
    vTgtPropValue = CVar(oAssemblyBase.DryWeight)
    
    GetDryWeight = True

End Function

Public Function GetWetWeight(oCandidate As IJAssembly, oTarget As IJAssembly, ByRef vCndtPropValue As Variant, ByRef vTgtPropValue As Variant) As Boolean

    Dim oAssemblyBase As IJWeightCGgrouping
    
    Set oAssemblyBase = oCandidate
    oAssemblyBase.UpdateWCG
    vCndtPropValue = oAssemblyBase.WetWeight
    
    Set oAssemblyBase = Nothing
    
    Set oAssemblyBase = oTarget
    oAssemblyBase.UpdateWCG
    vTgtPropValue = oAssemblyBase.WetWeight

    GetWetWeight = True
End Function

Public Function GetZAxis(oCandidate As IJAssembly, oTarget As IJAssembly, ByRef oCndVector As IJDVector, ByRef oTgtVector As IJDVector) As Boolean

    Dim oAssemblyBase As IJLocalCoordinateSystem
    
    Set oAssemblyBase = oCandidate
    Set oCndVector = oAssemblyBase.ZAxis

    Set oAssemblyBase = Nothing
    
    Set oAssemblyBase = oTarget
    Set oTgtVector = oAssemblyBase.ZAxis
    
    GetZAxis = True
End Function

Public Function GetShortString(oObj As Object, strAttName As String, lVal As Long) As String
Const METHOD As String = "FillPartCoatingAreaInfo"
On Error GoTo ErrorHandler
   
    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim oCodelist                   As IJDCodeListMetaData
    Dim oUserAttrColl                   As IJDAttributesCol
    Dim varInterfaceID                  As Variant
    
    Set oCodelist = GetPOM("Model")
                    
    Set oAttributeMetadata = oObj
    Set oAttrHelper = oObj
    
  Dim iID                             As Variant
    Dim ii As Long
    If oAttributeMetadata Is Nothing Then Exit Function
    
    If oAttrHelper Is Nothing Then Exit Function
       
    
    For Each varInterfaceID In oAttrHelper
        Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(varInterfaceID)

        iID = oInterfaceInfo.Type
        Set oUserAttrColl = oAttrHelper.CollectionOfAttributes(iID)

        For ii = 1 To oUserAttrColl.Count
            Set oAttribute = oUserAttrColl.Item(ii)
            If Not oAttribute Is Nothing Then
                If oAttribute.AttributeInfo.UserName = strAttName Then
                    GetShortString = GetStringFromId(oAttribute.AttributeInfo.CodeListTableName, lVal)
                    GoTo CleanUp
                End If
            End If
         Next ii

    Next
                           
CleanUp:

    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing
    Set oCodelist = Nothing
        
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Private Function GetPOM(strDBType As String) As IJDPOM
Const METHOD = "GetPOM"
On Error GoTo ErrHandler
    
    Dim oContext            As IJContext
    Dim oAccessMiddle       As IJDAccessMiddle
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDBType)
    
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
    
Exit Function
ErrHandler:
     Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function GetStringFromId(ByVal TableName As String, ByVal ValueID As Long) As String
Const METHOD = "GetStringFromId"
On Error GoTo ErrorHandler

    Dim oIJDCodeListMetaData      As IJDCodeListMetaData
    Set oIJDCodeListMetaData = GetPOM("Model")
    GetStringFromId = oIJDCodeListMetaData.DisplayStringByValueID(TableName, ValueID)
    
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Public Function GetPropertyValues(ByRef oCandidate As Object, ByRef vCndtPropValue As Variant, ByRef vTgtPropValue As Variant, ByRef bAPISuccess As Boolean _
                                    , ByRef bCommon As Boolean, ByRef strKeywordChecks As String, ByRef oCPSCommonHelper As IJPlnCompareHelperEx _
                                    , Optional ByRef bIsGeomSame As Boolean, Optional oConnectedObjPairs As IJElements, Optional oMatchedPCs As IJElements) As String
 
    Const METHOD = "GetPropertyValues"
    On Error GoTo ErrorHandler
    
        Dim bStatus                 As String
        Dim arr()                   As String
        Dim oCndtConnectedObjects       As IJElements
        Dim oTgtConnectedObjects        As IJElements
        
        Select Case strKeywordChecks

            Case "Area", "SurfaceArea"
                arr = Split(m_dAreaTolerance, GetDecimalSeparator)
                vCndtPropValue = FormatNumber(vCndtPropValue, Len(arr(1)))
                vTgtPropValue = FormatNumber(vTgtPropValue, Len(arr(1)))
                
            Case "Thickness"
                arr = Split(m_dThicknessTolerance, GetDecimalSeparator)
                vCndtPropValue = FormatNumber(vCndtPropValue, Len(arr(1)))
                vTgtPropValue = FormatNumber(vTgtPropValue, Len(arr(1)))
                
             Case "BoundaryLength", "Width", "Length", "CutLength", "LandingCurveLength"
                arr = Split(m_dDistanceTolerance, GetDecimalSeparator)
                vCndtPropValue = FormatNumber(vCndtPropValue, Len(arr(1)))
                vTgtPropValue = FormatNumber(vTgtPropValue, Len(arr(1)))
                
            Case "DryWeight", "WetWeight"
                arr = Split(m_dWeightTolerance, GetDecimalSeparator)
                vCndtPropValue = FormatNumber(vCndtPropValue, Len(arr(1)))
                vTgtPropValue = FormatNumber(vTgtPropValue, Len(arr(1)))
                                
            Case "ProfileType", "Curved", "PlateType", "Type", "Priority"
                vCndtPropValue = GetShortString(oCandidate, strKeywordChecks, CLng(vCndtPropValue))
                vTgtPropValue = GetShortString(oCandidate, strKeywordChecks, CLng(vTgtPropValue))
                
            Case "PrimaryOrientation", "SecondaryOrientation"
                Dim Orientation() As String
                Orientation = Split(vCndtPropValue, "/")
                arr = Split(m_dDistanceTolerance, GetDecimalSeparator)
                vCndtPropValue = FormatNumber(Orientation(0), Len(arr(1))) + "/ " + FormatNumber(Orientation(1), Len(arr(1))) + "/ " + FormatNumber(Orientation(2), Len(arr(1)))
                
                Orientation = Split(vTgtPropValue, "/")
                vTgtPropValue = FormatNumber(Orientation(0), Len(arr(1))) + "/ " + FormatNumber(Orientation(1), Len(arr(1))) + "/ " + FormatNumber(Orientation(2), Len(arr(1)))
                
            Case "BuildMethod"
                If Not IsEmpty(vCndtPropValue) Then vCndtPropValue = GetShortString(oCandidate, "Build Method", CLng(vCndtPropValue)) Else vCndtPropValue = ""
                If Not IsEmpty(vTgtPropValue) Then vTgtPropValue = GetShortString(oCandidate, "Build Method", CLng(vTgtPropValue)) Else vTgtPropValue = ""

            Case "SlotConnectivity"
                If Not IsEmpty(vCndtPropValue) Then vCndtPropValue = GetShortString(oCandidate, "Slot Connectivity", CLng(vCndtPropValue)) Else vCndtPropValue = ""
                If Not IsEmpty(vTgtPropValue) Then vTgtPropValue = GetShortString(oCandidate, "Slot Connectivity", CLng(vTgtPropValue)) Else vTgtPropValue = ""
                                
            Case "StageCode", "WorkCenter", "SectionName", "SectionStandard", "AssemblyShrinkages"
                If IsEmpty(vCndtPropValue) Or IsNull(vCndtPropValue) Then vCndtPropValue = ""
                If IsEmpty(vTgtPropValue) Or IsNull(vTgtPropValue) Then vTgtPropValue = ""

            Case "Direction"
                If vCndtPropValue = 0 Then vCndtPropValue = "Above"
                If vCndtPropValue = 1 Then vCndtPropValue = "Below"
                If vCndtPropValue = 3 Then vCndtPropValue = "Inboard"
                If vCndtPropValue = 4 Then vCndtPropValue = "Outboard"
                If vTgtPropValue = 0 Then vTgtPropValue = "Above"
                If vTgtPropValue = 1 Then vTgtPropValue = "Below"
                If vTgtPropValue = 3 Then vTgtPropValue = "Inboard"
                If vTgtPropValue = 4 Then vTgtPropValue = "Outboard"
                
            Case "MountingAngles"
                Dim MatchingPairs           As Object
                Dim tgt1                    As Double
                Dim tgt2                    As Double
                Dim totcount                As Integer
                
                If Not bAPISuccess Then bCommon = False
                If Not bIsGeomSame Then
                    vCndtPropValue = "Part Geometries are not common"
                    vTgtPropValue = "Part Geometries are not common"
                    bCommon = False

                Else
                    Dim GetMountingAngles() As String
                    
                    GetMountingAngles = Split(GetPCMountingAngles(oMatchedPCs, oCPSCommonHelper.TransMatrix), "@")
                    vCndtPropValue = GetMountingAngles(0)
                    vTgtPropValue = GetMountingAngles(1)
                End If
                
            Case "ConnectionBevels"
                Dim oConnectedItem          As Object
                Dim oPlnMatchingPair        As IJPlnPhysConnPair
                Dim CandCategoryValue       As String
                Dim TgtCategoryValue        As String
                Dim x                       As Integer
                Dim AttrValue               As String
                
                If Not bAPISuccess Then bCommon = False
                
                vCndtPropValue = ""
                vTgtPropValue = ""
                
                'Code To Display bevel information of Not bIsGeomSame Parts
                 If Not bIsGeomSame Then

                   vCndtPropValue = "Part Geometries are not common"
                   vTgtPropValue = "Part Geometries are not common"
                   bCommon = False

                 Else
                 'Code To Display bevel information of bIsGeomSame Parts
                    If oMatchedPCs.Count > 0 Then
                        
                        Dim oPhysConnPair       As IJPlnPhysConnPair
                        Dim bPCGeometryCommon As Boolean
                        bPCGeometryCommon = True
                        
                        For Each oPhysConnPair In oMatchedPCs
                            If Not oPhysConnPair.TypeOfMatch = PlnMatch_Full Then
                                vCndtPropValue = "PC Geometries are not common"
                                vTgtPropValue = "PC Geometries are not common"
                                
                                bPCGeometryCommon = False
                                Exit For
                            End If
                        Next
                                
                        If bPCGeometryCommon = True Then
                            For x = 1 To oMatchedPCs.Count
                                Set oPlnMatchingPair = oMatchedPCs.Item(x)
                                If Not oPlnMatchingPair.CandidatePhysConn Is Nothing Then
                                    AttrValue = getAttrValue("Category", oPlnMatchingPair.CandidatePhysConn)
                                    If AttrValue = "" Then CandCategoryValue = CandCategoryValue + AttrValue + "" Else CandCategoryValue = CandCategoryValue + AttrValue + "; "
                                    vCndtPropValue = CandCategoryValue
                                Else
                                    vCndtPropValue = vCndtPropValue + ""
                                End If
            
                               If Not oPlnMatchingPair.TargetPhysConn Is Nothing Then
                                    AttrValue = getAttrValue("Category", oPlnMatchingPair.TargetPhysConn)
                                    If AttrValue = "" Then TgtCategoryValue = TgtCategoryValue + AttrValue + "" Else TgtCategoryValue = TgtCategoryValue + AttrValue + "; "
                                    vTgtPropValue = TgtCategoryValue
                                Else
                                    vTgtPropValue = vTgtPropValue + ""
                               End If
            
                            Next
                        End If
                    End If
                End If
                
            Case "DefinitionAttributes", "ProfileSolid", "Geometry", "PartsGeometry", "Boundary"
                If bCommon Then
                    vCndtPropValue = "Same"
                    vTgtPropValue = "Same"
                Else
                    vCndtPropValue = "Different"
                    vTgtPropValue = "Different"
                End If
                      
             Case "EdgeFeatures", "CornerFeatures", "FreeEdgeTreatments", "MarkingLines", "Shrinkages", "Margins", "Features"
                If Not bIsGeomSame Then
                    bAPISuccess = oCPSCommonHelper.GetConnectedObjects(GeteObject(strKeywordChecks), oCndtConnectedObjects, oTgtConnectedObjects)
                    If oCndtConnectedObjects.Count > 0 Or oTgtConnectedObjects.Count > 0 Then
                        vCndtPropValue = "0" + " -Common; " + CStr(oCndtConnectedObjects.Count) + " -Unique"
                        vTgtPropValue = "0" + " -Common; " + CStr(oTgtConnectedObjects.Count) + " -Unique"
                        bCommon = False
                    Else
                        vCndtPropValue = "0" + " -Common; " + "0" + " -Unique"
                        vTgtPropValue = "0" + " -Common; " + "0" + " -Unique"
                    End If
                Else
                    Dim ConnectedObjPairsCount As Integer
                    Dim PlnMatchingPairItem As IJPlnMatchingPair
                    Dim CommonCount As Integer
                    Dim CanCount As Integer
                    Dim TgtCount As Integer
                    CommonCount = 0
                    CanCount = 0
                    TgtCount = 0
                    If Not oConnectedObjPairs Is Nothing Then
                    
                        If oConnectedObjPairs.Count > 0 Then
                            For ConnectedObjPairsCount = 1 To oConnectedObjPairs.Count
                            
                                Set PlnMatchingPairItem = oConnectedObjPairs.Item(ConnectedObjPairsCount)
                                If (Not PlnMatchingPairItem.CandidateConnObject Is Nothing) And (Not PlnMatchingPairItem.TargetConnObject Is Nothing) Then
                                    CommonCount = CommonCount + 1
                                ElseIf Not PlnMatchingPairItem.CandidateConnObject Is Nothing Then
                                    CanCount = CanCount + 1
                                ElseIf Not PlnMatchingPairItem.TargetConnObject Is Nothing Then
                                    TgtCount = TgtCount + 1
                                End If
                            Next ConnectedObjPairsCount
                        End If
                    End If
                    vCndtPropValue = CStr(CommonCount) + " -Common; " + CStr(CanCount) + " -Unique"
                    vTgtPropValue = CStr(CommonCount) + " -Common; " + CStr(TgtCount) + " -Unique"
                End If
                               
            'Plate Cases
            Case "Tightness"
                vCndtPropValue = GetShortString(oCandidate, "PlateTightness", CLng(vCndtPropValue))
                vTgtPropValue = GetShortString(oCandidate, "PlateTightness", CLng(vTgtPropValue))
                
            'Member Cases
            Case "TypeCategory"
                vCndtPropValue = GetShortString(oCandidate, "Type Category", CLng(vCndtPropValue))
                vTgtPropValue = GetShortString(oCandidate, "Type Category", CLng(vTgtPropValue))
                
            'StdAssembly Cases
            Case "AssemblyType"
                  vCndtPropValue = GetShortString(oCandidate, "Type", CLng(vCndtPropValue))
                  vTgtPropValue = GetShortString(oCandidate, "Type", CLng(vTgtPropValue))

            Case "AssemblyStage"
                 If vCndtPropValue <> -1 Then vCndtPropValue = GetShortString(oCandidate, "Stage", CLng(vCndtPropValue)) Else vCndtPropValue = ""
                 If vTgtPropValue <> -1 Then vTgtPropValue = GetShortString(oCandidate, "Stage", CLng(vTgtPropValue)) Else vTgtPropValue = ""
                 
            Case "AssemblyChildrenCount", "SubAssemblyCount"
                If (IsEmpty(vCndtPropValue)) Or (IsNull(vCndtPropValue)) Then vCndtPropValue = 0
                If (IsEmpty(vTgtPropValue)) Or (IsNull(vTgtPropValue)) Then vTgtPropValue = 0
                
        End Select
 
         'Populate test command outputs
         If (IsNull(vCndtPropValue) And IsNull(vTgtPropValue)) Or (IsEmpty(vCndtPropValue) And IsEmpty(vTgtPropValue)) Then
             
             If (bAPISuccess = False) And (bIsGeomSame) Then
                 vCndtPropValue = "API Failed"
                 vTgtPropValue = "API Failed"
             Else
                 vCndtPropValue = ""
                 vTgtPropValue = ""
             End If
             
        End If
           
    If bCommon = True Then
        bStatus = "Same"
    Else
        bStatus = "Different"
    End If
     
    GetPropertyValues = CStr(vCndtPropValue) + "@" + CStr(vTgtPropValue) + "@" + bStatus
        
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

'Compares the bevels of the candidate and target objects
Public Function CompareBevels(ByRef oCandidate As Object, ByRef oTarget As Object, ByVal oMatchedPCs As IJElements, Optional bFlip As Boolean, Optional strBevelsInfo As String) As Boolean
                              
Const METHOD = "CompareBevels"
On Error GoTo ErrorHandler

    Dim oPlnIntHelper           As IJDPlnIntHelper
    Dim oPlnMatchingPair        As IJPlnPhysConnPair
    Dim strResultMsg            As String
    
    CompareBevels = False
    Set oPlnIntHelper = New CPlnIntHelper
    
    
    For Each oPlnMatchingPair In oMatchedPCs
    
        If oPlnMatchingPair.TypeOfMatch = PlnMatch_Full Then
            
            Dim strCandidateRefSide As String
            Dim strTargetRefSide As String
            
            'If flip flag is false reference sides are Base-Base
            strCandidateRefSide = "Base"
            strTargetRefSide = "Base"
            strResultMsg = ""
            
            'If flip flag is true reference sides are Base-Offset
            If bFlip = True Then
                strTargetRefSide = "Offset"
            End If
            
            CompareBevels = oPlnIntHelper.CompareBevels(oCandidate, oPlnMatchingPair.CandidatePhysConn, strCandidateRefSide, Nothing, _
                                                            oTarget, oPlnMatchingPair.TargetPhysConn, strTargetRefSide, Nothing, _
                                                            m_dDistanceTolerance, m_dAngleTolerance, True, strResultMsg)
                                                            
            strBevelsInfo = strBevelsInfo & strResultMsg & "NextMatchedPC"
        
           
            If CompareBevels = False Then
                Exit Function
            End If
            
        Else
            CompareBevels = False
            Exit Function
        End If
    Next
    
    Exit Function
    
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function


Public Function GeteObject(s As String) As ConnectedObjectTypes

    Select Case s
        Case "EdgeFeatures"
        GeteObject = EdgeFeatures
        
        Case "CornerFeatures"
        GeteObject = CornerFeatures
        
        Case "FreeEdgeTreatments"
        GeteObject = FreeEdgeTreatments
        
        Case "MarkingLines"
        GeteObject = MarkingLines
        
        Case "Shrinkages"
        GeteObject = Shrinkages
        
        Case "AssemblyShrinkages"
        GeteObject = AssemblyShrinkages
        
        Case "Features"
        GeteObject = AllFeatures
        
        Case "Margins"
        GeteObject = Margins
    End Select
End Function

'Method returns the standard parts from the given assembly hierarchy path
Public Function GetStandardReferencePartsCollection(strReferencePath As String) As IJDObjectCollection
Const METHOD = "GetStandardReferencePartsCollection"
On Error GoTo ErrorHandler

    Dim oPlnCommonPartHelper   As IJPlnCommonPartHelper
    Set oPlnCommonPartHelper = New CPlnCommonPartHelper
    
    Dim oStdRefPartsAssemblyCollection As IJDObjectCollection
    Dim oStdRefPartsCollection As IJDObjectCollection
    
    Set oStdRefPartsCollection = oPlnCommonPartHelper.GetStandardReferencePartsCollection(strReferencePath, oStdRefPartsAssemblyCollection)
     
    Set GetStandardReferencePartsCollection = oStdRefPartsCollection
    Set oPlnCommonPartHelper = Nothing
    
Exit Function
        
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

''******************************************************************************************
''  Function             : GetDecimalSeparator
''  Return Value         : Returns the decimal seperator corresponding to regional settings
'*******************************************************************************************
Public Function GetDecimalSeparator() As String
    Dim lLocale As Long
    Dim lLCType As Long
    Dim lpLCData As String * 20
    Const LOCALE_SDECIMAL = &HE
    ''Get the decimal seperator corresponding to regionalsettings and store it in a variable
    lLocale = GetUserDefaultLCID
    lLCType = LOCALE_SDECIMAL
    Call GetLocaleInfo(lLocale, lLCType, lpLCData, 4)
    GetDecimalSeparator = Left$(lpLCData, InStr(1, lpLCData, vbNullChar) - 1)
End Function
