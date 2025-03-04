Attribute VB_Name = "CommonFuncs"
Option Explicit
Private Const E_ACCESSDENIED As Long = -2147467259

Public Function GetRuleValue(strRuleName As String, strColumn As String) As Variant
    Const METHOD = "GetRuleValue"
    On Error GoTo ErrorHandler

    Dim oOutputHelper As IJMfgOutputHelper
    Set oOutputHelper = New MfgCatalogQueryHelper
    If Not oOutputHelper Is Nothing Then
        On Error Resume Next
        Err.Clear
        GetRuleValue = oOutputHelper.GetOutputRuleValue(strRuleName, strColumn)

        If Err.Number = E_ACCESSDENIED Then
            GetRuleValue = ""
            Set oOutputHelper = Nothing
            Exit Function
        End If
    End If
CleanUp:
    Set oOutputHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Public Function GetXMLAttribute(oNode As IXMLDOMNode, ByVal strAttr As String, ByRef strVal As String) As Boolean
Const METHOD = "GetXMLAttribute"
On Error GoTo ErrHandler
    GetXMLAttribute = False

    If Not oNode.Attributes Is Nothing Then
        If Not oNode.Attributes.getNamedItem(strAttr) Is Nothing Then
            GetXMLAttribute = True
            strVal = oNode.Attributes.getNamedItem(strAttr).nodeValue
        End If
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number
End Function

Public Function GetResourceManager() As IJDPOM
    Const METHOD = "GetActiveConnection"
    On Error GoTo ErrorHandler
    
    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    Dim oPOM As IUnknown
    oCmnAppGenericUtil.GetResourceManager oPOM
    
    Set GetResourceManager = oPOM

    Set oCmnAppGenericUtil = Nothing
    Exit Function
    
ErrorHandler:
    Err.Clear

End Function


Public Function GetOID(oObject As Object) As String
    On Error GoTo ErrorHandler
    Const METHOD = "GetOID"
    
    Dim oPOM As IJDPOM
    Set oPOM = GetResourceManager()

    'Retrive the OID from the Manufacturing Object
    Dim oMoniker As IMoniker
    Set oMoniker = oPOM.GetObjectMoniker(oObject)
    GetOID = oPOM.DbIdentifierFromMoniker(oMoniker)
    
    Set oMoniker = Nothing
    
    Exit Function
    
ErrorHandler:
    Resume Next
End Function

Public Function CleanGUID(sGUID As String) As String
    On Error GoTo ErrorHandler
    Const METHOD = "CleanGUID"
    
    Dim sTemp As String
    
    If InStr(1, sGUID, "{") >= 1 And InStr(1, sGUID, "}") >= 1 Then
        'The string has brackets in it.
        sTemp = Mid(sGUID, InStr(1, sGUID, "{") + 1, Len(sGUID))
        CleanGUID = Mid(sTemp, 1, InStr(1, sTemp, "}") - 1)
    Else
        CleanGUID = sGUID
    End If
    Exit Function
ErrorHandler:
    Resume Next
End Function

Public Function GenerateOutputDocument(ByVal strID As String, ByVal strOutputDocumentName As String) As DOMDocument
Const METHOD = "GenerateOutputDocument"
On Error GoTo ErrorHandler

    'Generate Output ReportDocument
    Dim oOutputReportDoc As New DOMDocument
    Dim oOutputReportNode As IXMLDOMNode
    Set oOutputReportNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputReport", "")
    oOutputReportDoc.appendChild oOutputReportNode

    Dim oOutputDocumentNode As IXMLDOMNode
    Set oOutputDocumentNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputDocument", "")

    Dim oOutputDocumentElem As IXMLDOMElement
    Set oOutputDocumentElem = oOutputDocumentNode

    oOutputDocumentElem.setAttribute "ID", strID

    Dim lastPos As Long
    lastPos = InStrRev(strOutputDocumentName, "\")
    Dim strFileName As String
    strFileName = Trim(Right$(strOutputDocumentName, Len(strOutputDocumentName) - lastPos))
    oOutputDocumentElem.setAttribute "NAME", strFileName

    oOutputReportNode.appendChild oOutputDocumentNode
    
    Set GenerateOutputDocument = oOutputReportDoc
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function


Public Function IsProfileABuiltUp(ByVal oMfgProfilePart As IJMfgProfilePart) As Boolean
Const METHOD = "IsProfileABuiltUp"
On Error GoTo ErrorHandler

    IsProfileABuiltUp = False

    On Error Resume Next
    Dim oMfgProfileOutput As IJMfgProfileOutput
    Set oMfgProfileOutput = oMfgProfilePart
    
    On Error GoTo ErrorHandler
    If Not oMfgProfileOutput Is Nothing Then
    
        Dim lCells As Long: lCells = 0
        Dim strCellNames() As Variant
        oMfgProfileOutput.GetOutputKeys lCells, strCellNames
 
        If lCells = 0 Then
        
            Dim lManufacturedAsPlate As Long
            lManufacturedAsPlate = oMfgProfilePart.ManufactureAsPlate
            
            If lManufacturedAsPlate = 1 Then
            
                Dim oProfilePart As Object
                Set oProfilePart = oMfgProfilePart.GetDetailedPart
                
                Dim oPartSupport As IJPartSupport
                Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
                
                Set oPartSupport.Part = oProfilePart

                Dim oProfilePartSupport As IJProfilePartSupport
                Set oProfilePartSupport = oPartSupport
                
                
                Dim enumProfileSectionType As ProfileSectionType
                enumProfileSectionType = oProfilePartSupport.SectionType
                
                If Not enumProfileSectionType = Flat_Bar Then
                    IsProfileABuiltUp = True
                End If
                
                Set oProfilePartSupport = Nothing
                Set oProfilePart = Nothing
            ElseIf Not lManufacturedAsPlate = 0 Then
                IsProfileABuiltUp = True
            End If
            
        ElseIf lCells >= 2 Then
            IsProfileABuiltUp = True
        End If
       
        Set oMfgProfileOutput = Nothing
    End If
    

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function




Public Sub GetDXFFileNameAndPartID(ByVal pObject As Object, ByVal oPlateOrProfileNode As IXMLDOMNode, ByVal bstrBaseOutputName As String, ByVal strExtension As String, ByRef strFullFileName As String, ByRef strPartID As String)
Const METHOD = "GetDXFFileNameAndPartID"
On Error GoTo ErrorHandler

   
    Dim strPartName As String
    Dim strModelPartGuid As String
    Dim strPartGuid As String
    Dim strCellName As String
  
    If Not oPlateOrProfileNode Is Nothing Then

        strPartName = oPlateOrProfileNode.selectSingleNode(".//SMS_PART_INFO/@PART_NAME").nodeValue
        strModelPartGuid = oPlateOrProfileNode.selectSingleNode(".//SMS_PART_INFO/@MODEL_PART_GUID").nodeValue
        strPartGuid = oPlateOrProfileNode.selectSingleNode(".//SMS_PART_INFO/@PART_GUID").nodeValue

    End If

    
    strPartID = "{" + strPartGuid + "}"
    
    If TypeOf pObject Is IJMfgPlatePart Then
        strFullFileName = bstrBaseOutputName + "_" + "{" + strModelPartGuid + "}" + strExtension
    ElseIf TypeOf pObject Is IJMfgProfilePart Then
        
        strCellName = oPlateOrProfileNode.selectSingleNode(".//SMS_PART_INFO/@CELL_NAME").nodeValue
        If IsProfileABuiltUp(pObject) = True Then
           
            strPartID = strPartID + "-" + strCellName
            strFullFileName = bstrBaseOutputName + "-" + strCellName + "_" + "{" + strModelPartGuid + "}" + strExtension
        Else
            strFullFileName = bstrBaseOutputName + "_" + "{" + strModelPartGuid + "}" + strExtension
        End If
     
    ElseIf TypeOf pObject Is IJDMfgTemplateSet Then
        strPartID = GetOID(pObject) + "-" + strPartName
        strFullFileName = bstrBaseOutputName + "-" + strPartName + "_" + GetOID(pObject) + strExtension
     
    End If
     

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Sub


Public Function SanitizeFileName(ByVal strInput As String) As String
Const sMETHOD As String = "SanitizeFileName"
On Error GoTo ErrorHandler

    SanitizeFileName = "SanitizeFileName"
    '<>:\"/\\|?*
    
    Dim tempString As String
    tempString = strInput
    
    If InStr(tempString, "<") > 0 Then
        tempString = Replace(tempString, "<", "_")
    End If
    
    If InStr(tempString, ">") > 0 Then
        tempString = Replace(tempString, ">", "_")
    End If
    
    If InStr(tempString, ":") > 0 Then
        tempString = Replace(tempString, ":", "_")
    End If
    
    If InStr(tempString, "\") > 0 Then
        tempString = Replace(tempString, "\", "_")
    End If
    
    If InStr(tempString, "/") > 0 Then
        tempString = Replace(tempString, "/", "_")
    End If
    
    If InStr(tempString, "//") > 0 Then
        tempString = Replace(tempString, "//", "_")
    End If
    
    If InStr(tempString, "|") > 0 Then
        tempString = Replace(tempString, "|", "_")
    End If
    
    If InStr(tempString, "?") > 0 Then
        tempString = Replace(tempString, "?", "_")
    End If
    
    If InStr(tempString, "*") > 0 Then
        tempString = Replace(tempString, "*", "_")
    End If
    
    SanitizeFileName = tempString
    
    
    Exit Function
ErrorHandler:

End Function


Public Function ConvertNumberToENLocale(ByVal value As Double) As String
Const METHOD = "ConvertNumberToENLocale"
On Error GoTo ErrorHandler

    Dim returnValue As String
    returnValue = CStr(Format(value, "0.00"))

    If InStr(returnValue, ",") > 0 Then
        ConvertNumberToENLocale = Replace(returnValue, ",", ".")
    Else
        ConvertNumberToENLocale = returnValue
    End If

    Exit Function
ErrorHandler:
    Resume Next

End Function

Public Function GetRelatedObjects(ByVal oObject As Object, ByVal strInterfaceID As String, ByVal strCollectionName As String) As IJElements
Const METHOD = "GetRelatedObjects"
On Error GoTo ErrorHandler

    Dim oAssocRel As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    
    Dim oResultColl As IJElements
    Set oResultColl = New JObjectCollection
    
    On Error Resume Next
    Set oAssocRel = oObject

    If Not oAssocRel Is Nothing Then
        Set oTargetObjCol = oAssocRel.CollectionRelations(strInterfaceID, strCollectionName)
        
        If oTargetObjCol.Count > 0 Then
            Dim i As Long
            
            For i = 1 To oTargetObjCol.Count
                oResultColl.Add oTargetObjCol.Item(i)
            Next
        End If
    End If
    
    Set GetRelatedObjects = oResultColl

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Public Function IsRootPC(ByVal oPC As IJStructPhysicalConnection) As Boolean
Const METHOD = "IsRootPC"
On Error GoTo ErrorHandler

    IsRootPC = False
    On Error Resume Next
    Dim oDesignParent As IJDesignParent
    Set oDesignParent = oPC
    
    On Error GoTo ErrorHandler
    If Not oDesignParent Is Nothing Then
        Dim oChildPCs As IJDObjectCollection
        oDesignParent.GetChildren oChildPCs
        
        If Not oChildPCs Is Nothing Then
            Dim oChildObject As Object
            For Each oChildObject In oChildPCs
                If TypeOf oChildObject Is IJStructPhysicalConnection Then
                    IsRootPC = True
                    Set oChildObject = Nothing
                    Exit For
                End If
                Set oChildObject = Nothing
            Next
        End If
        Set oChildPCs = Nothing
    End If
    
    Set oDesignParent = Nothing

    Exit Function
ErrorHandler:
    Set oDesignParent = Nothing
    Err.Raise Err.Number, , Err.Description
End Function
Public Sub GenerateGeometryNode(ByVal oOutputXML As IXMLDOMDocument, ByRef oParentNode As IXMLDOMNode, ByRef oPhyConCurve As IJCurve)
    Const METHOD = "GenerateGeometryNode"
    On Error GoTo ErrorHandler

    ' Initial version, we only create the end-points as reference
   
    If Not oPhyConCurve Is Nothing Then
        
        On Error Resume Next
        Dim oPhyConComplexString As IJComplexString
        Set oPhyConComplexString = oPhyConCurve
        
        On Error GoTo ErrorHandler
        If oPhyConComplexString Is Nothing Then
            Set oPhyConComplexString = New ComplexString3d
            oPhyConComplexString.AddCurve oPhyConCurve, True
        End If
                
            
        Dim oSMSEdgeElem As IXMLDOMElement
        Dim oSMSEdgeNode As IXMLDOMNode
        Set oSMSEdgeNode = oOutputXML.createNode(NODE_ELEMENT, "SMS_EDGE", "")
        Set oSMSEdgeElem = oSMSEdgeNode

        
        Dim oMfgMathGeom As New MfgMathGeom
        Dim oLineArcComplexString As IJComplexString
        

        oMfgMathGeom.Convert3DCurvesToLineArc Nothing, oPhyConComplexString, 0.0001, oLineArcComplexString
        If Not oLineArcComplexString Is Nothing Then
        
            Dim oCurveElem As IJElements
            oLineArcComplexString.GetCurves oCurveElem
            
            Dim strSEGType As String
            Dim lCount As Long
            
            For lCount = 1 To oCurveElem.Count
            
                Dim oCVGCurveElem As IXMLDOMElement
                Dim oCVGCurveNode As IXMLDOMNode
                Set oCVGCurveNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_CURVE", "")
                Set oCVGCurveElem = oCVGCurveNode
                
                Dim oCVGVertexElem As IXMLDOMElement
                Dim oCVGVertexNode As IXMLDOMNode

                Dim oLine As IJLine
                Dim oCurve As IJCurve
                
                On Error Resume Next
                Set oLine = oCurveElem.Item(lCount)
                Set oCurve = oCurveElem.Item(lCount)
                
                Dim dStartX As Double, dStartY As Double, dStartZ As Double
                Dim dMiddleX As Double, dMiddleY As Double, dMiddleZ As Double
                Dim dEndX As Double, dEndY As Double, dEndZ As Double
                Dim dStartParam As Double
                Dim dEndParam As Double
                
                On Error GoTo ErrorHandler
                
                If Not oLine Is Nothing Then
                    strSEGType = "line"
                    oCVGCurveElem.setAttribute "CURVE_TYPE", "line"
                    If Not oCurve Is Nothing Then
                        oCVGCurveElem.setAttribute "LENGTH", ConvertNumberToENLocale(oCurve.length * 1000)
                        oCurve.ParamRange dStartParam, dEndParam
                        oCurve.Position dStartParam, dStartX, dStartY, dStartZ
                        oCurve.Position dEndParam, dEndX, dEndY, dEndZ
                    End If
                    
                    Set oCVGVertexNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_VERTEX", "")
                    Set oCVGVertexElem = oCVGVertexNode
                    
                    'Convert it to MM
                    dStartX = dStartX * 1000
                    dStartY = dStartY * 1000
                    dStartZ = dStartZ * 1000
            
                    oCVGVertexElem.setAttribute "POINT_CODE", "s_point"
                    oCVGVertexElem.setAttribute "SEG_TYPE", strSEGType
                    oCVGVertexElem.setAttribute "SX", ConvertNumberToENLocale(dStartX)
                    oCVGVertexElem.setAttribute "SY", ConvertNumberToENLocale(dStartY)
                    oCVGVertexElem.setAttribute "SZ", ConvertNumberToENLocale(dStartZ)
            
                    oCVGCurveNode.appendChild oCVGVertexNode
            
                    Set oCVGVertexNode = Nothing
                    Set oCVGVertexElem = Nothing
            
                    Set oCVGVertexNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_VERTEX", "")
                    Set oCVGVertexElem = oCVGVertexNode
            
                    'Convert it to MM
                    dEndX = dEndX * 1000
                    dEndY = dEndY * 1000
                    dEndZ = dEndZ * 1000
            
                    oCVGVertexElem.setAttribute "POINT_CODE", "e_point"
                    oCVGVertexElem.setAttribute "SEG_TYPE", "dummy"
                    oCVGVertexElem.setAttribute "SX", ConvertNumberToENLocale(dEndX)
                    oCVGVertexElem.setAttribute "SY", ConvertNumberToENLocale(dEndY)
                    oCVGVertexElem.setAttribute "SZ", ConvertNumberToENLocale(dEndZ)
                    
                    oCVGCurveNode.appendChild oCVGVertexNode
                
                Else
                    
                    strSEGType = "circle"
                    oCVGCurveElem.setAttribute "CURVE_TYPE", "arc"
                    If Not oCurve Is Nothing Then
                        oCVGCurveElem.setAttribute "LENGTH", ConvertNumberToENLocale(oCurve.length * 1000)
                        
                        oCurve.ParamRange dStartParam, dEndParam
                        oCurve.Position dStartParam, dStartX, dStartY, dStartZ
                        oCurve.Position (dStartParam + dEndParam) / 2#, dMiddleX, dMiddleY, dMiddleZ
                        oCurve.Position dEndParam, dEndX, dEndY, dEndZ
                    End If

                    Set oCVGVertexNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_VERTEX", "")
                    Set oCVGVertexElem = oCVGVertexNode
                    
                    'Convert it to MM
                    dStartX = dStartX * 1000
                    dStartY = dStartY * 1000
                    dStartZ = dStartZ * 1000
            
                    oCVGVertexElem.setAttribute "POINT_CODE", "s_point"
                    oCVGVertexElem.setAttribute "SEG_TYPE", strSEGType
                    oCVGVertexElem.setAttribute "SX", ConvertNumberToENLocale(dStartX)
                    oCVGVertexElem.setAttribute "SY", ConvertNumberToENLocale(dStartY)
                    oCVGVertexElem.setAttribute "SZ", ConvertNumberToENLocale(dStartZ)

                    oCVGCurveNode.appendChild oCVGVertexNode
            
                    Set oCVGVertexNode = Nothing
                    Set oCVGVertexElem = Nothing
                    
                    Set oCVGVertexNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_VERTEX", "")
                    Set oCVGVertexElem = oCVGVertexNode
                    
                    'Convert it to MM
                    dMiddleX = dMiddleX * 1000
                    dMiddleY = dMiddleY * 1000
                    dMiddleZ = dMiddleZ * 1000
            
                    oCVGVertexElem.setAttribute "POINT_CODE", "m_point"
                    oCVGVertexElem.setAttribute "SEG_TYPE", strSEGType
                    oCVGVertexElem.setAttribute "SX", ConvertNumberToENLocale(dMiddleX)
                    oCVGVertexElem.setAttribute "SY", ConvertNumberToENLocale(dMiddleY)
                    oCVGVertexElem.setAttribute "SZ", ConvertNumberToENLocale(dMiddleZ)

                    oCVGCurveNode.appendChild oCVGVertexNode
            
                    Set oCVGVertexNode = Nothing
                    Set oCVGVertexElem = Nothing
            
                    Set oCVGVertexNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_VERTEX", "")
                    Set oCVGVertexElem = oCVGVertexNode
            
                    'Convert it to MM
                    dEndX = dEndX * 1000
                    dEndY = dEndY * 1000
                    dEndZ = dEndZ * 1000
            
                    oCVGVertexElem.setAttribute "POINT_CODE", "e_point"
                    oCVGVertexElem.setAttribute "SEG_TYPE", "dummy"
                    oCVGVertexElem.setAttribute "SX", ConvertNumberToENLocale(dEndX)
                    oCVGVertexElem.setAttribute "SY", ConvertNumberToENLocale(dEndY)
                    oCVGVertexElem.setAttribute "SZ", ConvertNumberToENLocale(dEndZ)
                    
                    oCVGCurveNode.appendChild oCVGVertexNode

                    
                End If
                
                oSMSEdgeNode.appendChild oCVGCurveNode
                
                Set oLine = Nothing
                Set oCurve = Nothing
                Set oCVGVertexNode = Nothing
                Set oCVGVertexElem = Nothing
                Set oCVGCurveNode = Nothing
                Set oCVGCurveElem = Nothing
                
            Next

            oParentNode.appendChild oSMSEdgeNode
  
        End If

    End If

    Exit Sub
ErrorHandler:
    Resume Next

End Sub
