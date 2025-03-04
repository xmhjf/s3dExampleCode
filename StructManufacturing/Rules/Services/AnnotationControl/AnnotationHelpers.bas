Attribute VB_Name = "AnnotationHelpers"
'*******************************************************************
'  Copyright (C) 2011 Intergraph.  All rights reserved.
'  Project:
'  Abstract:    Helpers.bas
'  History:
'       Seth Eden Hollingsead     October 24, 2011  Created
'******************************************************************

''**************************************************************************************
'' Routine      : UpdateSMS_Text_Args
'' Abstract     : Replaces the SMS_TEXT_ARGS INCLUDE_PROPERTY nodes with PROPERTY nodes from the
''              : SMS_ANNOTATION.xml file.
'' Inputs       : oDefaultXMLDom = SMS_ANNOTATION.xml
''              : oDefaultTypeElement = SMS_OUTPUT_ANNOTATION for the annotation attribute list that is being created.
''**************************************************************************************
Public Function UpdateSMS_TEXT_Args(oDefaultTypeElement As IXMLDOMElement, oTextNodes As IXMLDOMElement) _
As IXMLDOMElement
    Dim oDomDoc As New DOMDocument
    Dim oElem As IXMLDOMElement
    Dim oEditElem As IXMLDOMElement
    Dim oPropertyElem As IXMLDOMElement
    Dim oNodes As IXMLDOMNodeList
    Dim oPropertyNodes As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim oTempNode As IXMLDOMNode
    Dim oPropertyNode As IXMLDOMNode
    Dim oPropertyAttributes As IXMLDOMNamedNodeMap
    Dim oAttribute As IXMLDOMAttribute
    Dim oAttributeNode As IXMLDOMNode
    Dim oTempElem1 As IXMLDOMElement
    Dim oTempElem2 As IXMLDOMElement
    Dim strAttribName1 As String
    Dim strAttribName2 As String
    Dim i As Integer
    
    If oDefaultTypeElement Is Nothing Or oTextNodes Is Nothing Then GoTo CleanUp
    
    Set oNodes = oDefaultTypeElement.selectNodes(".//SMS_INPUT_ARGS/SMS_TEXT_ARGS/INCLUDE_PROPERTY")
    Set oEditElem = oDefaultTypeElement.selectSingleNode(".//SMS_INPUT_ARGS/SMS_TEXT_ARGS")
    Set oPropertyNodes = oTextNodes.selectNodes("//PROPERTY")
    If Not oNodes Is Nothing And Not oPropertyNodes Is Nothing And Not oEditElem Is Nothing Then
        If oNodes.Length > 0 And oPropertyNodes.Length > 0 Then
            For Each oNode In oNodes
                For Each oPropertyNode In oPropertyNodes
                    If Not oNode Is Nothing And Not oPropertyNode Is Nothing Then
                        Set oTempElem1 = oNode
                        Set oTempElem2 = oPropertyNode
                        strAttribName1 = oTempElem1.getAttribute("NAME")
                        strAttribName2 = oTempElem2.getAttribute("NAME")
                        If strAttribName1 = strAttribName2 Then
                            oEditElem.removeChild oNode
                            Set oElem = oDomDoc.createElement("PROPERTY")
                            oElem.setAttribute "NAME", strAttribName1
                            Set oPropertyAttributes = oPropertyNode.Attributes
                            If Not oPropertyAttributes Is Nothing Then
                                For i = 1 To oPropertyAttributes.Length
                                    Set oAttributeNode = oPropertyAttributes.Item(i)
                                    If Not oAttributeNode Is Nothing Then
                                        oElem.setAttribute oAttributeNode.nodeName, oAttributeNode.nodeValue
                                    End If
                                Next i
                                Set oTempNode = oEditElem.appendChild(oElem)
                            End If
                            Exit For 'Found the one we are looking for, start the next one.
                        End If
                    End If
                Next oPropertyNode
            Next oNode
        End If
    End If
    Set UpdateSMS_TEXT_Args = oDefaultTypeElement
CleanUp:
    Set oDomDoc = Nothing
    Set oElem = Nothing
    Set oEditElem = Nothing
    Set oPropertyElem = Nothing
    Set oNodes = Nothing
    Set oPropertyNodes = Nothing
    Set oNode = Nothing
    Set oTempNode = Nothing
    Set oPropertyNode = Nothing
    Set oPropertyAttributes = Nothing
    Set oAttribute = Nothing
    Set oAttributeNode = Nothing
    Set oTempElem1 = Nothing
    Set oTempElem2 = Nothing
End Function

''**************************************************************************************
'' Routine      : GetAttributesList
'' Abstract     : Parses through the SMS_ANNOTATION.xml and gets all the attributes
''              : according to how they should be loaded with default attributes according
''              : to what group it is in, and what attributes should be included with that group.
'' Inputs       : oDefaultXMLDom = SMS_ANNOTATION.xml
''              : oDefaultTypeElement = ATTRIBUTE_GROUP for the group that is being created.
''              : sPartType = PLATE or PROFILE to indicate what kind of part we are dealing with.
''**************************************************************************************
Public Function GetAttributesList(oDefaultXMLDOM As DOMDocument, _
oDefaultTypeElement As IXMLDOMElement, Optional sPartType As String = "", _
Optional bGetTextArgs As Boolean = False) As IXMLDOMNodeList
    Dim oAddNodeList As IXMLDOMNodeList
    Dim oTempNodeList As IXMLDOMNodeList
    Dim oTempElement As IXMLDOMElement
    Dim oTempAddElem As IXMLDOMElement
    Dim oModifyElem As IXMLDOMElement
    Dim oTempGroupElem As IXMLDOMElement
    Dim oXmlAttribute As IXMLDOMAttribute
    Dim sTempName As String
    Dim oMasterElement As IXMLDOMElement
    Dim oMasterNodeElem As IXMLDOMElement
    Dim oChildNodeList As IXMLDOMNodeList
    Dim oChildElement As IXMLDOMElement
    Dim sChildName As String
    
    If oDefaultXMLDOM Is Nothing Or oDefaultTypeElement Is Nothing Then GoTo CleanUp
    
    Set oMasterElement = oDefaultXMLDOM.createElement("MASTER")
    
    'Get all the attributes from all the groups
    Set oTempNodeList = oDefaultTypeElement.getElementsByTagName("INCLUDE_GROUP")
    For Each oTempElement In oTempNodeList
        sTempName = ""
        sTempName = oTempElement.getAttribute("NAME")
        If sTempName <> "" Then
            If InStr(1, sTempName, "MARKING") > 0 Then
                If UCase(sPartType) = "PLATE" Then
                    sTempName = "PLATE_" & sTempName
                ElseIf UCase(sPartType) = "PROFILE" Then
                    sTempName = "PROFILE_" & sTempName
                ElseIf UCase(sPartType) = "TEMPLATE" Then
                    sTempName = "TEMPLATE_" & sTempName
                Else 'Default to plate marking type!
                    sTempName = "PLATE_" & sTempName
                End If
            End If
            Set oTempGroupElem = Nothing
            Set oTempGroupElem = oDefaultXMLDOM.selectSingleNode("//ATTRIBUTE_GROUP[@NAME=""" & sTempName & """]")
            If Not oTempGroupElem Is Nothing Then
                Set oAddNodeList = Nothing
                Set oAddNodeList = GetAttributesList(oDefaultXMLDOM, oTempGroupElem, , bGetTextArgs)
                For Each oTempAddElem In oAddNodeList
                    oMasterElement.appendChild oTempAddElem.cloneNode(True)
                Next oTempAddElem
            End If
        End If
    Next oTempElement
    
    'Get all the individual listed property attributes
    Set oTempNodeList = Nothing
    Set oTempNodeList = oDefaultTypeElement.getElementsByTagName("INCLUDE_PROPERTY")
    For Each oTempElement In oTempNodeList
        sTempName = ""
        sTempName = oTempElement.getAttribute("NAME")
        If sTempName <> "" Then
            Set oTempAddElem = Nothing
            Set oTempAddElem = oDefaultXMLDOM.selectSingleNode("//PROPERTY[@NAME=""" & sTempName & """]")
            If Not oTempAddElem Is Nothing Then
                'Need to first check and see if oMasterElement already has an oTempAddElem node. It may have been
                'included from the above group include. If it is there, then we just need to update it's values,
                'with the values that we find in this instance. This instance over-writes the other one.
                If oMasterElement.hasChildNodes = True Then
                    Set oChildNodeList = oMasterElement.childNodes
                    If oChildNodeList.Length >= 1 Then
                        For Each oChildElement In oChildNodeList
                            If Not oChildElement Is Nothing Then
                                sChildName = oChildElement.getAttribute("NAME")
                                If sChildName = sTempName Then
                                    Set oMasterNodeElem = oChildElement
                                    GoTo UpdatePropertyDefinition
                                End If
                            End If
                        Next oChildElement
                    End If
                End If
                Set oMasterNodeElem = oMasterElement.appendChild(oTempAddElem.cloneNode(True))
                'We are not going to have a MODIFY_PROPERTY any more, and if you want to over-write the default
                'values for a given attribute definition then you should just set them directly in the
                'INCLUDE_PROPERTY node. Here we will over-write the defaults with any definitions
                'that the user wishes to customize.
UpdatePropertyDefinition:
                If Not oMasterNodeElem Is Nothing Then
                    For Each oXmlAttribute In oTempElement.Attributes
                        If Not oXmlAttribute Is Nothing Then
                            If oXmlAttribute.Name <> "NAME" Then
                                oMasterNodeElem.setAttribute oXmlAttribute.Name, oXmlAttribute.Value
                            End If
                        End If
                    Next oXmlAttribute
                End If
            End If
        End If
    Next oTempElement
    
    'We are not using SMS_TEXT_ARG nodes any more, not explicitely, they are converted later after this function.
    'everything is pulled using the Attribute Definitions.
    'So we use the keyword INCLUDE_PROPERTY, which is done above.
    'However, even if we aren't using SMS_TEXT_ARG nodes any more they still can possibly show up when coming into here
    'under certain circumstances, so we still need this here!
    'Get all the individual listed PROPERTY attributes
    If bGetTextArgs = True Then
        Set oTempNodeList = Nothing
        Set oTempNodeList = oDefaultTypeElement.getElementsByTagName("PROPERTY")
        For Each oTempElement In oTempNodeList
            sTempName = ""
            sTempName = oTempElement.getAttribute("NAME")
            If sTempName <> "" Then
                Set oTempAddElem = Nothing
                Set oTempAddElem = oDefaultXMLDOM.selectSingleNode("//PROPERTY[@NAME=""" & sTempName & """]")
                If Not oTempAddElem Is Nothing Then
                    oMasterElement.appendChild oTempAddElem.cloneNode(True)
                End If
            End If
        Next oTempElement
    End If
    
    'We are not going to use the MODIFY_PROPERTY keyword any more, it's all done directly in the INCLUDE_PROPERTY.
    'See above for additional details.
    'We still need this because many of the properties are modified after being included from a group.
    'so there is no corresponding INCLUDE_PROPERTY at the same level that it is modified.
'    'Modify the nodes based on all the modify attribute nodes
'    Set oTempNodeList = Nothing
'    Set oTempNodeList = oDefaultTypeElement.getElementsByTagName(MODIFY_PROPERTY)
'    For Each oTempElement In oTempNodeList
'        For Each oModifyElem In oMasterElement.childNodes
'            If oModifyElem.GetAttribute(NAME_ATTRIB) = oTempElement.GetAttribute(NAME_ATTRIB) Then
'                For Each oXmlAttribute In oTempElement.Attributes
'                    If oXmlAttribute.Name <> NAME_ATTRIB Then
'                        oModifyElem.SetAttribute oXmlAttribute.Name, oXmlAttribute.VALUE
'                    End If
'                Next oXmlAttribute
'                Exit For
'            End If
'        Next oModifyElem
'    Next oTempElement
    
    Set GetAttributesList = oMasterElement.childNodes
CleanUp:
    Set oAddNodeList = Nothing
    Set oTempNodeList = Nothing
    Set oTempElement = Nothing
    Set oTempAddElem = Nothing
    Set oModifyElem = Nothing
    Set oTempGroupElem = Nothing
    Set oXmlAttribute = Nothing
    Set oMasterElement = Nothing
End Function

''**************************************************************************************
'' Routine      : UpdateAnnotationPropertyNodes
'' Abstract     : Update all the Property Nodes of the Annotation object correctly,
''              : no matter what type of annotation it is.
'' Inputs       : oSettingsXML = SMS_ANNOTATION.xml
''              : oDefaultTypeElement = ATTRIBUTE_GROUP for the group that is being created.
''              : sPartType = PLATE or PROFILE to indicate what kind of part we are dealing with.
'' Returns      : UpdateAnnotationPropertyNodes as an DOM XML Element that contains the updated annotation XML.
''**************************************************************************************
Public Function UpdateAnnotationPropertyNodes(oSettingsXML As DOMDocument, oEntityNameElem As IXMLDOMElement, _
sPartType As String) As IXMLDOMElement
    Dim oNode As IXMLDOMNode
    Dim oTempDoc As New DOMDocument
    Dim oTempElem As IXMLDOMElement
    Dim oTempNodeList As IXMLDOMNodeList
    
    If Not oSettingsXML Is Nothing And Not oEntityNameElem Is Nothing And sPartType <> "" Then
        Set oTempElem = oTempDoc.createElement("MASTER")
        Set oTempNodeList = GetAttributesList(oSettingsXML, oEntityNameElem, sPartType, True)
        If Not oTempNodeList Is Nothing Then
            For Each oNode In oTempNodeList
                oTempElem.appendChild oNode.cloneNode(True)
            Next oNode
        End If
'NOTE: We need to update oEntityNameElem with the proper structure as well!
''<SMS_TEXT_ARGS>
''    <INCLUDE_PROPERTY NAME="CHAMFER_ANGLE_M"/>
''    <INCLUDE_PROPERTY NAME="CHAMFER_DEPTH_M"/>
''</SMS_TEXT_ARGS>
'' Should be:
''<SMS_TEXT_ARGS>
''    <PROPERTY NAME="CHAMFER_ANGLE_M" DISPLAY_NAME="CHAMFER_ANGLE_M" VALUE="0" READONLY="0" UNIT="78" UNIT_TYPE="ANGLE" LBOUND="0.0" UBOUND="50.0" DISPLAY_ORDER="50" />
''    <PROPERTY NAME="CHAMFER_DEPTH_M" DISPLAY_NAME="CHAMFER_DEPTH_M" VALUE="0" READONLY="0" UNIT="61" UNIT_TYPE="DISTANCE" LBOUND="0.0" DISPLAY_ORDER="51" />
''</SMS_TEXT_ARGS>
        Set oEntityNameElem = UpdateSMS_TEXT_Args(oEntityNameElem, oTempElem)
    End If
    Set UpdateAnnotationPropertyNodes = oEntityNameElem
CleanUp:
    Set oNode = Nothing
    Set oTempDoc = Nothing
    Set oTempElem = Nothing
End Function

Public Function CheckIfLongestCommonSeam(oMfgPart As Object, oGeom2d As IJMfgGeom2d) As Boolean
Const METHOD = "GetProductionRoutingStageCode"
On Error GoTo ErrorHandler

    CheckIfLongestCommonSeam = True
    
    'Get Input Geometry's name
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = oGeom2d
    Dim sGeometryName As String
    sGeometryName = oNamedItem.Name
    
    Dim oInputCurve As IJCurve
    Set oInputCurve = oGeom2d.GetGeometry

    Dim pResMgr As IUnknown
    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeom2dColFactory As GSCADMfgGeometry.MfgGeomCol2dFactory
    Set oGeom2dColFactory = New GSCADMfgGeometry.MfgGeomCol2dFactory
    
    ' Get the goemetries before unfold
    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = oGeom2dColFactory.Create(pResMgr)

    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
    Set oMfgPlateWrapper.object = oMfgPart
    Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries
        
    Dim i As Long
    Dim oCurrGeom2d As IJMfgGeom2d
    Dim oCurrCurve As IJCurve
    
    If oGeomCol2d Is Nothing Then GoTo CleanUp
    
    For i = 1 To oGeomCol2d.Getcount
    
        Set oCurrGeom2d = oGeomCol2d.GetGeometry(i)
        'Get only common seam marks and check only NON-support only marks
        If oCurrGeom2d.GetGeometryType = STRMFG_COMMON_SEAM_MARK And oCurrGeom2d.IsSupportOnly = True Then
            Set oNamedItem = oCurrGeom2d
            If oNamedItem.Name = sGeometryName Then
                
                 Set oCurrCurve = oCurrGeom2d.GetGeometry
                 If oCurrCurve.Length > oInputCurve.Length Then
                    'This means there is another bigger common seam mark for same part
                    CheckIfLongestCommonSeam = False
                    Exit For
                 End If
            End If
        End If
    Next i
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


'**************************************************************************************
' Method       : GetProductionRoutingStageCode
' Abstract     : This Function gets the ProductionRouting stage code
'**************************************************************************************
Public Function GetProductionRoutingStageCode(oPart As Object) As String
Const METHOD = "GetProductionRoutingStageCode"
On Error GoTo ErrorHandler
    
    If oPart Is Nothing Then Exit Function
    
    Dim oPlnProdRouting As PlanningObjects.PlnProdRouting
    Set oPlnProdRouting = New PlanningObjects.PlnProdRouting
    Set oPlnProdRouting.object = oPart
    
    GetProductionRoutingStageCode = oPlnProdRouting.GetProductionRoutingCode
    
    Set oPlnProdRouting = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function isNumericUS(vInput As Variant) As Boolean
    Const METHOD = "isNumericUS"
    On Error GoTo ErrorHandler
    Dim strInput As String
    strInput = vInput
    strInput = Trim(strInput)
    strInput = Replace(strInput, "0", "")
    strInput = Replace(strInput, ".", "", , 1)
    strInput = Replace(strInput, ",", "")
    If Len(strInput) = 0 Then
        isNumericUS = True
        Exit Function
    End If
    Dim dNumber As Double
    dNumber = -100000000000#
    dNumber = Val(vInput)
    If Not dNumber = 0 Then
        isNumericUS = True
    Else
        isNumericUS = False
    End If
    
Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

