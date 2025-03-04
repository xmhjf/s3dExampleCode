VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IngrTemplateSet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CIngrTemplateSetXML2.cls
''
''  Description : Populates the XML DOM Document with TemplateSet related data for Sample 2
''
''
''  Author      : Intergraph
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************

Implements IJDCustomReport

Private Const MODULE = "CustomReports.CIngrTemplateSetXML2: "

Private m_objXMLDoc             As DOMDocument
Private m_oCurrentRootNode      As IXMLDOMNode
Private m_oFileRootNode         As IXMLDOMNode

' Variable that stores the Filtering Info.
Private m_objFilter             As IXMLDOMElement

'Private m_oDwgProgress         As IJDDwgProgress

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    InitializeFilter m_objFilter
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub Class_Terminate()
    Const METHOD = "Class_Terminate"
    On Error GoTo ErrorHandler

    Set m_objXMLDoc = Nothing
    Set m_oCurrentRootNode = Nothing
    Set m_oFileRootNode = Nothing
    Set m_objFilter = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub


Private Sub InitializeFilter(filterElement As IXMLDOMElement)
    On Error Resume Next

    ' Yes - Hard Code the filter Info Here
    
    Dim strFilter As String
    
    strFilter = "<report reportName='MfgTemplate' >" & _
                        "<elementList elementListName='Templates' >" & _
                            "<element elementName='Template' >" & _
                                "<propertyList propertyListName='Contents' report='YES' >" & _
                                    "<property propertyName='Type' report='YES' />" & _
                                "</propertyList>" & _
                            "</element>" & _
                        "</elementList>" & _
                    "</report>"

    Dim oFilterDoc As New DOMDocument
    If oFilterDoc.loadXML(strFilter) Then
        Set filterElement = oFilterDoc.documentElement
    Else
        Set filterElement = Nothing
    End If
End Sub

Private Sub CreateHeader(sReportName As String, sReportUID As String)
    Const METHOD As String = "CreateHeader"
    On Error GoTo ErrorHandler
    
    'Set m_oFileRootNode report to be the root for the XMLsheet
    Set m_oFileRootNode = m_objXMLDoc.createNode(NODE_ELEMENT, "report", "")
    Set m_oFileRootNode = m_objXMLDoc.appendChild(m_oFileRootNode)
    
    'Add attributes reportName and reportUID to the rootnode
    AddAttribute m_oFileRootNode, "reportName", sReportName
    AddAttribute m_oFileRootNode, "reportUID", sReportUID
    
    'Create elementlist and set to m_oCurrentRootNode
    Dim oElementList As IXMLDOMNode
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "MfgTemplate"
    Set m_oCurrentRootNode = oElementList
    Set oElementList = Nothing
    
CleanUp:
    Set oElementList = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
    
End Sub

Private Sub SortTemplateElements(oInElements As IJElements)
    On Error GoTo ErrorHandler
    Const METHOD As String = "SortTemplateElements"
    
    Dim oIJStifSyst As IJStiffenerSystem 'identifies a profilesystem
    Dim oStiffnerPart As IJStiffenerPart 'identifies a profilesystem
    Dim oSystem As IJSystem
    Dim oIJStructNavUtils As GSCADStructNavigationUtilitiesLib.StructNavigationUtils
    Dim oEnumUnk As IEnumUnknown
    Dim oATPStructHelper As CONVERTUTILITIESLib.IJIEnumIterator
    Dim oCreateVBColl As CONVERTUTILITIESLib.CCollectionConversions ' IJCreateVBCollection.
    Dim oCollObj As Object
    Dim oColItem As Object
    Dim oTemplateSet As IJDMfgTemplateSet
    Dim oObj As Object
    Dim bStiffnerSystem As Boolean
    Dim bStiffnerPart As Boolean
    Dim lCount As Long
    
    SET_PROGRESS_INFO oInElements.Count, 2, 1
    
    For lCount = 1 To oInElements.Count
    
        CHECK_POINT_ELEMENT oInElements(lCount)

        If TypeOf oInElements.Item(lCount) Is IJDMfgTemplateSet Then
            'Step through collection and get the objects which implemnts IJProfilePart
            'The following objects does: ProfilePart, StiffenerPart, BeamPart
            For Each oColItem In oInElements

                'QI for IJProfilePart
                Set oTemplateSet = oColItem

                If Not oTemplateSet Is Nothing Then

                    'if object is IJProfilePart, call method to extract to xml
                    'if extractmethod fails, don't stop, but continue to next object
                    On Error Resume Next 'prevent stop
                    ExtractTemplateSetToXML oTemplateSet, m_oCurrentRootNode
                    On Error GoTo ErrorHandler
                    Set oTemplateSet = Nothing
                End If
                Set oColItem = Nothing
            Next
        End If
        
        Set oObj = Nothing
    Next

CleanUpp:
    Set oObj = Nothing
    Set oIJStifSyst = Nothing
    Set oSystem = Nothing
    Set oIJStructNavUtils = Nothing
    Set oEnumUnk = Nothing
    Set oATPStructHelper = Nothing
    Set oCreateVBColl = Nothing
    Set oCollObj = Nothing
    Set oColItem = Nothing
    Set oTemplateSet = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUpp
    
End Sub

Public Function ExtractTemplateSetToXML(oObj As IUnknown, Parent As IXMLDOMNode) As Boolean
    Const METHOD As String = "ExtractTemplateSetToXML"
    On Error GoTo ErrorHandler

    Dim oElement As IXMLDOMNode
    Dim oPropList As IXMLDOMNode
    Dim oProp As IXMLDOMNode
    Dim oNamedItem As IJNamedItem
    Dim iCount As Integer
    Dim oTemplateSet As IJDMfgTemplateSet
    Dim oTemplate As IJMfgTemplate
    Dim oElements As IJElements
    Dim oSketchedElements As IJElements
    Dim oTemplateRpt As IJMfgTemplateReport
    Dim dVal As Double
    Dim oFrames As IJElements
    Dim oFrame As IHFrame
    Dim iSketchedTemplateCnt As Integer
    Dim sPlatePartName As String
    Dim dInterval As Double
    Dim dMaxLength As Double, dALength As Double, dBLength As Double
    Dim iNumberOfAInterval As Integer, iNumberOfBInterval As Integer
    Dim iCnt As Integer
    
    Set oTemplateSet = oObj
    
    If Not oTemplateSet Is Nothing Then
        
        Set oElements = oTemplateSet.GetTemplates
                
        For iCount = 1 To oElements.Count
        
            Set oTemplate = oTemplateSet.GetTemplateAtGivenIndex(iCount)
            
            'Create node "element"
            Set oElement = CreateChildNode(Parent, "element")
        
            'Add attribute elementName to element node and set value to to IJNamedItem.Name
            Set oNamedItem = oObj
            
            CHECK_POINT_STEP oNamedItem.name
            
            AddAttribute oElement, "elementName", oNamedItem.name & iCount
            Set oNamedItem = Nothing

            'Create node propertyList below element
            Set oPropList = CreateChildNode(oElement, "propertyList")

            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']")) Then
                AddAttribute oPropList, "propertyListName", "MfgTemplateList"
            End If

            'Create node propertyList below element
            Set oPropList = CreateChildNode(oElement, "propertyList")
            
            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                Set oProp = CreateChildNode(oPropList, "property")

                'Add attribute Type to property node
                AddAttribute oProp, "propertyName", "TEMPLATE"
                
                If iCount = 1 Then
                    AddAttribute oProp, "propertyValue", "A"
                ElseIf iCount = oElements.Count Then
                    AddAttribute oProp, "propertyValue", "F"
                Else
                    AddAttribute oProp, "propertyValue", CStr(iCount)
                End If
            End If
                                
            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                Set oProp = CreateChildNode(oPropList, "property")

                'Add attribute Type to property node
                AddAttribute oProp, "propertyName", "HEIGHT"
                
                dVal = GetMaxHeight(oTemplate)
                AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                AddAttribute oProp, "propertyType", "String"
                Set oProp = Nothing
            End If

            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                Set oProp = CreateChildNode(oPropList, "property")

                'Add attribute Type to property node
                AddAttribute oProp, "propertyName", "WIDTH"
                
                dVal = GetTopLineLength(oTemplate)
                AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                AddAttribute oProp, "propertyType", "String"
                Set oProp = Nothing
            End If

            Set oTemplate = Nothing
            Set oElement = Nothing
            Set oTemplateRpt = Nothing
            Set oPropList = Nothing
            
        Next
    End If
'm_objXMLDoc.Save App.Path & "\bslee.xml"
CleanUp:
    Set oElement = Nothing
    Set oPropList = Nothing
    Set oProp = Nothing
    Set oNamedItem = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Public Sub AddAttribute(Parent As IXMLDOMNode, AttName As String, AttValue As String)
    Const METHOD As String = "AddAttribute"
    On Error GoTo ErrorHandler
    Dim oATT As IXMLDOMAttribute
    Set oATT = m_objXMLDoc.createAttribute(AttName)
    Dim oNamedNodeMap As IXMLDOMNamedNodeMap
    Set oNamedNodeMap = Parent.Attributes
    oNamedNodeMap.setNamedItem oATT
    oATT.Value = AttValue

CleanUp:
    Set oATT = Nothing
    Set oNamedNodeMap = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Sub

Function CreateChildNode(Parentname As IXMLDOMNode, ChildName As String) As IXMLDOMNode
    Const METHOD As String = "CreateChildNode"
    On Error GoTo ErrorHandler
    Dim oChild As IXMLDOMNode
    'create childnode, and hook up to parent
    Set oChild = m_objXMLDoc.createElement(ChildName)
    Set oChild = Parentname.appendChild(oChild)
    Set CreateChildNode = oChild
    
CleanUp:
    Set oChild = Nothing
    
    Exit Function
    
ErrorHandler:
    Set CreateChildNode = Nothing
    Err.Raise Err.Number
    GoTo CleanUp
End Function

'This method getts the root node
Function GetRootNode(oDoc As DOMDocument) As IXMLDOMNode
    Const METHOD As String = "GetRootNode"
    On Error GoTo ErrorHandler

    Set GetRootNode = oDoc.firstChild 'msxml
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function


'This method creates a data element, sets the value and add it to the parent.
Function AddDataElement(Parent As IXMLDOMNode, name As String, Value As Variant) As IXMLDOMElement
    Const METHOD As String = "AddDataElement"
    On Error GoTo ErrorHandler

    Set AddDataElement = m_objXMLDoc.createElement(name)
    'AddDataElement.nodeValue = value
    Set AddDataElement = Parent.appendChild(AddDataElement)
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
        
    CHECK_POINT_ELEMENT "Gathering elements... "
    
    If pElements.Count > 0 Then
    
        '- create the xmldom document
        Set m_objXMLDoc = New DOMDocument
        m_objXMLDoc.loadXML strFileName
        
        'load the xmldocument and create headerinformation
        CreateHeader "CTemplateSet Serializer", "CTemplateSet"
        
        'call methods to read from objects and add to xml
        SortTemplateElements pElements
        m_objXMLDoc.Save strFileName
        eCustomReportStatus = StrMfgProcessFinished
    End If
    
WrapUp:
    Exit Sub
    
ErrorHandler:
    eCustomReportStatus = StrMfgErrorUnknown
    Err.Raise Err.Number
    GoTo WrapUp
End Sub

Private Function CHECK_REPORTABILITY(pattern As String) As Boolean
    ' Default True
    CHECK_REPORTABILITY = True
    
    If m_objFilter Is Nothing Then GoTo CleanUp
    
    On Error GoTo CleanUp
    
    Dim oInNode As IXMLDOMNode
    Set oInNode = m_objFilter.selectSingleNode("//" & pattern)
    
    ' The node coming in is not recognised
    If oInNode Is Nothing Then GoTo CleanUp
    
    ' Get the report attribute which is a sibling of the incoming node
    ' and check for its value. If not present: return true as usual (Error Case)
    ' If we are able to clearly identify that NO is marked then return False
    If oInNode.parentNode.Attributes.getNamedItem("report").nodeValue = "NO" Then CHECK_REPORTABILITY = False
    
CleanUp:
    ' Worst case the Report must be printed out : So return True
End Function

Private Function GetTopLineLength(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetTopLineLength"
    On Error GoTo ErrorHandler
    
    Dim oLine As IJLine

    Set oLine = oTemplate.TopLine
    GetTopLineLength = oLine.Length
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetMaxHeight(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetMaxHeight"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    
    Set oTemplateRpt = oTemplate
    GetMaxHeight = oTemplateRpt.GetMaxHeightOfTemplate
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Sub CHECK_POINT_PHASE(name As Variant)
'    If Not m_oDwgProgress Is Nothing Then
'        m_oDwgProgress.AtPhase name
'    End If
End Sub

Private Sub CHECK_POINT_ELEMENT(name As Variant)
'    If Not m_oDwgProgress Is Nothing Then
'        m_oDwgProgress.AtElement name
'    End If
End Sub

Private Sub CHECK_POINT_STEP(name As Variant)
'    If Not m_oDwgProgress Is Nothing Then
'        m_oDwgProgress.AtStep name
'    End If
End Sub

Private Sub SET_PROGRESS_INFO(stepCt As Long, weightage As Double, incrementStatus As Long)
'    Dim oShipDwgProgress As IJDShipDwgProgress
'
'    If Not m_oDwgProgress Is Nothing Then
'        Set oShipDwgProgress = m_oDwgProgress
'
'        If Not oShipDwgProgress Is Nothing Then
'            oShipDwgProgress.SetProgressInfo stepCt, weightage, incrementStatus
'        End If
'    End If
End Sub
