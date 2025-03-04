VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IngrTemplateSet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CIngrTemplateSet3.cls
''
''  Description : Populates the XML DOM Document with TemplateSet related data for Sample 3
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

Private Const MODULE = "CustomReports.CIngrTemplateSet 3: "

Private m_objXMLDoc             As DOMDocument
Private m_oCurrentRootNode      As IXMLDOMNode
Private m_oFileRootNode         As IXMLDOMNode

'Private m_oDwgProgress         As IJDDwgProgress

' Variable that stores the Filtering Info.
Private m_objFilter             As IXMLDOMElement
Public Enum eRefCurve
    BottomFlatLine = 0
    SideParallelLine
    ArcEndingLine
End Enum
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
    
    Dim oColItem As Object
    Dim oTemplateSet As IJDMfgTemplateSet
    Dim lCount As Long
    
    SET_PROGRESS_INFO oInElements.Count, 2, 1

    'Step through all elements
    For lCount = 1 To oInElements.Count
        
        CHECK_POINT_ELEMENT oInElements(lCount)
        
        If TypeOf oInElements.Item(lCount) Is IJDMfgTemplateSet Then
            'Step through collection and get the objects which implemnts IJProfilePart
            'The following objects does: ProfilePart, StiffenerPart, BeamPart
            For Each oColItem In oInElements

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
    Next

CleanUpp:
    Set oColItem = Nothing
    Set oTemplateSet = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUpp
    
End Sub

Private Function ExtractTemplateSetToXML(oObj As IUnknown, Parent As IXMLDOMNode) As Boolean
    Const METHOD As String = "ExtractTemplateSetToXML"
    On Error Resume Next

    Dim oElement As IXMLDOMNode
    Dim oPropList As IXMLDOMNode
    Dim oProp As IXMLDOMNode
    Dim oNamedItem As IJNamedItem
    Dim iCount As Integer
    Dim iNonSketchedTemplateCount As Integer
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
    Dim bRealData As Boolean
'    Dim iCnt As Integer
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Dim sViewDirection As String
    
    Set oTemplateSet = oObj
    
    If Not oTemplateSet Is Nothing Then
        
        Set oElements = oTemplateSet.GetTemplates
        
        'Check what type of templateset
        Set oProcessSettings = oTemplateSet.GetProcessSettings
        sViewDirection = oProcessSettings.TemplateDirection
        
        'Get Distance
        GetDistanceForDrawing oTemplateSet, dMaxLength, dALength, dBLength
        
        'Get interval value
        dInterval = oTemplateSet.GetTemplateOffset
        
        'Get numer of intervals
        GetIntervals dALength, dBLength, dInterval, iNumberOfAInterval, iNumberOfBInterval
        
        'Get platepart name
        sPlatePartName = GetPlatePartName(oTemplateSet)
        
        'Get SketchedElements from TemplateSet
        Set oSketchedElements = oTemplateSet.GetSketchingTemplates
        iSketchedTemplateCnt = oSketchedElements.Count
        
        iNonSketchedTemplateCount = oElements.Count - iSketchedTemplateCnt
        
        'Get frames from Templateset
        Set oFrames = GetFrames(oTemplateSet)
        
        If iNonSketchedTemplateCount = oFrames.Count Then
            bRealData = False
        Else
            bRealData = False
        End If
'        iCnt = 1
        
        For iCount = 1 To iNonSketchedTemplateCount
        
            Set oTemplate = oTemplateSet.GetTemplateAtGivenIndex(iCount)
            
            Set oTemplateRpt = oTemplate
            
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
                AddAttribute oProp, "propertyName", "FR.NO"
                
'                If iCount <= oElements.Count - iSketchedTemplateCnt Then
                    'Get each frame
                If bRealData = True Then
                    If iCount = 1 Then
                        Set oFrame = oFrames.Item(iCount)
                        AddAttribute oProp, "propertyValue", oFrame.name & " - " & GetAftButtName(oTemplateSet)
                    ElseIf iCount = iNonSketchedTemplateCount Then
                        Set oFrame = oFrames.Item(iCount - 2)
                        AddAttribute oProp, "propertyValue", oFrame.name & " - " & GetForeButtName(oTemplateSet)
                    Else
                        Set oFrame = oFrames.Item(iCount - 1)
                        AddAttribute oProp, "propertyValue", oFrame.name
                    End If
                Else
                    Set oFrame = oFrames.Item(iCount)
                    If iCount = 1 Then
                        AddAttribute oProp, "propertyValue", oFrame.name & " - " & GetAftButtName(oTemplateSet)
                    ElseIf iCount = iNonSketchedTemplateCount Then
                        AddAttribute oProp, "propertyValue", oFrame.name & " - " & GetForeButtName(oTemplateSet)
                    Else
                        AddAttribute oProp, "propertyValue", oFrame.name
                    End If
                End If
'                Else
'                    AddAttribute oProp, "propertyValue", sPlatePartName & "Sketched" & iCnt
'                    iCnt = iCnt + 1
'                End If
                AddAttribute oProp, "propertyType", "String"
                Set oProp = Nothing
            End If
                                
            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                Set oProp = CreateChildNode(oPropList, "property")

                'Add attribute Type to property node
                AddAttribute oProp, "propertyName", "ANGLE"
                
                dVal = GetAngle(oTemplate)
                AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                AddAttribute oProp, "propertyType", "String"
                Set oProp = Nothing
            End If
    
            If sViewDirection = "Transversal" Then 'For Frame Template
            
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", GetLowerSeamName(oTemplateSet) & "(L*M)"
                    
                    dVal = GetLSeamNBCtlLine(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2)) '
                    AddAttribute oProp, "propertyType", "String"
                    Set oProp = Nothing
                End If
    
    
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", GetLowerSeamName(oTemplateSet) & "(L)"
                                    
                    dVal = GetDistBtnLSeamAndClosest(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2)) '
                    AddAttribute oProp, "propertyType", "String"
                    Set oProp = Nothing
                End If
    
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", GetUpperSeamName(oTemplateSet) & "(U*M)"
                    
                    dVal = GetUSeamNBCtlLine(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
                    Set oProp = Nothing
                End If
    
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", GetUpperSeamName(oTemplateSet) & "(U)"
                    
                    dVal = GetDistBtnUSeamAndClosest(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
                    Set oProp = Nothing
                End If
    
    
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", "BK"
                                        
                    Dim dBK As Double
                    
                    dBK = GetBackSet(oTemplate)
                    
                    If dBK <> 1000000 Then
                        dVal = dBK
                        AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    Else
                        AddAttribute oProp, "propertyValue", ""
                    End If
                    
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
            End If

            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                Set oProp = CreateChildNode(oPropList, "property")

                'Add attribute Type to property node
                
                Dim dRBB As Double
                Dim eRCurve As eRefCurve
                GetRBB oTemplate, dRBB, eRCurve
                
                Select Case eRCurve
                    Case BottomFlatLine
                        AddAttribute oProp, "propertyName", "RB.B"
                    Case SideParallelLine
                        AddAttribute oProp, "propertyName", "RS.B"
                    Case ArcEndingLine
                        AddAttribute oProp, "propertyName", "RW.B"
                End Select
                               
                If dRBB <> 1000000 Then
                    dVal = dRBB
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2)) '"rbb"
                Else
                    AddAttribute oProp, "propertyValue", ""
                End If
                
                AddAttribute oProp, "propertyType", "String"

                Set oProp = Nothing
            End If
            
            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                Set oProp = CreateChildNode(oPropList, "property")

                'Add attribute Type to property node
                
                Dim dRBH As Double
                GetRBH oTemplate, dRBH, eRCurve
                
                Select Case eRCurve
                    Case BottomFlatLine
                        AddAttribute oProp, "propertyName", "RB.H"
                    Case SideParallelLine
                        AddAttribute oProp, "propertyName", "RS.H"
                    Case ArcEndingLine
                        AddAttribute oProp, "propertyName", "RW.H"
                End Select
                                
                
                If dRBH <> 1000000 Then
                    dVal = dRBH
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2)) '"rbh"
                Else
                    AddAttribute oProp, "propertyValue", ""
                End If
                
                AddAttribute oProp, "propertyType", "String"

                Set oProp = Nothing
            End If

            If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                Set oProp = CreateChildNode(oPropList, "property")
                
                Dim iAInterval As Integer, iBInterval As Integer
                GetIndexAtBaseControlLine oTemplate, dInterval, iAInterval, iBInterval

                'Add attribute Type to property node
                AddAttribute oProp, "propertyName", "StandardPointNumber"
                AddAttribute oProp, "propertyValue", CStr(iAInterval)
                AddAttribute oProp, "propertyType", "String"
                Set oProp = Nothing
            End If
            
            Dim i As Integer
            Dim dRealInterval As Double
            Dim iTotalInterval As Integer
            Dim dEdgeADistance As Double
            Dim dOffSet As Double
            Dim bCheck As Boolean
            
            dOffSet = 0
            
            bCheck = False

            iTotalInterval = iNumberOfAInterval + iNumberOfBInterval
            
            Dim dTemplateLength As Double, dTemplateALength As Double, dTemplateBLength As Double
            Dim dNumberOfIntervalA As Integer, dNumberOfIntervalB As Integer
            
            dTemplateLength = GetTopLineLength(oTemplate)
            
            dTemplateALength = oTemplateRpt.GetDistFromStartingEdgeToBaseCtlLine
            
            dTemplateBLength = dTemplateLength - dTemplateALength

            GetIntervals dTemplateALength, dTemplateBLength, dInterval, dNumberOfIntervalA, dNumberOfIntervalB

            dEdgeADistance = dTemplateALength - ((dNumberOfIntervalA - 1) * dInterval)
            
            For i = 0 To iTotalInterval
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")

                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", CStr(i)
                    
                    If i = 0 Then
                        dRealInterval = 0
                    ElseIf i = 1 Then
                        dRealInterval = dEdgeADistance
                    Else
                        dRealInterval = dInterval
                    End If
                    
                    dOffSet = dOffSet + dRealInterval
                    
                    If bCheck = False Then
                        If dOffSet > dTemplateLength Then
                            dOffSet = dTemplateLength
                            bCheck = True
                        End If
                        dVal = GetOffSetValue(oTemplate, dOffSet)
                        AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    Else
                        AddAttribute oProp, "propertyValue", ""
                    End If
                    
                    AddAttribute oProp, "propertyType", "String"

                    Set oProp = Nothing
                End If
            Next
            
            If sViewDirection = "Transversal" Then
            
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    Dim strLowerSeamName As String
                    strLowerSeamName = GetLowerSeamName(oTemplateSet)
                    If strLowerSeamName = "" Then strLowerSeamName = "LSeam"
                    AddAttribute oProp, "propertyName", strLowerSeamName
                    
                    dVal = GetGirthLengthAtLSeam(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
                
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    Dim strUpperSeamName As String
                    strUpperSeamName = GetUpperSeamName(oTemplateSet)
                    If strUpperSeamName = "" Then strUpperSeamName = "USeam"
                    AddAttribute oProp, "propertyName", strUpperSeamName
                    
                    dVal = GetGirthLengthAtUSeam(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
                
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", "A.LINE"
                    
                    dVal = GetGirthLengthAtBaseCtlLine(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
                
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", "L-M"
                    
                    dVal = GetGirthLengthBtwLSeamAndBCtlLine(oTemplate)
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", "L-U"
                    
                    dVal = GetGirthLengthBtwUSeamAndBCtlLine(oTemplate)
                    
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
                
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", "F"
                    
                    dVal = GetMaxHiFromButt(oTemplate)
                    
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2))
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
                
                Dim dHLVal As Double, dHUVal As Double
                
                GetHLU oTemplate, dHUVal, dHLVal
                
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", "HL"
                    
                    dVal = dHLVal
                    
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2)) '"hl" '
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
                If (CHECK_REPORTABILITY("@propertyListName[.='Contents']//@propertyName[.='Type']")) Then
                    Set oProp = CreateChildNode(oPropList, "property")
    
                    'Add attribute Type to property node
                    AddAttribute oProp, "propertyName", "HU"
                    
                    dVal = dHUVal
                    
                    AddAttribute oProp, "propertyValue", CStr(Round(dVal, 2)) '"hu" '
                    AddAttribute oProp, "propertyType", "String"
    
                    Set oProp = Nothing
                End If
            End If
            
            Set oTemplate = Nothing
            Set oElement = Nothing
            Set oTemplateRpt = Nothing
            Set oPropList = Nothing
            
        Next
    End If
'm_objXMLDoc.save App.Path & "\bslee.xml"
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

Private Sub AddAttribute(Parent As IXMLDOMNode, AttName As String, AttValue As String)
    Const METHOD As String = "AddAttribute"
    On Error GoTo ErrorHandler
    
    Dim oATT As IXMLDOMAttribute
    Dim oNamedNodeMap As IXMLDOMNamedNodeMap
    
    Set oATT = m_objXMLDoc.createAttribute(AttName)
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

Private Function CreateChildNode(Parentname As IXMLDOMNode, ChildName As String) As IXMLDOMNode
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
Private Function GetRootNode(oDoc As DOMDocument) As IXMLDOMNode
    Const METHOD As String = "GetRootNode"
    On Error GoTo ErrorHandler

    Set GetRootNode = oDoc.firstChild 'msxml
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function


'This method creates a data element, sets the value and add it to the parent.
Private Function AddDataElement(Parent As IXMLDOMNode, name As String, Value As Variant) As IXMLDOMElement
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
    Const METHOD = "IJDSerializer_Update"
    On Error GoTo ErrorHandler
        
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

Private Function GetInterval(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetInterval"
    On Error GoTo ErrorHandler
    
    Dim oTopLine As IJLine
    Set oTopLine = oTemplate.TopLine

    GetInterval = oTopLine.Length / 8

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Sub InitUOMService()
    Const METHOD As String = "InitUOMService"
    On Error GoTo ErrorHandler

'    Set m_oUnitFormat = New UomFormat
'    Set m_oUnitsOfMeasure = m_oTrader.Service(TKUnitsOfMeasure, "")
'
'    m_oUnitFormat.PrecisionType = PRECISIONTYPE_DECIMAL
'    m_oUnitFormat.DecimalPrecision = 2
'    m_oUnitFormat.UnitsDisplayed = True
'
'    m_oUnitsOfMeasure.GetDefaultUnits m_eUnitType, m_CurrentDefaultUnits

    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Function GetGirthLengthAtLSeam(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetGirthLengthAtLSeam"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    
    Set oTemplateRpt = oTemplate
    GetGirthLengthAtLSeam = oTemplateRpt.GetGirthLengthAtLSeam
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetGirthLengthAtUSeam(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetGirthLengthAtUSeam"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    
    Set oTemplateRpt = oTemplate
    GetGirthLengthAtUSeam = oTemplateRpt.GetGirthLengthAtUSeam
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetGirthLengthAtBaseCtlLine(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetGirthLengthAtBaseCtlLine"
    On Error GoTo ErrorHandler

    Dim oTemplateRpt As IJMfgTemplateReport

    Set oTemplateRpt = oTemplate
    GetGirthLengthAtBaseCtlLine = oTemplateRpt.GetGirthLengthAtBaseCtlLine

    Exit Function

ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetGirthLengthBtwLSeamAndBCtlLine(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetGirthLengthBtwLSeamAndBCtlLine"
    On Error GoTo ErrorHandler

    Dim oTemplateRpt As IJMfgTemplateReport

    Set oTemplateRpt = oTemplate
    GetGirthLengthBtwLSeamAndBCtlLine = oTemplateRpt.GetGirthLengthBtwLSeamAndBaseCtlLine

    Exit Function

ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetGirthLengthBtwUSeamAndBCtlLine(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetGirthLengthBtwUSeamAndBCtlLine"
    On Error GoTo ErrorHandler

    Dim oTemplateRpt As IJMfgTemplateReport

    Set oTemplateRpt = oTemplate
    GetGirthLengthBtwUSeamAndBCtlLine = oTemplateRpt.GetGirthLengthBtwUSeamAndBaseCtlLine

    Exit Function

ErrorHandler:
    Err.Raise Err.Number
End Function

Private Sub GetRBB(oTemplate As IJMfgTemplate, dRBreadth As Double, eRCurve As eRefCurve)
    Const METHOD As String = "GetRBB"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Dim bExist As Boolean
    Dim dRBB As Double
    eRCurve = BottomFlatLine
    
    Set oTemplateRpt = oTemplate
    oTemplateRpt.GetBottomFlatLineBreadth bExist, dRBB
    If bExist = True Then
        dRBreadth = dRBB
        eRCurve = BottomFlatLine
    Else
        oTemplateRpt.GetSideParallelLineBreadth bExist, dRBB
        If bExist = True Then
            dRBreadth = dRBB
            eRCurve = SideParallelLine
        Else
            oTemplateRpt.GetEndingLineOfWLBreadth bExist, dRBB
            If bExist = True Then
                dRBreadth = dRBB
                eRCurve = ArcEndingLine
            Else
                dRBreadth = 1000000
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub GetRBH(oTemplate As IJMfgTemplate, dRHeight As Double, eRCurve As eRefCurve)
    Const METHOD As String = "GetRBH"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Dim bExist As Boolean
    Dim dRBH As Double
    eRCurve = BottomFlatLine
    
    Set oTemplateRpt = oTemplate
    oTemplateRpt.GetBottomFlatLineHeight bExist, dRBH
    If bExist = True Then
        dRHeight = dRBH
        eRCurve = BottomFlatLine
    Else
        oTemplateRpt.GetSideParallelLineHeight bExist, dRBH
        If bExist = True Then
            dRHeight = dRBH
            eRCurve = SideParallelLine
        Else
            oTemplateRpt.GetEndingLineOfWLHeight bExist, dRBH
            If bExist = True Then
                dRHeight = dRBH
                eRCurve = ArcEndingLine
            Else
                dRHeight = 1000000
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub GetHLU(oTemplate As IJMfgTemplate, dHU As Double, dHL As Double)
    Const METHOD As String = "GetHLU"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport

    Set oTemplateRpt = oTemplate
    oTemplateRpt.GetHeightsFromGroundFloor dHU, dHL
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Function GetHU(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetHU"
    On Error GoTo ErrorHandler
    
    
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetBackSet(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetBackSet"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    
    Set oTemplateRpt = oTemplate
    GetBackSet = oTemplateRpt.GetBackSet
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetMaxHiFromButt(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetMaxHiFromButt"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    
    Set oTemplateRpt = oTemplate
    GetMaxHiFromButt = oTemplateRpt.GetMaxHeightFromButtLine
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetLSeamNBCtlLine(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetLSeamNBCtlLine"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Set oTemplateRpt = oTemplate
    GetLSeamNBCtlLine = oTemplateRpt.GetDistBtwLSeamAndBaseCtlLine
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetUSeamNBCtlLine(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetUSeamNBCtlLine"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Set oTemplateRpt = oTemplate
    GetUSeamNBCtlLine = oTemplateRpt.GetDistBtwUSeamAndBaseCtlLine
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetDistBtnLSeamAndClosest(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetDistBtnLSeamAndClosest"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Set oTemplateRpt = oTemplate
    GetDistBtnLSeamAndClosest = oTemplateRpt.GetDistBtwLSeamAndClosestTemplate
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetDistBtnUSeamAndClosest(oTemplate As IJMfgTemplate) As Double
    Const METHOD As String = "GetDistBtnUSeamAndClosest"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Set oTemplateRpt = oTemplate
    GetDistBtnUSeamAndClosest = oTemplateRpt.GetDistBtwUSeamAndClosestTemplate
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetOffSetValue(oTemplate As IJMfgTemplate, dInterval As Double) As Double
    Const METHOD As String = "GetOffSetValue"
    On Error GoTo ErrorHandler
    
    Dim oTemplateRpt As IJMfgTemplateReport
    Set oTemplateRpt = oTemplate
    
    GetOffSetValue = oTemplateRpt.GetTemplateHeightByOffset(dInterval)
    

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function CheckMaxLength(oTemplateSet As IJDMfgTemplateSet) As Double
    Const METHOD As String = "CheckMaxLength"
    On Error GoTo ErrorHandler
    
    Dim oTemplate As IJMfgTemplate
    Dim oElements As IJElements
    Dim oElement As Object
    Dim iCount As Integer
    Dim oLine As IJLine
    Dim dLength As Double
    
    Set oElements = oTemplateSet.GetTemplates

    For Each oElement In oElements
        If TypeOf oElement Is IJMfgTemplate Then
            Set oTemplate = oElement
            Set oLine = oTemplate.TopLine
            If oLine.Length > dLength Then
                dLength = oLine.Length
            End If
        End If
        Set oLine = Nothing
        Set oTemplate = Nothing
    Next
    
    CheckMaxLength = dLength
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetUpperSeamName(oTemplateSet As IJDMfgTemplateSet) As String
    Const METHOD As String = "GetUpperSeamName"
    On Error GoTo ErrorHandler

    Dim oTemplateSetRpt As IJMfgTemplateSetReport
    Dim oPlatePart As IJPlatePart

    Set oTemplateSetRpt = oTemplateSet
    Set oPlatePart = oTemplateSetRpt.GetShellPlatePart

    GetUpperSeamName = oTemplateSetRpt.GetUpperSeamName(oPlatePart, IsBaseSide(oTemplateSet))

    Exit Function

ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetLowerSeamName(oTemplateSet As IJDMfgTemplateSet) As String
    Const METHOD As String = "GetLowerSeamName"
    On Error GoTo ErrorHandler

    Dim oTemplateSetRpt As IJMfgTemplateSetReport
    Dim oPlatePart As IJPlatePart

    Set oTemplateSetRpt = oTemplateSet
    Set oPlatePart = oTemplateSetRpt.GetShellPlatePart

    GetLowerSeamName = oTemplateSetRpt.GetLowerSeamName(oPlatePart, IsBaseSide(oTemplateSet))

    Exit Function

ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetPlatePartName(oTemplateSet As IJDMfgTemplateSet) As String
    Const METHOD As String = "GetPlatePartName"
    On Error GoTo ErrorHandler
    
    Dim oTemplateSetRpt As IJMfgTemplateSetReport
    Dim oNamedItem As IJNamedItem
    Dim oPlatePart As IJPlatePart
    
    Set oTemplateSetRpt = oTemplateSet
    Set oPlatePart = oTemplateSetRpt.GetShellPlatePart
    Set oNamedItem = oPlatePart
    
    GetPlatePartName = oNamedItem.name
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Sub GetDistanceForDrawing(oTemplateSet As IJDMfgTemplateSet, dMaxLength As Double, dALength As Double, dBLength As Double)
    Const METHOD As String = "GetDistanceForDrawing"
    On Error GoTo ErrorHandler
    
    Dim oTemplate As IJMfgTemplate
    Dim oElements As IJElements
    Dim oElement As Object
    Dim iCount As Integer
    Dim oLine As IJLine
    Dim oMaxTemplate As IJMfgTemplate
    Dim oTemplateRpt As IJMfgTemplateReport
    
    Dim dLength As Double
    
    Set oElements = oTemplateSet.GetTemplates

    For Each oElement In oElements
        If TypeOf oElement Is IJMfgTemplate Then
            Set oTemplate = oElement
            Set oLine = oTemplate.TopLine
            If oLine.Length > dLength Then
                dLength = oLine.Length
                Set oMaxTemplate = Nothing
                Set oMaxTemplate = oTemplate
            End If
        End If
        Set oLine = Nothing
        Set oTemplate = Nothing
    Next
    
    'Get max length
    Set oLine = oMaxTemplate.TopLine
    dMaxLength = oLine.Length
    
    'Get ADistance
    Set oTemplateRpt = oMaxTemplate
    dALength = oTemplateRpt.GetDistFromStartingEdgeToBaseCtlLine

    'Get BDistance
    dBLength = dMaxLength - dALength
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub GetIntervals(dALength As Double, dBLength As Double, dInterval As Double, iNumberOfAInterval As Integer, iNumberOfBInterval As Integer)
    Const METHOD As String = "GetIntervals"
    On Error GoTo ErrorHandler

    Dim iRemainder As Integer

    iRemainder = (dALength * 100) Mod (dInterval * 100)

    If iRemainder = 0 Then
        iNumberOfAInterval = (dALength * 100) \ (dInterval * 100)
    Else
        iNumberOfAInterval = (dALength * 100) \ (dInterval * 100) + 1
    End If

    iRemainder = (dBLength * 100) Mod (dInterval * 100)
    
    If iRemainder = 0 Then
        iNumberOfBInterval = (dBLength * 100) \ (dInterval * 100)
    Else
        iNumberOfBInterval = (dBLength * 100) \ (dInterval * 100) + 1
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

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

Private Sub GetIndexAtBaseControlLine(oTemplate As IJMfgTemplate, ByVal dInterval As Double, ByRef iAInterval As Integer, ByRef iBInterval As Integer)
    Const METHOD As String = "GetIndexAtBaseControlLine"
    On Error GoTo ErrorHandler

    Dim dLocalTemplateLength As Double, dALocalLength As Double, dBLocalLength As Double
    Dim oTemplateRpt As IJMfgTemplateReport
    
    Set oTemplateRpt = oTemplate
    dLocalTemplateLength = GetTopLineLength(oTemplate)
    dALocalLength = oTemplateRpt.GetDistFromStartingEdgeToBaseCtlLine
    dBLocalLength = dLocalTemplateLength - dALocalLength

    GetIntervals dALocalLength, dBLocalLength, dInterval, iAInterval, iBInterval

Exit Sub
ErrorHandler:
    Err.Raise Err.Number
End Sub

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


 
