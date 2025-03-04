VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIngrPinJig3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CIngrPinJig3.cls
''
''  Description : Populates the XML DOM Document with PinJig related data for Sample 3
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

Private Const MODULE = "::CustomReports.CIngrPinJigXML3"

Private m_objXMLDoc             As DOMDocument
Private m_objElements           As IJElements

'root node of each table
Private m_oAngleNode            As IXMLDOMNode
Private m_oDiagonalNode         As IXMLDOMNode
Private m_o1stRootNode          As IXMLDOMNode
Private m_o2ndRootNode          As IXMLDOMNode
Private m_o3rdRootNode          As IXMLDOMNode
Private m_o4thRootNode          As IXMLDOMNode

'Surface diagonal length
Private DistALtoFU              As Double
Private DistAUtoFL              As Double

Dim sVertNames()                As String
Dim sHorNames()                 As String

Private m_oFileRootNode         As IXMLDOMNode

' Variable that stores the Filtering Info.
Private m_objFilter             As IXMLDOMElement

'Type for marking data
Private Type MarkingDataInfo
    oMfgPin                  As IJMfgPin
    RemarkingHorCurveName    As String
    RemarkingVertCurveName   As String
    dx                       As Double
    dy                       As Double
    Height                   As Double
    dCurve                   As Double
    bFlag                    As Boolean
    HorLength                As Double
    VertLength               As Double
    AFAngle                  As Double
    FOAngle                  As Double
End Type

Private mRemarkingData()     As MarkingDataInfo

'Private m_oDwgProgress         As IJDDwgProgress

Private Sub Class_Initialize()
    InitializeFilter m_objFilter
End Sub

Private Sub Class_Terminate()
    Set m_objElements = Nothing
    Set m_objXMLDoc = Nothing
    Set m_o1stRootNode = Nothing
    Set m_o2ndRootNode = Nothing
    Set m_o3rdRootNode = Nothing
    Set m_o4thRootNode = Nothing
    Set m_oFileRootNode = Nothing

End Sub


Private Sub GetAllPinAndCreateXMLDocument(oInElements As IJElements)
    Const METHOD = "GetAllPinAndCreateXMLDocument"
    On Error GoTo ErrorHandler
    
    Dim oPinJig As IJPinJig
    Dim oMfgParent As IJMfgGeomParent
    Dim oColl As IJDTargetObjectCol
    Dim oProjectedData As IJJigProjectedData
    Dim oPin As IJMfgPin
    
    Dim iCnt As Integer
    
    Set oPinJig = oInElements.Item(1)
    Set oMfgParent = oPinJig
    
    Set oColl = oMfgParent.GetChildren()
    
    
    'Get last two objects(ProjectedData and PinSet object)
    For iCnt = 1 To oColl.Count
        If TypeOf oColl.Item(iCnt) Is IJJigOutput Then
            'This will be JigOutPut object
            Set oMfgParent = oColl.Item(iCnt)
            Set oColl = oMfgParent.GetChildren()
        End If
    Next
    
    'Now oColl has 'Projected Data' and 'PinSet' object
    For iCnt = 1 To oColl.Count
        If TypeOf oColl.Item(iCnt) Is IJMfgPinSet Then
            Set oMfgParent = oColl.Item(iCnt)
        End If
    Next

    'Get all pins
    Set oColl = oMfgParent.GetChildren()
'    MsgBox oColl.Count
    
    'Make xml document
    Dim i As Integer
    For i = 1 To oColl.Count
        ExtractXMLFromPin oColl.Item(i), m_o1stRootNode
    Next

WrapUp:
    Set oMfgParent = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    'ReportAndRaiseUnanticipatedError METHOD, MODULE
    GoTo WrapUp
End Sub

Private Sub ExtractXMLFromPin(oPin As IJMfgPin, Parent As IXMLDOMNode)
    On Error GoTo ErrorHandler
    Const METHOD = "ExtractXMLFromPin"
    
    Dim oElement As IXMLDOMNode
    Dim oPropList As IXMLDOMNode
    Dim oProp1 As IXMLDOMNode
    Dim oProp2 As IXMLDOMNode
    Dim oProp3 As IXMLDOMNode
    Dim oProp4 As IXMLDOMNode
    Dim oProp5 As IXMLDOMNode
    'Dim FileName As String
    Dim oName As IJNamedItem
    
    Set oElement = CreateChildNode(Parent, "element")
    
    Set oName = oPin
    AddAttribute oElement, "elementName", oName.name
    Set oName = Nothing
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    
    If (CHECK_REPORTABILITY("@propertyListName[.='Frame']")) Then
        AddAttribute oPropList, "PropertyListName", "Frame"
        If (CHECK_REPORTABILITY("@propertyListName[.='Frame']//@propertyName[.='Type']")) Then
            Set oProp1 = CreateChildNode(oPropList, "property")
            Set oProp2 = CreateChildNode(oPropList, "property")
            Set oProp3 = CreateChildNode(oPropList, "property")
            Set oProp4 = CreateChildNode(oPropList, "property")
            Set oProp5 = CreateChildNode(oPropList, "property")
'''
'''                 'Add attribute Type to property node
                AddAttribute oProp1, "propertyName", "UpperSeam"
                AddAttribute oProp2, "propertyName", "MiddleSeam1"
                AddAttribute oProp3, "propertyName", "MiddleSeam2"
                AddAttribute oProp4, "propertyName", "MiddleSeam3"
                AddAttribute oProp5, "propertyName", "LowerSeam"
                
'''                AddAttribute oProp, "propertyValue", "111"
'''                AddAttribute oProp, "propertyType", "String"
'''                Set oProp = Nothing
        End If
    End If
    'FileName = Environ("TEMP")
    'If FileName = "" Or FileName = vbNullString Then
    '    FileName = "c:\temp" 'Only use C:\Temp if there is a %TEMP% failure
    'End If
    'FileName = FileName & "\aa.xml"
    'm_objXMLDoc.Save FileName
WrapUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo WrapUp
End Sub

'*************************************************************
'************ UTILITIES **************************************
'*************************************************************



'******************************************************************
' Method : AddAttribute
'
' Description:
'   This method create an attribute assigns a value and link it to the parent node.
'
' Arguments:
'   [in] Parent As IXMLDOMNode, the node to be parent of the attribute
'   [in] AttName As String,     the name of the attribute
'   [out] AttValue As String    the value of the attribute (can be blank)
'
'   Return Values:
'
'******************************************************************
Public Sub AddAttribute(Parent As IXMLDOMNode, AttName As String, AttValue As String)
    Const METHOD As String = "AddAttribute"
    On Error GoTo ErrorHandler
    Dim oATT As IXMLDOMAttribute
    Set oATT = m_objXMLDoc.createAttribute(AttName)
    Dim oNamedNodeMap As IXMLDOMNamedNodeMap
    Set oNamedNodeMap = Parent.Attributes
    oNamedNodeMap.setNamedItem oATT
    oATT.Value = AttValue

Cleanup:
    Set oATT = Nothing
    Set oNamedNodeMap = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
    
End Sub


'******************************************************************
' Method : CreateChildNode
'
' Description:
'   This method creates an node (of type NODE_ELEMENT) and links it to the
'   parent node.
'
' Arguments:
'   [in] Parent As IXMLDOMNode, the node to be parent of the new Node
'   [in] NodeName As String,    the name new node
'
'   Return Values:
'
'******************************************************************
Function CreateChildNode(Parentname As IXMLDOMNode, ChildName As String) As IXMLDOMNode
    Const METHOD As String = "CreateChildNode"
    On Error GoTo ErrorHandler
    Dim oChild As IXMLDOMNode
    'create childnode, and hook up to parent
    Set oChild = m_objXMLDoc.createElement(ChildName)
    Set oChild = Parentname.appendChild(oChild)
    Set CreateChildNode = oChild
    
Cleanup:
    Set oChild = Nothing
    Exit Function
    
ErrorHandler:
    Set CreateChildNode = Nothing
    Err.Raise Err.Number
    GoTo Cleanup

End Function


'******************************************************************
' Method : CreateHeader
'
' Description:
'   This method creates the header, the first (top) node (report) and
'   the second node (elementList). It adds attributes to the first node.
'   XML-hierarchy:
'
'   report
'      |- Attribute reportName
'      |- Attribute reportUID
'      |- elementList
'           |-
'         All other stuff appears below this node.
'
' Arguments:
'   [in] sReportName As String, Name of report
'   [in] sReportUID As String,  Idendifier of report
'
'   Return Values:
'
'******************************************************************
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
    
    'Make angle table
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "TR-WL"
    AddAttribute oElementList, "LocReport", "In"
    Set m_oAngleNode = oElementList
    
    'Make diagonal table
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "Diagonal Length"
    AddAttribute oElementList, "LocReport", "In"
    Set m_oDiagonalNode = oElementList
    
    'Make 1st element list for 1st table
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "1stTABLE"
    AddAttribute oElementList, "LocReport", "Out"
    Set m_o1stRootNode = oElementList
    
'    Make 2nd element list for 2nd table
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "2ndTABLE"
    AddAttribute oElementList, "LocReport", "Out"
    Set m_o2ndRootNode = oElementList

'    Make 3rd element list for 3rd table
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "3rdTABLE"
    AddAttribute oElementList, "LocReport", "Out"
    Set m_o3rdRootNode = oElementList

    'Make 4th element list for 4th table
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "4thTABLE"
    AddAttribute oElementList, "LocReport", "Out"
    Set m_o4thRootNode = oElementList
    
    
    Set oElementList = Nothing
    
Cleanup:
    Set oElementList = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
    
End Sub


'This method getts the root node
Function GetRootNode(oDoc As DOMDocument) As IXMLDOMNode
    Set GetRootNode = oDoc.firstChild 'msxml
End Function


'This method creates a data element, sets the value and add it to the parent.
Function AddDataElement(Parent As IXMLDOMNode, name As String, Value As Variant) As IXMLDOMElement
    Set AddDataElement = m_objXMLDoc.createElement(name)
    'AddDataElement.nodeValue = value
    Set AddDataElement = Parent.appendChild(AddDataElement)
End Function

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
    
    CHECK_POINT_ELEMENT "Gathering elements... "
    
    If pElements.Count > 0 Then
    
        '- create the xmldom document
        Set m_objXMLDoc = New DOMDocument
        m_objXMLDoc.loadXML strFileName

        'load the xmldocument and create headerinformation
        CreateHeader "PinJig", "CINgrPinJig3"
        
        CreateTableFromPinJig pElements.Item(1)
        
        'call methods to read from objects and add to xml
        'Get Each pin object to write XML for reproting
'        GetAllPinAndCreateXMLDocument oTempElements
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Helpers for Filtering

Private Sub InitializeFilter(filterElement As IXMLDOMElement)
    On Error Resume Next

    ' Yes - Hard Code the filter Info Here
    
    Dim strFilter As String
    strFilter = "<report reportName='PinJig' >" & _
                        "<elementList elementListName='1stTable' >" & _
                            "<element elementName='***' >" & _
                                "<propertyList propertyListName='Seam' report='YES' >" & _
                                    "<property propertyName='Type' report='YES' />" & _
                                "</propertyList>" & _
                                "<propertyList propertyListName='Seam2' report='YES' >" & _
                                    "<property propertyName='MaterialType' report='YES' />" & _
                                    "<property propertyName='MaterialGrade' report='YES' />" & _
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

Private Function CHECK_REPORTABILITY(pattern As String) As Boolean
    ' Default True
    CHECK_REPORTABILITY = True
    
    If m_objFilter Is Nothing Then GoTo Cleanup
    
    On Error GoTo Cleanup
    
    Dim oInNode As IXMLDOMNode
    Set oInNode = m_objFilter.selectSingleNode("//" & pattern)
    
    ' The node coming in is not recognised
    If oInNode Is Nothing Then GoTo Cleanup
    
    ' Get the report attribute which is a sibling of the incoming node
    ' and check for its value. If not present: return true as usual (Error Case)
    ' If we are able to clearly identify that NO is marked then return False
    If oInNode.parentNode.Attributes.getNamedItem("report").nodeValue = "NO" Then CHECK_REPORTABILITY = False
    
Cleanup:
    ' Worst case the Report must be printed out : So return True
End Function

Private Function CreateTableFromPinJig(oPinJig As IJPinJig) As IJElements
    Const METHOD = "CreateTableFromPinJig"
    On Error GoTo ErrorHandler
    
    Dim oMfgChild           As IJMfgChild
    Dim oParent             As Object
    Dim oPlatePart          As IJPlatePart
    Dim oAssembly           As IJAssembly
    Dim oProjectedData      As IJJigProjectedData
    'Dim FileName            As String
    Dim oSurface            As GSCADMfgUtilSurface.IJSurfaceBody
   
    'Set PinJing as MfgChild
    Set oMfgChild = oPinJig
    Set oParent = oMfgChild.getParent()
    
    'Get projected data
    Set oProjectedData = GetProjectedPartFromPinJig(oPinJig)
    
    'Get vertical and horizontal remarking names.
    oProjectedData.GetTableNames sHorNames(), sVertNames()
        
    If TypeOf oParent Is IJPlanningAssembly Then
    
        Dim oMfgSurfaceUtil As GSCADMfgUtilSurface.MfgUtilSurface
        Set oMfgSurfaceUtil = New GSCADMfgUtilSurface.MfgUtilSurface
        
        Set oAssembly = oParent
        oMfgSurfaceUtil.GetCompositePanelSurfaceFromAssembly oAssembly, oSurface
        
        Set oMfgSurfaceUtil = Nothing
    
    ElseIf TypeOf oParent Is IJPlatePart Then
        
        Dim oGeomHlper As MfgGeomHelper
        Set oGeomHlper = New MfgGeomHelper
        
        Set oPlatePart = oParent
        Set oSurface = oGeomHlper.GetSurfaceFromPlatePart(oPlatePart, True)
        
        Set oGeomHlper = Nothing
    End If
    
    'Store all infos related to intersect points into global array. -> mRemarkingData()
    GetAllInfosRelatedToIntersectPnts oSurface, oPinJig
    
    Dim iIdx    As Integer
    
    Select Case TypeName(oParent)
        Case "IJPlanningAssembly"
        
            MakeXMLForDeserializer2 oSurface, oPinJig, m_oAngleNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_oDiagonalNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o1stRootNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o2ndRootNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o3rdRootNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o4thRootNode
            
        Case "IJPlatePart"
        
            MakeXMLForDeserializer2 oSurface, oPinJig, m_oAngleNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_oDiagonalNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o1stRootNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o2ndRootNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o3rdRootNode
            MakeXMLForDeserializer2 oSurface, oPinJig, m_o4thRootNode
    End Select
    
    Set oSurface = Nothing
    'FileName = Environ("TEMP")
    'If FileName = "" Or FileName = vbNullString Then
    '    FileName = "c:\temp" 'Only use C:\Temp if there is a %TEMP% failure
    'End If
    'FileName = FileName & "\aa.xml"
    'm_objXMLDoc.Save FileName
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Sub MakeXMLForDeserilizer(oJigSurface As Object, oXFrameColl As Collection, oSeam As IJSeam, oParent As IXMLDOMNode)
    On Error GoTo ErrorHandler
    Const METHOD = "MakeXMLForDeserilizer"
    
    Dim oGeomHelper     As New MfgGeomHelper
    Dim oDetSeam        As New StructDetailObjects.SEAM
    
    Dim oEdgeLists      As IEnumUnknown
    Dim oAftButt        As IUnknown
    Dim oForeButt       As IUnknown
    Dim oUpperSeam      As IUnknown
    Dim oLowerSeam      As IUnknown
    
    Dim oElement        As IXMLDOMNode
    Dim oPropList       As IXMLDOMNode
    Dim oProp           As IXMLDOMNode
    Dim i               As Integer
    Dim j               As Integer
    Dim oName           As IJNamedItem
    
    'Temporary --------------------------------------
    Dim Value           As Integer
    
    Dim ValueX          As Integer
    Dim ValueY          As Integer
    Dim ValueZ          As Integer
    Dim Curve           As Integer
    
    Dim AValue          As Integer
    Dim FValue          As Integer
    
    Value = 0
    
    ValueX = 1
    ValueY = 2
    ValueZ = 3
    Curve = 4
    
    AValue = 10
    FValue = 20
    
    '------------------------------------------------
    
    Set oElement = CreateChildNode(oParent, "element")
    
    '--- 1. Insert Row field names
    Set oDetSeam.object = oSeam
    AddAttribute oElement, "elementName", oDetSeam.name
    
    'Branch here
    Dim oParentElem As IXMLDOMElement
    Set oParentElem = oParent
    
    Select Case oParentElem.getAttributeNode("elementListName").nodeValue
        Case "1stTABLE"
            '--- 2. Insert row values(Frame lines) ---------
            'Create 'PropertyList element
            
            For i = 1 To oXFrameColl.Count
                Set oName = oXFrameColl.Item(i)
                
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", oName.name
            
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "X"
                
                'Using method of pinjig entity
                AddAttribute oProp, "propertyValue", Str(ValueX)
                ValueX = ValueX + 1
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "Y"
                
                'Using method of pinjig entity
                AddAttribute oProp, "propertyValue", Str(ValueY)
                ValueY = ValueY + 1
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "Z"
                
                'Using method of pinjig entity
                AddAttribute oProp, "propertyValue", Str(ValueZ)
                ValueZ = ValueZ + 1
                
                
                'If AFT butt or FORE butt
                If i = 1 Or i = oXFrameColl.Count Then
                    Set oProp = CreateChildNode(oPropList, "property")
                    AddAttribute oProp, "propertyName", "CURVE"
                    
                    'Using method of pinjig entity
                    AddAttribute oProp, "propertyValue", Str(Curve)
                    Curve = Curve + 1
                End If
            Next
        
        Case "2ndTABLE"
            
            '--- 2. Insert row values(Frame lines) ---------
            'Create 'PropertyList element
            Set oPropList = CreateChildNode(oElement, "propertyList")
            AddAttribute oPropList, "propertyListName", "Frame"
            
            For i = 1 To oXFrameColl.Count
                Set oName = oXFrameColl.Item(i)
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", oName.name
                
                '-----------------------------------------------------------------------------------------
                'Will add value using GetGirthLengthFromBasicMiddleFrameLine::IJJigIntersectPoint method
                '
                AddAttribute oProp, "propertyValue", Str(Value)
                Set oProp = Nothing
                Set oName = Nothing
                Value = Value + 1
                
                '-----------------------------------------------------------------------------------------
                
            Next
        
        Case "3rdTABLE"
        
            '--- 2. Insert row values(Frame lines) ---------
            'Create 'PropertyList element
            
            For i = 2 To oXFrameColl.Count - 1
                Set oName = oXFrameColl.Item(i)
                
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", oName.name
            
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "A"
                
                'Using method of pinjig entity
                AddAttribute oProp, "propertyValue", Str(AValue)
                AValue = AValue + 1
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "F"
                
                'Using method of pinjig entity
                AddAttribute oProp, "propertyValue", Str(FValue)
                FValue = FValue + 1
                
            Next
        
        Case "4thTABLE"
            '--- 2. Insert row values(Frame lines) ---------
            'Create 'PropertyList element
            Set oPropList = CreateChildNode(oElement, "propertyList")
            AddAttribute oPropList, "propertyListName", "Frame"
            
            For i = 1 To oXFrameColl.Count
                Set oName = oXFrameColl.Item(i)
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", oName.name
                
                '-----------------------------------------------------------------------------------------
                'Will add value using GetGirthLengthFromBasicMiddleFrameLine::IJJigIntersectPoint method
                '
                AddAttribute oProp, "propertyValue", Str(Value)
                Set oProp = Nothing
                Set oName = Nothing
                Value = Value + 1
                
                '-----------------------------------------------------------------------------------------
                
            Next
        
        Case Else
            'MsgBox "Never occur"
    End Select
  
WrapUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo WrapUp
End Sub

Private Sub MakeXMLForDeserializer2(oSurface As GSCADMfgUtilSurface.IJSurfaceBody, oPinJig As IJPinJig, oParent As IXMLDOMNode)
    On Error GoTo ErrorHandler
    Const METHOD = "MakeXMLForDeserializer2"
    
    Dim oGeomHelper     As New MfgGeomHelper
    
    Dim oElement        As IXMLDOMNode
    Dim oPropList       As IXMLDOMNode
    Dim oProp           As IXMLDOMNode
    Dim i               As Integer
    Dim j               As Integer
    Dim oName           As IJNamedItem
    
    'Temporary --------------------------------------
    Dim Value           As Integer
    
    Dim ValueX          As Integer
    Dim ValueY          As Integer
    Dim ValueZ          As Integer
    Dim Curve           As Integer
    
    Dim AValue          As Integer
    Dim FValue          As Integer
    
    Dim TRAngle         As Double
    Dim WLAngle         As Double
    Dim TmpAngle        As Double
    
    Value = 0
    
    ValueX = 1
    ValueY = 2
    ValueZ = 3
    Curve = 4
    
    AValue = 10
    FValue = 20
    
    Dim oJigOutput          As GSCADPinJigEntity.IJJigOutput
    Dim oPart3D             As GSCADPinJigEntity.IJJigPart3D
    Dim oProcessData        As GSCADPinJigEntity.IJJigProcessData
    Dim oProjectedData      As GSCADPinJigEntity.IJJigProjectedData
    Dim oIntersectPoint     As GSCADPinJigEntity.IJJigIntersectPoint
    Dim oElemsContours      As IJElements
    Dim oElemLines          As IJElements
    Dim oIntersectPnts      As IJElements
    Dim oRemarkingLines     As IJElements
    Dim oCurve              As IJCurve
    
''    Dim sVertNames()        As String
''    Dim sHorNames()         As String
    Dim iHorCnt             As Integer
    
    Dim oParentElem         As IXMLDOMElement
    Dim sTableName          As String
    Dim strNewVal           As String
    
    Set oJigOutput = oPinJig.GetJigOutput
    Set oPart3D = oPinJig.GetJigPart3D
'    Set oProjectedData = oJigOutput.GetJigProjectedData
    oPart3D.GetDiagonalLengths DistALtoFU, DistAUtoFL
    'Set oElemsContours = oProjectedData.GetEdgesByType(STRMFG_PinJigContourLine2D)

'    'Get vertical and horizontal remarking names.
'    oProjectedData.GetTableNames sHorNames(), sVertNames()
    
    Set oParentElem = oParent
    sTableName = oParentElem.getAttributeNode("elementListName").nodeValue
    
    If sTableName = "TR-WL" Or sTableName = "Diagonal Length" Then
        
        Select Case sTableName
            Case "TR-WL"
                Set oElement = CreateChildNode(oParent, "element")
                AddAttribute oElement, "elementName", "TR"
                
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", ""
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "Angle"
                
                oJigOutput.GetAngleBetweenBasePlaneAndGlobalCoordinate WLAngle, TmpAngle, TRAngle
                AddAttribute oProp, "propertyValue", CStr(Round(TRAngle, 2))
                
                '--------------------------------------------------------------------------------------
                Set oElement = CreateChildNode(oParent, "element")
                AddAttribute oElement, "elementName", "WL"
                
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", ""
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "Angle"
                
                'oJigOutput.GetAngleBetweenBasePlaneAndGivenPlane True, WLAngle
                AddAttribute oProp, "propertyValue", CStr(Round(WLAngle, 2))
                
                
            Case "Diagonal Length"
                Set oElement = CreateChildNode(oParent, "element")
                AddAttribute oElement, "elementName", "AL-FU"
                
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", ""
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "Distance"
                AddAttribute oProp, "propertyValue", CStr(Round(DistALtoFU, 2))
                
'                Set oProp = CreateChildNode(oPropList, "property")
'                AddAttribute oProp, "propertyName", "Length"
'                m_oUnitsOfMeasure.FormatUnit UNIT_DISTANCE, DistAUtoFL, strNewVal, m_oUomFormat, m_CurrentDistanceUnits
'                AddAttribute oProp, "propertyValue", strNewVal
                
                '---------------------------------------------------
                
                Set oElement = CreateChildNode(oParent, "element")
                AddAttribute oElement, "elementName", "AU-FL"
                
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", ""
                
                Set oProp = CreateChildNode(oPropList, "property")
                AddAttribute oProp, "propertyName", "Distance"
                AddAttribute oProp, "propertyValue", CStr(Round(DistAUtoFL, 2))
                
'                Set oProp = CreateChildNode(oPropList, "property")
'                AddAttribute oProp, "propertyName", "Length"
'                m_oUnitsOfMeasure.FormatUnit UNIT_DISTANCE, GirthLenAUtoFL, strNewVal, m_oUomFormat, m_CurrentDistanceUnits
'                AddAttribute oProp, "propertyValue", strNewVal
                
                
            Case Else
                'MsgBox "Never occur"
        End Select
    
    Else
    
        'Set oPart3D = oPinJig.GetJigPart3D
        Set oProjectedData = oJigOutput.GetJigProjectedData
        
        'oPart3D.GetDiagonalLengths DistALtoFU, DistAUtoFL
    
        Set oProcessData = oPart3D.GetJigProcessData
        
        Set oElemsContours = oProcessData.GetEdgesByType(STRMFG_PinJigContourLine3D)
        Set oIntersectPnts = oProjectedData.GetEdgesByType(STRMFG_PinJigCrossPoint)
        Set oRemarkingLines = oProcessData.GetEdgesByType(STRMFG_PinJigRemarkingLine3D)
        
        'Currently not used. This will be used after code is completed by Jeong. 2002/8/12
        'This will be used in distinguish Horizontal and Vertical remarking lines.
        For i = 1 To oIntersectPnts.Count
            Set oIntersectPoint = oIntersectPnts.Item(i)
    '        oIntersectPoint.HorizontalRemarkingLine
    '        oIntersectPoint.VerticalRemarkingLine
        Next

'''''   'To test Jig method(GetTableNames), tempoarary prohibit below codes.
'''''        ReDim sVertNames(1 To oRemarkingLines.Count + 2)
'''''        ReDim sHorNames(1 To 2)
'''''
'''''        For i = 1 To oRemarkingLines.Count
'''''            Set oName = oRemarkingLines.Item(i)
'''''            sVertNames(i + 1) = oName.name
'''''        Next
'''''
'''''        For i = 1 To oElemsContours.Count
'''''            Set oName = oElemsContours.Item(i)
'''''            If oName.name = "AftButt" Then
'''''                sVertNames(1) = oName.name
'''''
'''''            ElseIf oName.name = "ForeButt" Then
'''''                sVertNames(oRemarkingLines.Count + 2) = oName.name
'''''
'''''            ElseIf oName.name = "UpperSeam" Then
'''''                sHorNames(1) = oName.name
'''''
'''''            ElseIf oName.name = "LowerSeam" Then
'''''                sHorNames(2) = oName.name
'''''
'''''            End If
'''''        Next
        
        
    '    For i = 1 To UBound(sVertNames)
    '        MsgBox sVertNames(i)
    '    Next
        
       
        'Branch here
    ''''    Dim oParentElem As IXMLDOMElement
    ''''    Set oParentElem = oParent
    
        'These are used in first table
        Dim dx                  As Double
        Dim dy                  As Double
        Dim dz                  As Double
        Dim dCurve              As Double
        
        Dim dHorGirthLength     As Double
        Dim dVertGirthLength    As Double
        
        Dim dAFAngle            As Double
        Dim dFOAngle            As Double
        
        For iHorCnt = 0 To UBound(sHorNames)
            'Create element.
            Set oElement = CreateChildNode(oParent, "element")
            AddAttribute oElement, "elementName", sHorNames(iHorCnt)
            
            
            Select Case sTableName
                Case "1stTABLE"
                    '--- 2. Insert row values(Frame lines) ---------
                    'Create 'PropertyList element
                    
    '                    Set oElement = CreateChildNode(oParent, "element")
                        'AddAttribute oElement, "elementName", sHorNames(iHorCnt)
                    
                        For i = 0 To UBound(sVertNames)
            
                            Set oPropList = CreateChildNode(oElement, "propertyList")
                            AddAttribute oPropList, "propertyListName", sVertNames(i)
                            
                            
                            'Get each necessary values
                            GetFirstTableValues sHorNames(iHorCnt), sVertNames(i), dx, dy, dz, dCurve
                            
                            Set oProp = CreateChildNode(oPropList, "property")
                            AddAttribute oProp, "propertyName", "X"
            
                            'Using method of pinjig entity
                            AddAttribute oProp, "propertyValue", CStr(Round(dx, 2))
                            
'                            MsgBox "Xvalue->" & sHorNames(iHorCnt) & ":" & sVertNames(i) & ">" & strNewVal
                            
'                            ValueX = ValueX + 1
            
                            Set oProp = CreateChildNode(oPropList, "property")
                            AddAttribute oProp, "propertyName", "Y"
            
                            'Using method of pinjig entity
                            AddAttribute oProp, "propertyValue", CStr(Round(dy, 2))
'                            MsgBox "Yvalue->" & sHorNames(iHorCnt) & ":" & sVertNames(i) & ">" & strNewVal
'                            ValueY = ValueY + 1
                          
                            Set oProp = CreateChildNode(oPropList, "property")
                            AddAttribute oProp, "propertyName", "Z"
                            
                            'Using method of pinjig entity
                            AddAttribute oProp, "propertyValue", CStr(Round(dz, 2))
'                            ValueZ = ValueZ + 1
            
                            'If AFT butt or FORE butt
                            If i = 0 Or i = UBound(sVertNames) Then
                                Set oProp = CreateChildNode(oPropList, "property")
                                AddAttribute oProp, "propertyName", "CURVE"
'                                MsgBox dCurve
                                'Using method of pinjig entity
'                                MsgBox "After"
                                AddAttribute oProp, "propertyValue", CStr(Round(dCurve, 2))

                            End If
                        Next
                   '
                Case "2ndTABLE"
        
                    '--- 2. Insert row values(Frame lines) ---------
                    'Create 'PropertyList element
                    Set oPropList = CreateChildNode(oElement, "propertyList")
                    AddAttribute oPropList, "propertyListName", "Frame"
        
                    For i = 0 To UBound(sVertNames)
                        
                        GetSecondTableValues sHorNames(iHorCnt), sVertNames(i), dHorGirthLength
                        
                        Set oProp = CreateChildNode(oPropList, "property")
                        AddAttribute oProp, "propertyName", sVertNames(i)
        
                        '-----------------------------------------------------------------------------------------
                        'Will add value using GetGirthLengthFromBasicMiddleFrameLine::IJJigIntersectPoint method
                        
                        'Using method of pinjig entity
                                     
                        AddAttribute oProp, "propertyValue", CStr(Round(dHorGirthLength, 2))
                        Set oProp = Nothing
'                        Value = Value + 1
        
                        '-----------------------------------------------------------------------------------------
        
                    Next
        '
                Case "3rdTABLE"
        
                    '--- 2. Insert row values(Frame lines) ---------
                    'AddAttribute oElement, "elementName", sHorNames(iHorCnt)
                    
                    
                    
                    For i = 0 To UBound(sVertNames)
                        
                        GetThirdTableValues sHorNames(iHorCnt), sVertNames(i), dAFAngle, dFOAngle
                        
                        Set oPropList = CreateChildNode(oElement, "propertyList")
                        AddAttribute oPropList, "propertyListName", sVertNames(i)
    
                        Set oProp = CreateChildNode(oPropList, "property")
                        AddAttribute oProp, "propertyName", "A"
        
                        'Using method of pinjig entity
                        AddAttribute oProp, "propertyValue", CStr(Round(dAFAngle, 2))
                        AValue = AValue + 1
        
                        Set oProp = CreateChildNode(oPropList, "property")
                        AddAttribute oProp, "propertyName", "F"
        
                        'Using method of pinjig entity
                        AddAttribute oProp, "propertyValue", CStr(Round(dFOAngle, 2))
                        FValue = FValue + 1
                    
                    Next
        
                    
        '
                Case "4thTABLE"
                    '--- 2. Insert row values(Frame lines) ---------
                    'Create 'PropertyList element
                    Set oPropList = CreateChildNode(oElement, "propertyList")
                    AddAttribute oPropList, "propertyListName", "Frame"
        
                    For i = 0 To UBound(sVertNames)
                    
                        GetFourthTableValues sHorNames(iHorCnt), sVertNames(i), dVertGirthLength
                    
                        Set oProp = CreateChildNode(oPropList, "property")
                        AddAttribute oProp, "propertyName", sVertNames(i)
        
                        '-----------------------------------------------------------------------------------------
                        'Will add value using GetGirthLengthFromBasicMiddleFrameLine::IJJigIntersectPoint method
                        '
                        'Using method of pinjig entity
                        AddAttribute oProp, "propertyValue", CStr(Round(dVertGirthLength, 2))
                        Set oProp = Nothing
                        Value = Value + 1
        
                        '-----------------------------------------------------------------------------------------
                    Next
        '
                Case Else
                    'MsgBox "Never occur"
            End Select
        
        Next
    End If
  
WrapUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo WrapUp
End Sub


'Private Function PlateBoundaries(oPlatePart As IJPlatePart) As IJElements
'On Error GoTo ErrorHandler
'Const METHOD = "PlateBoundaries"
'
'    On Error Resume Next
'    Set PlateBoundaries = New Collection
'
'    Dim sTypeObject As String
'    Dim sObjectType As String
'    'Dim cBoundaryData As BoundaryData
'
'    Dim oPort As IJPort
'    Dim oLandCurve As IJLandCurve
'    Dim oBoundaryObject As Object
'
'    Dim nPorts As Long
'    Dim iIndex1 As Long
'    Dim iIndex2 As Long
'    Dim oOrderedIndex() As Long
'
'    Dim oPortList As IJElements
'    Dim oBoundList As IJElements
'    Dim pBoundaries As New DynElements
'
'    Dim oDetSeam As New StructDetailObjects.SEAM
'
'    Dim oPartInfo As GSCADStructGeomUtilities.PartInfo
'    Dim pStructGeometryHelper As IMSStructGenericBusinessObjects.StructGeometryHelper
'
'    Dim oPlateSystem As IJPlateSystem
'    Dim oStructDetailHelper As StructDetailHelper
'
'    ' check if Standalone Plate Part (Get the plate part's Parent System)
'    Set oStructDetailHelper = New StructDetailHelper
'    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oPlateSystem
'    Set oStructDetailHelper = Nothing
'
'    Set oPortList = New DynElements
'    Set oBoundList = New DynElements
'    Set oPartInfo = New GSCADStructGeomUtilities.PartInfo
'    Set pStructGeometryHelper = New IMSStructGenericBusinessObjects.StructGeometryHelper
'
'    If Not oPlateSystem Is Nothing Then
'        oPartInfo.GetPlatePartPortBoundaries oPlatePart, oPortList, oBoundList
'        nPorts = oPortList.Count
'
'    Else
'        nPorts = 0
'    End If
'
'    If nPorts > 0 Then
'        GetOrderPortsList oPortList, oOrderedIndex
'''        FindOutEdgePositions oPortList, oOrderedIndex
'
'        For iIndex1 = 1 To nPorts
'            'iIndex2 will be indext of corresponding seam of iIndex1
'            iIndex2 = oOrderedIndex(iIndex1)
'
'
'
''''            Dim jj As Integer
''''            Dim oNameUtil As New SDNameRulesUtilHelper
''''            Dim oLowPos As IJDPosition
''''            Dim oHighPos As IJDPosition
''''
''''
''''            oNameUtil.GetRangeCorners oPortList.Item(iIndex2), oLowPos, oHighPos
''''            MsgBox (oLowPos.x + oHighPos.x) / 2# & "," & (oLowPos.y + oHighPos.y) / 2# & "," & (oLowPos.z + oHighPos.z) / 2#
'
''            MsgBox 1
'            Set oBoundaryObject = oBoundList.Item(iIndex2)
''            MsgBox 2
''            MsgBox TypeName(oBoundaryObject)
'''            Set oBoundaryObject = oPortList.Item(iIndex2)
'
'            If TypeOf oBoundaryObject Is IJPort Then
''                MsgBox 1
'                Set oPort = oBoundaryObject
'                Set oBoundaryObject = Nothing
'                Set oBoundaryObject = oPort.Connectable
'                Set oPort = Nothing
'
'            ElseIf TypeOf oBoundaryObject Is IJLandCurve Then
''                MsgBox 2
'                Set oLandCurve = oBoundaryObject
'                Set oBoundaryObject = Nothing
'                pStructGeometryHelper.RecursiveGetStructEntityIUnknown oLandCurve, _
'                                                                       oBoundaryObject
'                Set oLandCurve = Nothing
'            End If
'
''            MsgBox 3
'            Get_ObjectTypeData oBoundaryObject, sTypeObject, sObjectType
'
'''''            Set cBoundaryData.Boundary = oBoundaryObject
'''''            cBoundaryData.ObjectType = sTypeObject
'''''            cBoundaryData.TypeObject = sObjectType
'
'            Dim edgeType As String
''            MsgBox 4
'            'Get boundary type(Aft, Fore, Upper, Lower)
'            edgeType = FindOutEdgePositions(oPortList, iIndex1)
''            MsgBox 5
'            'Add seam object wity returned boundary type.
'            pBoundaries.Add oBoundaryObject, edgeType
'
'        Next iIndex1
'
'    Else
'        ' Get the Plate Part Boundaries from the Part's CustomGeometry
'        Set oPartInfo = New GSCADStructGeomUtilities.PartInfo
'        oPartInfo.GetPlatePartBoundaries oPlatePart, oBoundList
'        Set oPartInfo = Nothing
'
'        For Each oBoundaryObject In oBoundList
'            If TypeOf oBoundaryObject Is IJPort Then
'                Set oPort = oBoundaryObject
'                Set oBoundaryObject = Nothing
'                Set oBoundaryObject = oPort.Connectable
'                Set oPort = Nothing
'            ElseIf TypeOf oBoundaryObject Is IJLandCurve Then
'                Set oLandCurve = oBoundaryObject
'                Set oBoundaryObject = Nothing
'                pStructGeometryHelper.RecursiveGetStructEntityIUnknown oLandCurve, _
'                                                                       oBoundaryObject
'                Set oLandCurve = Nothing
'            End If
'
'            Get_ObjectTypeData oBoundaryObject, sTypeObject, sObjectType
'            Set cBoundaryData.Boundary = oBoundaryObject
'            cBoundaryData.ObjectType = sTypeObject
'            cBoundaryData.TypeObject = sObjectType
'            PlateBoundaries.Add cBoundaryData
'            Set cBoundaryData.Boundary = Nothing
'        Next oBoundaryObject
'
'    End If
'
'    Set PlateBoundaries = pBoundaries
'
'    Exit Function
'
'ErrorHandler:
'  Err.Raise Err.Number
'
'End Function

'Gather information of the intersect points.
Private Sub GetAllInfosRelatedToIntersectPnts(oSurface As GSCADMfgUtilSurface.IJSurfaceBody, oPinJig As IJPinJig)

    Const METHOD = "GetAllInfosRelatedToIntersectPnts"
    On Error GoTo ErrorHandler

    Dim oElemCrossPnts          As IJElements
    Dim oElemCornerPnts         As IJElements
    Dim oInterSectPnt           As IJJigIntersectPoint
    Dim tRemarkingData          As MarkingDataInfo
    Dim strSeamName             As String
    Dim strButtName             As String
    Dim oExtMfgPin              As IJMfgPin
    Dim oJigOut                 As IJJigOutput
    Dim oProjectedData          As IJJigProjectedData
    Dim iCnt                    As Integer
    Dim oXLine                  As Object
    Dim oYLine                  As Object
    
    Set oJigOut = oPinJig.GetJigOutput()
    Set oProjectedData = oJigOut.GetJigProjectedData()
 
    Set oElemCrossPnts = oProjectedData.GetEdgesByType(STRMFG_PinJigCrossPoint)
    Set oElemCornerPnts = oProjectedData.GetEdgesByType(STRMFG_PinJigCornerPoint)

    'Union
    oElemCrossPnts.AddElements oElemCornerPnts
    
    'Re dimension array
    ReDim mRemarkingData(1 To oElemCrossPnts.Count) As MarkingDataInfo
    Dim ddCurve As Double
 
    oJigOut.CreateCoordinateLines oXLine, oYLine
    
    SET_PROGRESS_INFO oElemCrossPnts.Count, 1.75, 1
    
    For iCnt = 1 To oElemCrossPnts.Count
        Set oInterSectPnt = oElemCrossPnts.Item(iCnt)
        
        CHECK_POINT_ELEMENT oInterSectPnt
        
        If oInterSectPnt.HorizontalRemarkingLine = "UpperSeam" Or oInterSectPnt.HorizontalRemarkingLine = "LowerSeam" Then
            strSeamName = GetTheRealSeamName(sHorNames(), oInterSectPnt.HorizontalRemarkingLine)
            tRemarkingData.RemarkingHorCurveName = strSeamName
        Else
            tRemarkingData.RemarkingHorCurveName = oInterSectPnt.HorizontalRemarkingLine
        End If

        If oInterSectPnt.VerticalRemarkingLine = "AftButt" Or oInterSectPnt.VerticalRemarkingLine = "ForeButt" Then
            strButtName = GetTheRealButtName(sVertNames(), oInterSectPnt.VerticalRemarkingLine)
            tRemarkingData.RemarkingVertCurveName = strButtName
        Else
            tRemarkingData.RemarkingVertCurveName = oInterSectPnt.VerticalRemarkingLine
        End If

        'Get coordinate values along the Jig Coordinate lines.
        oInterSectPnt.GetPositionFromJigCoordinate oXLine, oYLine, tRemarkingData.dx, tRemarkingData.dy
        tRemarkingData.Height = oInterSectPnt.Height
        
        If oInterSectPnt.VerticalRemarkingLine = "AftButt" Or oInterSectPnt.VerticalRemarkingLine = "ForeButt" Then

            ddCurve = oInterSectPnt.GetCurveLengthAtSeam()
            tRemarkingData.dCurve = ddCurve

        End If
        
        'Get Horizontal and Vertical girth length from Jig coordinate line.
        oInterSectPnt.GetGirthLengthFromJigCoordLine oXLine, oYLine, tRemarkingData.HorLength, tRemarkingData.VertLength
        
        'Get aft and forward angle from bulkhead
        oInterSectPnt.GetAngleFromBulkhead oSurface, tRemarkingData.AFAngle, tRemarkingData.FOAngle

        mRemarkingData(iCnt) = tRemarkingData
        
    Next
        
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
End Sub

'Get all positions in the port
Private Function GetTheRealSeamName(aLongies() As String, strSeamName As String) As String
    Const METHOD = "GetTheRealSeamName"
    On Error GoTo ErrorHandler
    
    Dim iCnt As Integer
    
    For iCnt = 0 To UBound(aLongies())
        If strSeamName = VBA.Left(aLongies(iCnt), 9) Then
'            MsgBox aLongies(iCnt)
            GetTheRealSeamName = aLongies(iCnt)
            Exit Function
        End If
    Next
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

'Get all positions in the port
Private Function GetTheRealButtName(aFrames() As String, strButtName As String) As String
    Const METHOD = "GetTheRealButtName"
    On Error GoTo ErrorHandler
    
    Dim iCnt As Integer
    
    For iCnt = 0 To UBound(aFrames())
        If VBA.Left(strButtName, 7) = VBA.Left(aFrames(iCnt), 7) Then
'            MsgBox aFrames(iCnt)
            GetTheRealButtName = aFrames(iCnt)
            Exit Function
        End If
    Next
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetProjectedPartFromPinJig(oPinJig As Object) As IJJigProjectedData
    Const METHOD = "GetProjectedPartFromPinJig"
    On Error GoTo ErrorHandler
    
    Dim oMfgParent          As IJMfgGeomParent
    Dim oColl               As IJDTargetObjectCol
    Dim oProjectedData      As IJJigProjectedData
    
    Dim iCnt                As Integer
    
    
    Set oMfgParent = oPinJig
    
    Set oColl = oMfgParent.GetChildren()
    
    
    'Get last two objects(ProjectedData and PinSet object)
    For iCnt = 1 To oColl.Count
        If TypeOf oColl.Item(iCnt) Is IJJigOutput Then
            'This will be JigOutPut object
            Set oMfgParent = oColl.Item(iCnt)
            Set oColl = oMfgParent.GetChildren()
        End If
    Next
    
    'Now oColl has 'Projected Data' and 'PinSet' object
    For iCnt = 1 To oColl.Count
        If TypeOf oColl.Item(iCnt) Is IJJigProjectedData Then
            Set GetProjectedPartFromPinJig = oColl.Item(iCnt)
        End If
    Next
   
    Set oMfgParent = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Sub GetFirstTableValues(sHorName As String, sVertName As String, dx As Double, dy As Double, dz As Double, dCurve As Double)
    Const METHOD = "GetFirstTableValues"
    On Error GoTo ErrorHandler
    
    Dim iIdx            As Integer
    Dim tRemarkingData  As MarkingDataInfo
    
    For iIdx = 1 To UBound(mRemarkingData())
        tRemarkingData = mRemarkingData(iIdx)
        If sHorName = tRemarkingData.RemarkingHorCurveName And sVertName = tRemarkingData.RemarkingVertCurveName Then
            dx = tRemarkingData.dx
            dy = tRemarkingData.dy
            dz = tRemarkingData.Height
            dCurve = tRemarkingData.dCurve
        End If
    Next
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number
End Sub
Private Sub GetSecondTableValues(sHorName As String, sVertName As String, dHorValue As Double)
    Const METHOD = "GetFirstTableValues"
    On Error GoTo ErrorHandler
    
    Dim iIdx            As Integer
    Dim tRemarkingData  As MarkingDataInfo
    
    For iIdx = 1 To UBound(mRemarkingData())
        tRemarkingData = mRemarkingData(iIdx)
        If sHorName = tRemarkingData.RemarkingHorCurveName And sVertName = tRemarkingData.RemarkingVertCurveName Then
            dHorValue = tRemarkingData.HorLength
        End If
    Next
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub GetThirdTableValues(sHorName As String, sVertName As String, dAFAngle As Double, dFOAngle As Double)
    Const METHOD = "GetFirstTableValues"
    On Error GoTo ErrorHandler
    
    Dim iIdx            As Integer
    Dim tRemarkingData  As MarkingDataInfo
    
    For iIdx = 1 To UBound(mRemarkingData())
        tRemarkingData = mRemarkingData(iIdx)
        If sHorName = tRemarkingData.RemarkingHorCurveName And sVertName = tRemarkingData.RemarkingVertCurveName Then
            dAFAngle = tRemarkingData.AFAngle
            dFOAngle = tRemarkingData.FOAngle
        End If
    Next
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub GetFourthTableValues(sHorName As String, sVertName As String, dVertValue As Double)
    Const METHOD = "GetFirstTableValues"
    On Error GoTo ErrorHandler
    
    Dim iIdx            As Integer
    Dim tRemarkingData  As MarkingDataInfo
    
    For iIdx = 1 To UBound(mRemarkingData())
        tRemarkingData = mRemarkingData(iIdx)
        If sHorName = tRemarkingData.RemarkingHorCurveName And sVertName = tRemarkingData.RemarkingVertCurveName Then
            dVertValue = tRemarkingData.VertLength
        End If
    Next
    
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
