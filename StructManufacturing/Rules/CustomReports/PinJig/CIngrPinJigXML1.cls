VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIngrPinJig1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : Serializer
''  File        : CIngrPinJigXML1.cls
''
''  Description : Populates the XML DOM Document with PinJig related data for Sample 1
''
''
''  Author      : Intergraph Development.
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************
Implements IJDCustomReport

Private Const MODULE = "Serializer.CIngrPinJigXML1->"

Private m_oXmlDoc               As DOMDocument

'Root node for each table
Private m_oGagdo                As IXMLDOMNode
Private m_oDiagonal             As IXMLDOMNode
Private m_oGround               As IXMLDOMNode
Private m_oLong                 As IXMLDOMNode
Private m_oMarkingDataPins      As IXMLDOMNode
Private m_oMarkingDataLong      As IXMLDOMNode
Private m_oMarkingDataFrame     As IXMLDOMNode

Private m_oFileRootNode         As IXMLDOMNode

' Variable that stores the Filtering Info.
Private m_objFilter             As IXMLDOMElement

'will hold seam names(longi marking line names)
Private m_LongiNames()          As String
'will hold frame names(transverse marking line names)
Private m_FrameNames()          As String

'will hold longitudinal remarking line entities
Private m_LongiObjects        As IJElements
'will hold transverse remarking line entities
Private m_FrameObjects        As IJElements

'Coner Extreme Pins information
Private aDX()                   As Double
Private aDY()                   As Double
Private aHeight()               As Double
Private aMfgPins()              As IJMfgPin

'Surface diagonal length
Private DistALtoFU              As Double
Private DistAUtoFL              As Double

Private m_oPlateEdges           As IJElements

'Type for marking data
Private Type MarkingData
    oMfgPin                  As IJMfgPin
    RemarkingHorCurveName    As String
    RemarkingVertCurveName   As String
    dx                       As Double
    dy                       As Double
    Height                   As Double
    bFlag                    As Boolean
    HorGirthLength           As Double
    VertGirthLength          As Double
End Type

Private RemarkingData()         As MarkingData

'Private m_oDwgProgress         As IJDDwgProgress

Private Sub Class_Initialize()
    InitializeFilter m_objFilter
    InitUOMService
End Sub

Private Sub Class_Terminate()
    Set m_oXmlDoc = Nothing
    
    Set m_oMarkingDataPins = Nothing
    Set m_oMarkingDataLong = Nothing
    Set m_oMarkingDataFrame = Nothing
    
    Set m_oFileRootNode = Nothing
        
    Dim i As Integer
    Dim nCount As Long: nCount = -1
    
    On Error Resume Next
    nCount = UBound(RemarkingData())
    
    If nCount > -1 Then
        For i = 1 To nCount
            Set RemarkingData(i).oMfgPin = Nothing
        Next
    End If
    
End Sub

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
        
    CHECK_POINT_ELEMENT "Gathering elements... "
    
    If pElements.Count > 0 Then
        
        '- create the xmldom document
        Set m_oXmlDoc = New DOMDocument
        m_oXmlDoc.loadXML strFileName
        
        CreateHeader "PinJig", "CIngrPinJig1"
        
        CreateTableFromPinJig pElements.Item(1)
        m_oXmlDoc.Save strFileName
        eCustomReportStatus = StrMfgProcessFinished
    End If
    
    m_oXmlDoc.Save strFileName
        
End Sub

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
    Set m_oFileRootNode = m_oXmlDoc.createNode(NODE_ELEMENT, "report", "")
    Set m_oFileRootNode = m_oXmlDoc.appendChild(m_oFileRootNode)
    
    'Add attributes reportName and reportUID to the rootnode
    AddAttribute m_oFileRootNode, "reportName", sReportName
    AddAttribute m_oFileRootNode, "reportUID", sReportUID
    
    
    'Create elementlist and set to m_oCurrentRootNode
    Dim oElementList As IXMLDOMNode
    
    'Make Head Table
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "PART GAGDO"
    AddAttribute oElementList, "LocReport", "In"
    Set m_oGagdo = oElementList
    
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "DIAGONAL LENGTH"
    AddAttribute oElementList, "LocReport", "In"
    Set m_oDiagonal = oElementList
    
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "GROUND INFO."
    AddAttribute oElementList, "LocReport", "In"
    Set m_oGround = oElementList
    
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "LONGITUDINAL DATA"
    AddAttribute oElementList, "LocReport", "In"
    Set m_oLong = oElementList
    
    'Make 1st element list for marking data of extreme pins.
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "MARKING DATA FOR EXTREME PINS"
    AddAttribute oElementList, "LocReport", "Out"
    Set m_oMarkingDataPins = oElementList
    
'    Make 2nd element list for marking data of longi.
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "MARKING DATA FOR LONG(AFT BUTT TO FWD BUTT)"
    AddAttribute oElementList, "LocReport", "Out"
    Set m_oMarkingDataLong = oElementList

'    Make 3rd element list for marking data of frame
    Set oElementList = CreateChildNode(m_oFileRootNode, "elementList")
    AddAttribute oElementList, "elementListName", "MARKING DATA FOR FRAME(LOWER SEAM TO UPPER SEAM)"
    AddAttribute oElementList, "LocReport", "Out"
    Set m_oMarkingDataFrame = oElementList

    Set oElementList = Nothing
    
Cleanup:
    Set oElementList = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
    
End Sub

Private Function CreateTableFromPinJig(oPinJig As IJPinJig) As IJElements
    Const METHOD = "CreateTableFromPinJig"
    On Error GoTo ErrorHandler
    
    Dim oMfgChild           As IJMfgChild
    Dim oParent             As Object
    Dim oPlatePart          As IJPlatePart
    Dim oAssembly           As IJAssembly
    Dim oPart3D             As IJJigPart3D
    Dim oProjectedData      As IJJigProjectedData
    
    Dim oVector             As IJDVector
    Dim oXFrameColl         As Collection
    Dim oFrame              As IHFrame
    Dim oFrameSystem        As IHFrameSystem
    
    'For longitudinal objects
    Dim oColLongis          As New Collection
    Dim oPartSupport        As IJPartSupport
    
    'For remarking lines
    Dim oElemsRemarkingLines As IJElements
    
    Dim iCnt                As Integer
    Dim oName               As IJNamedItem
    
    Dim oElemLines          As IJElements
    
    Dim dx As Double, dy As Double, dz As Double
    Dim dXV As Double, dYV As Double, dZV As Double
    
    'Physical connection
    Dim oPhysicalConn       As IJStructPhysicalConnection
    Dim oDetPhysicalConn    As StructDetailObjects.PhysicalConn
    Dim oPort1              As IJPort
    Dim oPort2              As IJPort
    Dim oGeomUtilities      As GSCADStructGeomUtilities.TopologyLocate
    Dim oReferencePart      As Object
    Dim dAngle1             As Double
    Dim dAngle2             As Double
    Dim oColPlatesOfAssm    As IJDTargetObjectCol
    Dim iIdx                As Integer
        
    'Set PinJing as MfgChild
    Set oMfgChild = oPinJig
    
    Set oParent = oMfgChild.GetParent()
    
    Call oPinJig.GetBasePlane(dXV, dYV, dZV, dx, dy, dz)
    
    'Get projected data
    Set oProjectedData = GetProjectedPartFromPinJig(oPinJig)
    
    If TypeOf oParent Is IJPlanningAssembly Then
        Set oAssembly = oParent

        Set oColPlatesOfAssm = oAssembly.GetChildren()
        
    ElseIf TypeOf oParent Is IJPlatePart Then
        Set oPlatePart = oParent
    
    End If
    
    Set oVector = New DVector
    oVector.Set 1, 0, 0
    oVector.Length = 1
    
    'Get horizontal and vertical names
    oProjectedData.GetVerAndHorRemark2D m_LongiObjects, m_FrameObjects
    
    ReDim m_LongiNames(1 To m_LongiObjects.Count)
    For iIdx = 1 To m_LongiObjects.Count
        Dim RemarkLineName As IJNamedItem
        Set RemarkLineName = m_LongiObjects.Item(iIdx)
        m_LongiNames(iIdx) = RemarkLineName.name
    Next
    
    ReDim m_FrameNames(1 To m_FrameObjects.Count)
    For iIdx = 1 To m_FrameObjects.Count
        Set RemarkLineName = m_FrameObjects(iIdx)
        m_FrameNames(iIdx) = RemarkLineName.name
    Next

    '--- Get longitudinal profiles ------------------------------
    Dim oConObjsCol As Collection
    Dim oTotalObjCol As New Collection
    Dim oTotalConnectionsCol As New Collection
    Dim oConnectionsCol  As Collection
    Dim oThisPartPortCol  As Collection
    Dim oOtherPartPortCol As Collection
    Dim iInner As Integer
    
    Set oPartSupport = New GSCADSDPartSupport.PartSupport
    
    If TypeOf oParent Is IJPlatePart Then
        
        Set oPartSupport.Part = oPlatePart
        oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjsCol, oConnectionsCol, _
                                                                oThisPartPortCol, oOtherPartPortCol
        For iInner = 1 To oConObjsCol.Count
            oTotalObjCol.Add oConObjsCol.Item(iInner)
            oTotalConnectionsCol.Add oConnectionsCol.Item(iInner)
        Next
        
    ElseIf TypeOf oParent Is IJPlanningAssembly Then
        
        For iIdx = 1 To oColPlatesOfAssm.Count
            If TypeOf oColPlatesOfAssm.Item(iIdx) Is IJPlatePart Then
                
                Set oPartSupport.Part = oColPlatesOfAssm.Item(iIdx)
                
                oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjsCol, _
                                            oConnectionsCol, oThisPartPortCol, oOtherPartPortCol
                For iInner = 1 To oConObjsCol.Count
                    oTotalObjCol.Add oConObjsCol.Item(iInner)
                    oTotalConnectionsCol.Add oConnectionsCol.Item(iInner)
                Next
                
            End If


        Next
        
    End If
    
''    Set oPartSupport = New GSCADSDPartSupport.PartSupport
''
''    Set oPartSupport.Part = oPlatePart
''
''    oPartSupport.GetConnectedObjects ConnectionPhysical, oConObjsCol, oConnectionsCol, oThisPartPortCol, oOtherPartPortCol
    
    'Check whether profile is longitudinal or not.
    For iCnt = 1 To oTotalObjCol.Count
        
        If TypeOf oTotalObjCol.Item(iCnt) Is IJProfilePart Then
            If CheckLongiProfile(oTotalObjCol.Item(iCnt)) = True Then

                oColLongis.Add oConObjsCol.Item(iCnt)
            
            End If

        End If
    Next
    
    'Check whether plate is longitudinal or not. If so, add it to longitudinal collection.
    For iCnt = 1 To oTotalObjCol.Count

        If TypeOf oTotalObjCol.Item(iCnt) Is IJPlatePart Then
        
            If CheckPlateType(oTotalObjCol.Item(iCnt)) = LBulkheadPlate Then 'In case of longitudinal
                oColLongis.Add oTotalObjCol.Item(iCnt)
                
                'Caculate angle between connected plate part and connecting plate part.
                
                ' commented call to "GetPhysicalConnection", since this fails to get the correct physical connection
                ' between plates if there are profiles also under the same parent plate system - Maruvada
                Set oPhysicalConn = oTotalConnectionsCol.Item(iCnt)
                
                Set oDetPhysicalConn = New StructDetailObjects.PhysicalConn
                Set oDetPhysicalConn.object = oPhysicalConn
                
                Set oPort1 = oDetPhysicalConn.Port1
                Set oPort2 = oDetPhysicalConn.Port2
                
                Set oGeomUtilities = New GSCADStructGeomUtilities.TopologyLocate
                
                Set oReferencePart = oConObjsCol.Item(iCnt)
                
                oGeomUtilities.ComputeMountingAngles oPort1, oPort2, oReferencePart, dAngle1, dAngle2
                
'                dGirderAngle = (dAngle1 / 3.141592) * 180
               ' dGirderAngle = dAngle1
                
            ElseIf CheckPlateType(oTotalObjCol.Item(iCnt)) = TBulkheadPlate Then 'In case of trasverse
            
                'Caculate angle between connected plate part and connecting plate part.
                
                ' commented call to "GetPhysicalConnection", since this fails to get the correct physical connection
                ' between plates if there are profiles also under the same parent plate system - Maruvada
'                Set oPhysicalConn = oTotalConnectionsCol.Item(iCnt)
'                Set oDetPhysicalConn = New StructDetailObjects.PhysicalConn
'                Set oDetPhysicalConn.object = oPhysicalConn
'
'                Set oPort1 = oDetPhysicalConn.Port1
'                Set oPort2 = oDetPhysicalConn.Port2
'
'                Set oGeomUtilities = New GSCADStructGeomUtilities.TopologyLocate
'
'                Set oReferencePart = oTotalObjCol.Item(iCnt)
'
'                oGeomUtilities.ComputeMountingAngles oPort1, oPort2, oReferencePart, dAngle1, dAngle2
                
'                dFrameAngle = (dAngle1 / 3.141592) * 180
                'dFrameAngle = dAngle1
                
'                MsgBox "1:" & (dAngle1 / 3.141592) * 180
'                MsgBox "2:" & (dAngle2 / 3.141592) * 180

            ElseIf CheckPlateType(oTotalObjCol.Item(iCnt)) = DeckPlate Then
                
                'Caculate angle between connected plate part and connecting plate part.
                
                ' commented call to "GetPhysicalConnection", since this fails to get the correct physical connection
                ' between plates if there are profiles also under the same parent plate system - Maruvada
'                Set oPhysicalConn = oTotalConnectionsCol.Item(iCnt)
'                Set oDetPhysicalConn = New StructDetailObjects.PhysicalConn
'                Set oDetPhysicalConn.object = oPhysicalConn
'
'                Set oPort1 = oDetPhysicalConn.Port1
'                Set oPort2 = oDetPhysicalConn.Port2
'
'                Set oGeomUtilities = New GSCADStructGeomUtilities.TopologyLocate
'
'                Set oReferencePart = oTotalObjCol.Item(iCnt)
'
'                oGeomUtilities.ComputeMountingAngles oPort1, oPort2, oReferencePart, dAngle1, dAngle2
                
'                dStringerAngle = (dAngle1 / 3.141592) * 180
                'dStringerAngle = dAngle1
                
            End If
        End If
        
    Next
    

    '--- End of getting longitudinal profiles --------------------
'
'    '--- Get Edges of plate part -------------------------------------------------------
'    Dim oGeomHelper     As New MfgGeomHelper
'    Dim oDetPlatePart      As New StructDetailObjects.PlatePart
''    Dim oEdges          As IJElements
'
'    Set oDetPlatePart.Object = oPlatePart
'
'    If oDetPlatePart.ThicknessDirection = AboveDir Then
    
'    Set m_oPlateEdges = oGeomHelper.GetButtLinesAndSeamLines(oplatepart,
    
    '-----------------------------------------------------------------------------------
    
'''    If oColLongis.Count <> 0 Then
'''
'''        ReDim m_LongiNames(1 To oColLongis.Count)
'''
'''        'Get longi names
'''        For iCnt = 1 To oColLongis.Count
'''            Set oName = oColLongis.Item(iCnt)
'''            m_LongiNames(iCnt) = oName.name
'''        Next
'''
'''    'If longi objects don't exist...... Just showing garbage values.
'''    ElseIf oColLongis.Count = 0 Then
'''
'''        'Store each longitudinal(including seam line) line names.
'''        'Temporary... Now I cannot remarking names. I'm just waiting for somebody to finish this job
'''        ReDim m_LongiNames(0 To 0)
''''''        m_LongiNames(1) = "SEAM-140-0-0"
''''''        m_LongiNames(2) = "SEAM-7-20-400"
''''''        m_LongiNames(3) = "TRAC-8-293-0"
''''''        m_LongiNames(4) = "TRAC-9-249-2"
''''''        m_LongiNames(5) = "TRAC-10-212-6"
'''
'''
'''    End If
    
''    Set oProjectedData = GetProjectedPartFromPinJig(oPinJig)
    'GetRemarkingLines oProjectedData, oElemsRemarkingLines
    'Get extreme pin information.
    GetInfoOfConerPntsPin oProjectedData, aDX, aDY, aHeight
    'Get Diagonal length.
    Set oPart3D = GetPart3DFromPinJig(oPinJig)
    
    oPart3D.GetDiagonalLengths DistALtoFU, DistAUtoFL
    
    'Get marking data information.
    GetRemarkingNameWithExtremePins oProjectedData

    'Make XML data for being used in table.
    MakeAngleTable oPinJig, m_oGagdo

    MakeDiagonalLengthTable m_oDiagonal
    
    MakeGroundInfoTable m_oGround

    MakeLongDataTable m_oLong, oPinJig, oColLongis, oPlatePart
    
    MakeXMLForTableMarkingPins oPinJig, m_oMarkingDataPins
    
    MakeXMLForTableMarkingLongFrame m_oMarkingDataLong, m_FrameNames(), m_LongiNames()
    
    MakeXMLForTableMarkingLongFrame m_oMarkingDataFrame, m_FrameNames(), m_LongiNames()
    
Cleanup:
     
     
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
End Function

Private Sub MakeAngleTable(oPinJig As IJPinJig, rootNode As IXMLDOMNode)
    Const METHOD = "MakeAngleTable"
    On Error GoTo ErrorHandler
    
    Dim oElement            As IXMLDOMNode
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    
    Dim dFrameAngle        As Double
    Dim dStringerAngle    As Double
    Dim dGirderAngle       As Double
    'Dim oFormat             As New UomFormat
    Dim strNewVal           As String
    
    Dim oJigOutput As IJJigOutput
    Set oJigOutput = oPinJig.GetJigOutput
    
    oJigOutput.GetAngleBetweenBasePlaneAndGlobalCoordinate dStringerAngle, dGirderAngle, dFrameAngle
    
    '------------------------------------------------------------------------------------------
       
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "FRAME"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "CHIBU GAGDO"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(dFrameAngle, 2))
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "SINE GAP"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", Str(Format(Sin(CDbl(strNewVal) * 3.141592 / 180), "0.000"))
    
    
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "STRINGER"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"

    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "CHIBU GAGDO"

    AddAttribute oProperty, "propertyValue", CStr(Round(dStringerAngle, 2))
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "SINE GAP"

    AddAttribute oProperty, "propertyValue", Str(Format(Sin(CDbl(strNewVal) * 3.141592 / 180), "0.000"))
    

    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "GIRDER"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "CHIBU GAGDO"
 
    AddAttribute oProperty, "propertyValue", CStr(Round(dGirderAngle, 2))
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "SINE GAP"

    AddAttribute oProperty, "propertyValue", Str(Format(Sin(CDbl(strNewVal) * 3.141592 / 180), "0.000"))
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub MakeDiagonalLengthTable(rootNode As IXMLDOMNode)
    Const METHOD = "MakeDiagonalLengthTable"
    On Error GoTo ErrorHandler
    
    Dim oElement            As IXMLDOMNode
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    
    Dim strNewVal           As String
    
'-----------------------------------------------------------------------------------------------------------------------------
    
    
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "AL-FU"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "LENGTH"
    'Using method of pinjig entity
    
    AddAttribute oProperty, "propertyValue", CStr(Round(DistALtoFU, 2))
    
    
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "AU-FL"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "LENGTH"
    'Using method of pinjig entity
    
    AddAttribute oProperty, "propertyValue", CStr(Round(DistAUtoFL, 2))
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

'This function is completed when pin column and row is gotten.(2002/9/02)
Private Sub MakeGroundInfoTable(rootNode As IXMLDOMNode)
    Const METHOD = "MakeGroundInfoTable"
    On Error GoTo ErrorHandler
    
    Dim oElement            As IXMLDOMNode
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    
    Dim strNewVal           As String

    
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "1"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "POINT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(aMfgPins(1).row, aMfgPins(1).column)

    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DX"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDX(1), 2))
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DY"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDY(1), 2))

    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "HEIGHT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aHeight(1), 2))
    
    
    
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "2"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "POINT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(aMfgPins(2).row, aMfgPins(2).column)

    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DX"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDX(2), 2))
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DY"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDY(2), 2))

    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "HEIGHT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aHeight(2), 2))
    
    
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "3"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "POINT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(aMfgPins(3).row, aMfgPins(3).column)
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DX"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDX(3), 2))
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DY"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDY(3), 2))

    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "HEIGHT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aHeight(3), 2))
    
    
    Set oElement = CreateChildNode(rootNode, "element")
    AddAttribute oElement, "elementName", "4"
    
    Set oPropList = CreateChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Null"
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "POINT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(aMfgPins(4).row, aMfgPins(4).column)
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DX"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDX(4), 2))
    
    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DY"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(Round(aDY(4), 2))

    Set oProperty = CreateChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "HEIGHT"
    'Using method of pinjig entity
    AddAttribute oProperty, "propertyValue", CStr(aHeight(aDX(4), 2))
    
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub MakeLongDataTable(rootNode As IXMLDOMNode, oPinJig As IJPinJig, oCollConObjs As Collection, oPlatePart As IJPlatePart)
    Const METHOD = "MakeLongDataTable"
    On Error GoTo ErrorHandler

    Dim oElement            As IXMLDOMNode
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    Dim oCurName            As IJNamedItem
    Dim oNextName           As IJNamedItem
    
    Dim oProfileUtil        As ProfileUtils
    Dim oProfileAttribute   As IJProfileAttributes
    Dim oProfilePart        As New StructDetailObjects.ProfilePart
    
    Dim iCnt                As Integer
    
    'Physical connection
    Dim oPhysicalConn       As IJStructPhysicalConnection
    Dim oDetPhysicalConn    As StructDetailObjects.PhysicalConn
    Dim oPort1              As IJPort
    Dim oPort2              As IJPort
    Dim oReferencePart      As Object
    Dim dAngle1             As Double
    Dim dAngle2             As Double
    Dim oInterSectPnt       As IJDPosition
    
    Dim strNewVal           As String
    Dim oCurve              As IJLandCurve
    Dim oWireBody           As IJWireBody
    Dim oSGOWireBodyUtil    As SGOWireBodyUtilities
    
    Dim sx                  As Double
    Dim sy                  As Double
    Dim sz                  As Double
    
    Dim ex                  As Double
    Dim ey                  As Double
    Dim ez                  As Double
    
    Dim dAlpha As Double, dBeta As Double, dGamma As Double, dTwistRate As Double
    Dim colRegions          As IJElements
    Dim oSRegion            As IJDProfileRegion
    Dim oERegion            As IJDProfileRegion
    Dim oBound1             As Object
    Dim oBound2             As Object
    Dim oBoundList          As IJElements
    Dim i                   As Integer
    Dim oAngleUnit          As Units
    Dim Value               As Double
    Dim ABD1                As Double
    Dim ABD2                As Double
    Dim FBD1                As Double
    Dim FBD2                As Double
    
    Dim ABWebSize           As Double
    Dim FBWebSize           As Double

    Dim oProfileOrientation As IMSProfileEntity.IJDProfileOrientation

'-----------------------------------------------------------------------------------------------------------------------------
        
    
    Set oProfileUtil = New GSCADCreateModifyUtilities.ProfileUtils
    Set oProfileAttribute = oProfileUtil
    Set oSGOWireBodyUtil = New SGOWireBodyUtilities
    
    '########## For AFT Butt line.... #################################
    For iCnt = 1 To oCollConObjs.Count
        If TypeOf oCollConObjs.Item(iCnt) Is IJStiffenerPart Then
            Set oElement = CreateChildNode(rootNode, "element")
            AddAttribute oElement, "elementName", GetTheRealButtName(m_FrameNames(), "AftButt")
            
            Set oPropList = CreateChildNode(oElement, "propertyList")
            AddAttribute oPropList, "propertyListName", "Null"
            
            'Current profile
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "SLong(G)"
            
            'Get longi profile name. Current profile
            Set oCurName = oCollConObjs.Item(iCnt)
            'Set current profile to detailed profile part.
            Set oProfilePart.object = oCollConObjs.Item(iCnt)
            Set oProfileOrientation = oProfilePart.ParentSystem
            
            Set colRegions = oProfileOrientation.RegionCollection
            Set oSRegion = colRegions.Item(1)
                       
            oSRegion.GetPositionAnglesExtremity ProfileRegionBeginning, sx, sy, sz, dAlpha, dBeta, dGamma
            Value = dBeta
            AddAttribute oProperty, "propertyValue", oCurName.name + "(" + Str(Format(Value, "0.00")) + ")"
            
            'for the Next profile
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "NLong(G)"
            
            If iCnt <> oCollConObjs.Count Then
                Set oNextName = oCollConObjs.Item(iCnt + 1)
                
                Set oProfilePart.object = oCollConObjs.Item(iCnt + 1)
                Set oProfileOrientation = oProfilePart.ParentSystem
                
                Set colRegions = oProfileOrientation.RegionCollection
                Set oSRegion = colRegions.Item(1)
                
                oSRegion.GetPositionAnglesExtremity ProfileRegionBeginning, sx, sy, sz, dAlpha, dBeta, dGamma
                Value = dBeta
                AddAttribute oProperty, "propertyValue", oNextName.name + "(" + Str(Format(Value, "0.00")) + ")"
                
                'Get the profile diagonal length.
                oPinJig.GetProfileDiagonal oCollConObjs.Item(iCnt), oCollConObjs.Item(iCnt + 1), ABD1, ABD2, FBD1, FBD2
                
            End If
            
            '------------------------------------------------------------------------------------------------------
             
''            'Get the profile diagonal length.
''            oPinJig.GetProfileDiagonal oCollConObjs.Item(iCnt), oCollConObjs.Item(iCnt + 1), ABD1, ABD2, FBD1, FBD2
            
            'D1
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "D1"
            'Get the diagonal length for profiels.
            AddAttribute oProperty, "propertyValue", CStr(Round(ABD1, 2))
            
            'D2
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "D2"
            AddAttribute oProperty, "propertyValue", CStr(Round(ABD2, 2))
            
            'Web SIZE of Start stiffener
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "WebS"
            
            'Actually, this value should be changed with web size of start stiffener at aft butt.
'            AddAttribute oProperty, "propertyValue", m_oUnitsOfMeasure.ConvertDbuToUnit(UNIT_DISTANCE, oProfilePart.WebLength, m_CurrentDistanceUnits)
            oPinJig.GetProfileIntersectSize oCollConObjs.Item(iCnt), ABWebSize, FBWebSize
            AddAttribute oProperty, "propertyValue", CStr(Round(ABWebSize, 2))
            
            'Web SIZE of Next stiffener
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "WebN"
            AddAttribute oProperty, "propertyValue", CStr(Round(oProfilePart.WebLength, 2))
        End If
    Next
    
    
    '##### For forward butt line ####################################
    For iCnt = 1 To oCollConObjs.Count
        If TypeOf oCollConObjs.Item(iCnt) Is IJStiffenerPart Then
            Set oElement = CreateChildNode(rootNode, "element")
            AddAttribute oElement, "elementName", GetTheRealButtName(m_FrameNames(), "ForeButt")
            
            Set oPropList = CreateChildNode(oElement, "propertyList")
            AddAttribute oPropList, "propertyListName", "Null"
            
            'Current profile
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "SLong(G)"
            
            'Get longi profile name
            Set oCurName = oCollConObjs.Item(iCnt)
            Set oProfilePart.object = oCollConObjs.Item(iCnt)
            Set oProfileOrientation = oProfilePart.ParentSystem
            
            'Get regions.
            Set colRegions = oProfileOrientation.RegionCollection
            If colRegions.Count = 1 Then
                Set oERegion = colRegions.Item(1)
            ElseIf colRegions.Count >= 2 Then
                Set oERegion = colRegions.Item(colRegions.Count)
            End If
            
            oERegion.GetPositionAnglesExtremity ProfileRegionEnd, ex, ey, ez, dAlpha, dBeta, dGamma
            Value = dBeta
            AddAttribute oProperty, "propertyValue", oCurName.name + "(" + Str(Format(Value, "0.00")) + ")"
            
            'Next profile
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "NLong(G)"
            If iCnt <> oCollConObjs.Count Then
                Set oNextName = oCollConObjs.Item(iCnt + 1)
                Set oProfilePart.object = oCollConObjs.Item(iCnt)
                Set oProfileOrientation = oProfilePart.ParentSystem
                
                'Get regions.
                Set colRegions = oProfileOrientation.RegionCollection
                If colRegions.Count = 1 Then
                    Set oERegion = colRegions.Item(1)
                ElseIf colRegions.Count >= 2 Then
                    Set oERegion = colRegions.Item(colRegions.Count)
                End If
                
                oERegion.GetPositionAnglesExtremity ProfileRegionEnd, ex, ey, ez, dAlpha, dBeta, dGamma
                Value = dBeta
                AddAttribute oProperty, "propertyValue", oNextName.name + "(" + Str(Format(Value, "0.00")) + ")"
                
                'Get the profile diagonal length.
                oPinJig.GetProfileDiagonal oCollConObjs.Item(iCnt), oCollConObjs.Item(iCnt + 1), ABD1, ABD2, FBD1, FBD2
                
            End If
            
'            oPinJig.GetProfileDiagonal oCollConObjs.Item(iCnt), oCollConObjs.Item(iCnt + 1), ABD1, ABD2, FBD1, FBD2
            
            'D1
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "D1"
            AddAttribute oProperty, "propertyValue", CStr(Round(FBD1, 2))
            
            'D2
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "D2"
            AddAttribute oProperty, "propertyValue", CStr(Round(FBD2, 2))
            
            'Web SIZE of Start stiffener
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "WebS"
            oPinJig.GetProfileIntersectSize oCollConObjs.Item(iCnt), ABWebSize, FBWebSize
            AddAttribute oProperty, "propertyValue", CStr(Round(FBWebSize, 2))
            
            'Web SIZE of Next stiffener
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "WebN"
            AddAttribute oProperty, "propertyValue", CStr(Round(oProfilePart.WebLength, 2))
        End If
            
    Next
        
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub MakeXMLForTableMarkingPins(oPinJig As IJPinJig, rootNode As IXMLDOMNode)
    Const METHOD = "MakeXMLForTableMarkingPins"
    On Error GoTo ErrorHandler
    
    Dim oElement            As IXMLDOMNode
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    
    Dim iCnt                As Integer
    Dim CurveName           As String
    Dim PointName           As String
    Dim aDX(1 To 2)         As Double
    Dim aDY(1 To 2)         As Double
    Dim aHeight(1 To 2)     As Double
    Dim oExtMfgPins         As Collection
    Dim oProjectedData      As IJJigProjectedData
    
    Dim strNewVal           As String
    
'-----------------------------------------------------------------------------------------------------------------------------
        
    'Get the projected data from PinJig Object.
    Set oProjectedData = GetProjectedPartFromPinJig(oPinJig)
    
    SET_PROGRESS_INFO (m_FrameObjects.Count + m_LongiObjects.Count), 4, 1
    
    If m_LongiObjects.Count <> 0 Then
    
        '1. Insert row field names
        For iCnt = 1 To m_LongiObjects.Count
            
            Dim PinCnt As Integer
           
           CHECK_POINT_ELEMENT m_LongiNames(iCnt)
           
'            Set oExtMfgPins = New DynElements
            Set oExtMfgPins = GetMarkingValues(oProjectedData, m_LongiObjects(iCnt), aDX(), aDY(), aHeight())
            
            PinCnt = oExtMfgPins.Count

            'For the first point
            If (PinCnt > 0) Then
            
                Set oElement = CreateChildNode(rootNode, "element")
                AddAttribute oElement, "elementName", m_LongiNames(iCnt)
       
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", "Null"
                
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "POINT"
                
                'Using method of pinjig entity
                AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(oExtMfgPins.Item(1).row, oExtMfgPins.Item(1).column)
    
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "DX"
                'Using method of pinjig entity
                AddAttribute oProperty, "propertyValue", CStr(Round(aDX(1), 2))
                
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "DY"
                'Using method of pinjig entity
                AddAttribute oProperty, "propertyValue", CStr(Round(aDY(1), 2))
        
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "HEIGHT"
                'Using method of pinjig entity
                AddAttribute oProperty, "propertyValue", CStr(Round(aHeight(1), 2))
                
            End If
            
            'For the second point
            If (PinCnt > 1) Then
                
                Set oElement = CreateChildNode(rootNode, "element")
                AddAttribute oElement, "elementName", "" 'Longies(iCnt)
                
                Set oPropList = CreateChildNode(oElement, "propertyList")
                AddAttribute oPropList, "propertyListName", "Null"
                
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "POINT"
                'Using method of pinjig entity
                AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(oExtMfgPins.Item(2).row, oExtMfgPins.Item(2).column)
        
                
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "DX"
                'Using method of pinjig entity
                AddAttribute oProperty, "propertyValue", CStr(Round(aDX(2), 2))
                
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "DY"
                'Using method of pinjig entity
                
                AddAttribute oProperty, "propertyValue", CStr(Round(aDY(2), 2))
        
                Set oProperty = CreateChildNode(oPropList, "property")
                AddAttribute oProperty, "propertyName", "HEIGHT"
                'Using method of pinjig entity
                AddAttribute oProperty, "propertyValue", CStr(Round(aHeight(2), 2))
                
            End If

        Next
        
    End If
    
    If m_FrameObjects.Count <> 0 Then
    
        For iCnt = 1 To m_FrameObjects.Count

            CHECK_POINT_ELEMENT m_FrameNames(iCnt)
            
            Set oExtMfgPins = GetMarkingValues(oProjectedData, m_FrameObjects(iCnt), aDX(), aDY(), aHeight())
            
            Set oElement = CreateChildNode(rootNode, "element")
            AddAttribute oElement, "elementName", m_FrameNames(iCnt)
            
             Set oPropList = CreateChildNode(oElement, "propertyList")
            AddAttribute oPropList, "propertyListName", "Null"
            
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "POINT"
            AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(oExtMfgPins.Item(1).row, oExtMfgPins.Item(1).column)
            
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "DX"
            AddAttribute oProperty, "propertyValue", CStr(Round(aDX(1), 2))
            
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "DY"
            AddAttribute oProperty, "propertyValue", CStr(Round(aDY(1), 2))
            
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "HEIGHT"
            AddAttribute oProperty, "propertyValue", CStr(Round(aHeight(1), 2))
            
        '---------------------------------------------------------------------------------------------------------------------
        
            Set oElement = CreateChildNode(rootNode, "element")
            AddAttribute oElement, "elementName", "" 'Frames(iCnt)
 
             Set oPropList = CreateChildNode(oElement, "propertyList")
            AddAttribute oPropList, "propertyListName", "Null"

            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "POINT"

            AddAttribute oProperty, "propertyValue", MakePinNameWithRowColumn(oExtMfgPins.Item(2).row, oExtMfgPins.Item(2).column)

            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "DX"

            AddAttribute oProperty, "propertyValue", CStr(Round(aDX(2), 2))
            
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "DY"
            AddAttribute oProperty, "propertyValue", CStr(Round(aDY(2), 2))
            
            Set oProperty = CreateChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", "HEIGHT"
            AddAttribute oProperty, "propertyValue", CStr(Round(aHeight(2), 2))
                      
        Next
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

Private Sub MakeXMLForTableMarkingLongFrame(rootNode As IXMLDOMNode, Frames() As String, Longies() As String)
    Const METHOD = "MakeXMLForTableMarkingLongFrame"
    On Error GoTo ErrorHandler
    
    Dim oElement            As IXMLDOMNode
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    
    Dim iCnt                As Integer
    Dim jCnt                As Integer
    
    Dim HorValue            As Double
    Dim VertValue           As Double
    
    Dim strNewVal           As String
    Dim topNode             As String
    
'-----------------------------------------------------------------------------------------------------------------------------
    
    
    If UBound(Longies()) <> 0 Then
    
        '1. Insert row field names
        For iCnt = 1 To UBound(Longies())
            Set oElement = CreateChildNode(rootNode, "element")
            AddAttribute oElement, "elementName", Longies(iCnt)
            
            Set oPropList = CreateChildNode(oElement, "propertyList")
            AddAttribute oPropList, "propertyListName", "Null"
            
            If UBound(Frames()) <> 0 Then
            
                For jCnt = 1 To UBound(Frames())
                    Set oProperty = CreateChildNode(oPropList, "property")
                    AddAttribute oProperty, "propertyName", Frames(jCnt)

                    topNode = rootNode.Attributes.getNamedItem("elementListName").nodeValue

                    Select Case topNode
                        Case "MARKING DATA FOR LONG(AFT BUTT TO FWD BUTT)"

                            HorValue = GetHorizontalValues(Longies(iCnt), Frames(jCnt))
                        
                            'm_oUnitsOfMeasure.FormatUnit UNIT_DISTANCE, HorValue, strNewVal, oFormat, m_CurrentDistanceUnits
                            strNewVal = CStr(Round(HorValue, 2))
                        Case "MARKING DATA FOR FRAME(LOWER SEAM TO UPPER SEAM)"

                            VertValue = GetVerticalValues(Longies(iCnt), Frames(jCnt))
                            strNewVal = CStr(Round(VertValue, 2))
                           ' m_oUnitsOfMeasure.FormatUnit UNIT_DISTANCE, VertValue, strNewVal, oFormat, m_CurrentDistanceUnits
                    End Select
                    
                    AddAttribute oProperty, "propertyValue", strNewVal
                Next
            
            End If
            
        Next
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub

'''Private Sub MakeXMLForTableMarkingFrame(rootNode As IXMLDOMNode, Frames() As String, Longies() As String)
'''    Const METHOD = "CreateTableFromPinJig"
'''    On Error GoTo ErrorHandler
'''
'''    Dim oElement            As IXMLDOMNode
'''    Dim oPropList           As IXMLDOMNode
'''    Dim oProperty           As IXMLDOMNode
'''
'''    Dim iCnt                As Integer
'''
'''
'''
'''    '1. Insert row field names
'''    For iCnt = 1 To UBound(Longies())
'''        Set oElement = CreateChildNode(rootNode, "element")
'''        AddAttribute oElement, "elementName", Longies(iCnt)
'''
'''        Set oPropList = CreateChildNode(oElement, "propertyList")
'''        AddAttribute oPropList, "propertyListName", "Null"
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "POINT"
'''        'Using method of pinjig entity
'''        AddAttribute oProperty, "propertyValue", "B4/M4"
'''
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "DX"
'''        'Using method of pinjig entity
'''        AddAttribute oProperty, "propertyValue", "123"
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "DY"
'''        'Using method of pinjig entity
'''        AddAttribute oProperty, "propertyValue", "456"
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "HEIGHT"
'''        'Using method of pinjig entity
'''        AddAttribute oProperty, "propertyValue", "234"
'''
'''    Next
'''
'''    For iCnt = 1 To UBound(Frames())
'''        Set oElement = CreateChildNode(rootNode, "element")
'''        AddAttribute oElement, "elementName", Frames(iCnt)
'''
'''         Set oPropList = CreateChildNode(oElement, "propertyList")
'''        AddAttribute oPropList, "propertyListName", "Null"
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "POINT"
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "DX"
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "DY"
'''
'''        Set oProperty = CreateChildNode(oPropList, "property")
'''        AddAttribute oProperty, "propertyName", "HEIGHT"
'''    Next
'''
'''    Exit Sub
'''
'''ErrorHandler:
'''    ReportAndRaiseUnanticipatedError MODULE, METHOD
'''End Sub





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
    Set oATT = m_oXmlDoc.createAttribute(AttName)
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
    Set oChild = m_oXmlDoc.createElement(ChildName)
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

'Get corner points pin information
'aHeight() as Double and aPosExtPin() as String will be added in the parameter lists.
Private Sub GetInfoOfConerPntsPin(oProjectedPart As IJJigProjectedData, aDX() As Double, aDY() As Double, aHeight() As Double)

    Const METHOD = "GetInfoOfConerPntsPin"
    On Error GoTo ErrorHandler

    Dim oProjectedData      As IJJigProjectedData
    Dim oCornerPts          As IJElements
    Dim oIntersectionPt     As IJJigIntersectPoint
    Dim oObject             As Object
    Dim iCnt                As Integer
    Dim CntPts              As Integer
    Dim oMfgPin             As IJMfgPin
        
    Set oProjectedData = oProjectedPart

    'Now working for only 4 corners.
    Set oCornerPts = oProjectedData.GetEdgesByType(STRMFG_PinJigCornerPoint)

    CntPts = oCornerPts.Count
    
'    ReDim aPoint(CntPts)
    ReDim aDX(1 To CntPts)
    ReDim aDY(1 To CntPts)
    ReDim aHeight(1 To CntPts)
    ReDim aMfgPins(1 To CntPts)
    
    SET_PROGRESS_INFO oCornerPts.Count, 10, 1
    
    For iCnt = 1 To oCornerPts.Count
        
        Set oObject = oCornerPts.Item(iCnt)
        Set oIntersectionPt = oObject
        
        CHECK_POINT_ELEMENT oIntersectionPt
        
        oIntersectionPt.GetHoriAndVertDistWithNearestMfgPin aDX(oIntersectionPt.Order), aDY(oIntersectionPt.Order), oMfgPin
        aHeight(oIntersectionPt.Order) = oIntersectionPt.Height
        Set aMfgPins(oIntersectionPt.Order) = oMfgPin
    Next
    
    Exit Sub
    
ErrorHandler:
  
    Err.Raise Err.Number
End Sub

Private Sub GetRemarkingNameWithExtremePins(oProjectedData As IJJigProjectedData)

    Const METHOD = "GetRemarkingNameWithExtremePins"
    On Error GoTo ErrorHandler
    
    Dim oElemCrossPts       As IJElements
    Dim oElemConerPts       As IJElements
    Dim oIntersectPt        As IJJigIntersectPoint
    Dim iCnt                As Integer
    Dim eRemarkingData      As MarkingData
    Dim strSeamName         As String
    Dim strButtName         As String
    
    Dim oExtMfgPin             As IJMfgPin
    
    Set oElemCrossPts = oProjectedData.GetEdgesByType(STRMFG_PinJigCrossPoint)
    Set oElemConerPts = oProjectedData.GetEdgesByType(STRMFG_PinJigCornerPoint)
    
    oElemCrossPts.AddElements oElemConerPts
    
    ReDim RemarkingData(1 To oElemCrossPts.Count) As MarkingData
    
    SET_PROGRESS_INFO oElemCrossPts.Count, 1.75, 1
    
    For iCnt = 1 To oElemCrossPts.Count
        Set oIntersectPt = oElemCrossPts.Item(iCnt)
        
        CHECK_POINT_ELEMENT oIntersectPt
        
        If oIntersectPt.HorizontalRemarkingLine = "UpperSeam" Or oIntersectPt.HorizontalRemarkingLine = "LowerSeam" Then

            strSeamName = GetTheRealSeamName(m_LongiNames(), oIntersectPt.HorizontalRemarkingLine)
            eRemarkingData.RemarkingHorCurveName = strSeamName 'oIntersectPt.HorizontalRemarkingLine
            
        Else
            eRemarkingData.RemarkingHorCurveName = oIntersectPt.HorizontalRemarkingLine
            
            
        End If

        If oIntersectPt.VerticalRemarkingLine = "AftButt" Or oIntersectPt.VerticalRemarkingLine = "ForeButt" Then

            strButtName = GetTheRealButtName(m_FrameNames(), oIntersectPt.VerticalRemarkingLine)

            eRemarkingData.RemarkingVertCurveName = strButtName 'oIntersectPt.VerticalRemarkingLine
            
        Else
            eRemarkingData.RemarkingVertCurveName = oIntersectPt.VerticalRemarkingLine
            
        End If
        
        oIntersectPt.GetHoriAndVertDistWithNearestMfgPin eRemarkingData.dx, eRemarkingData.dy, oExtMfgPin
       
        Set eRemarkingData.oMfgPin = oExtMfgPin
        
        Dim HLenNoShrk As Double
        Dim VLenNoShrk As Double
        oIntersectPt.GetSHIGirthLength eRemarkingData.HorGirthLength, eRemarkingData.VertGirthLength, HLenNoShrk, VLenNoShrk

        eRemarkingData.Height = oIntersectPt.Height

        RemarkingData(iCnt) = eRemarkingData
    Next
    
    Exit Sub
    
ErrorHandler:
  
    Err.Raise Err.Number
End Sub

'Get each values for marking pins table
Private Function GetMarkingValues(oJigProjectedData As IJJigProjectedData, _
                            CurveObj As Object, aDX() As Double, aDY() As Double, aHeight() As Double) As Collection
    Const METHOD = "GetMarkingValues"
    On Error GoTo ErrorHandler
    
    Dim oIdx            As Integer
    Dim iIdx            As Integer
    Dim MarkingDt       As MarkingData
    Dim oInterPnts      As IJElements
    Dim oInterSectPnt   As IJJigIntersectPoint
    Dim dxExtPin        As Double
    Dim dyExtPin        As Double
    Dim oExtMfgPin      As IJMfgPin
    Dim oElemPins       As New Collection

    Set oInterPnts = oJigProjectedData.GetIntersectPointsOfRemarkingLine(CurveObj)

    For oIdx = 1 To oInterPnts.Count
        
        Set oInterSectPnt = oInterPnts.Item(oIdx)
        oInterSectPnt.GetHoriAndVertDistWithNearestMfgPin dxExtPin, dyExtPin, oExtMfgPin

        For iIdx = 1 To UBound(RemarkingData())
            MarkingDt = RemarkingData(iIdx)

            If oExtMfgPin Is MarkingDt.oMfgPin Then
                RemarkingData(iIdx).bFlag = True
                aDX(oIdx) = MarkingDt.dx
                aDY(oIdx) = MarkingDt.dy
                aHeight(oIdx) = MarkingDt.Height
                oElemPins.Add MarkingDt.oMfgPin
            End If
        Next
         
        Set oExtMfgPin = Nothing
    Next

    Set GetMarkingValues = oElemPins
    
''    oElemPins.Clear
    Set oElemPins = Nothing
    
    Exit Function
    
ErrorHandler:
    
    Err.Raise Err.Number
End Function

'Get horizontal values
Private Function GetHorizontalValues(HorName As String, VertName As String) As Double
    Const METHOD = "GetHorizontalValues"
    On Error GoTo ErrorHandler

    Dim iIdx                As Integer
    Dim MarkingDt           As MarkingData
    Dim oGlobals            As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
    Dim bAB                 As Boolean
    Dim bFB                 As Boolean
    Dim bUS                 As Boolean
    Dim bLS                 As Boolean
    
    
    For iIdx = 1 To UBound(RemarkingData())
    
        MarkingDt = RemarkingData(iIdx)

        If MarkingDt.RemarkingHorCurveName = HorName And MarkingDt.RemarkingVertCurveName = VertName Then

            GetHorizontalValues = MarkingDt.HorGirthLength
        
        End If
        
    Next
    

Exit Function
    
ErrorHandler:
    
    Err.Raise Err.Number

End Function

'Get vertical values
Private Function GetVerticalValues(HorName As String, VertName As String) As Double
    Const METHOD = "GetVerticalValues"
    On Error GoTo ErrorHandler

    Dim iIdx        As Integer
    Dim MarkingDt   As MarkingData
    
    For iIdx = 1 To UBound(RemarkingData())
        MarkingDt = RemarkingData(iIdx)
        
        If MarkingDt.RemarkingHorCurveName = HorName And MarkingDt.RemarkingVertCurveName = VertName Then
            
            GetVerticalValues = MarkingDt.VertGirthLength
        
        End If
        
    Next
    

Exit Function
    
ErrorHandler:
    
    Err.Raise Err.Number

End Function

'Get remarking lines
Private Sub GetRemarkingLines(oProjectedData As IJJigProjectedData, oElems As IJElements)
    Const METHOD = "GetRemarkingLines"
    On Error GoTo ErrorHandler

    Dim oName       As IJNamedItem
    Dim iCnt        As Integer
    
    Set oElems = oProjectedData.GetEdgesByType(STRMFG_PinJigRemarkingLine2D)
    
    Exit Sub
    
ErrorHandler:
    
    Err.Raise Err.Number
End Sub

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

Private Function GetPart3DFromPinJig(oPinJig As Object) As IJJigPart3D
    Const METHOD = "GetPart3DFromPinJig"
    On Error GoTo ErrorHandler
    
    Dim oMfgParent          As IJMfgGeomParent
    Dim oColl               As IJDTargetObjectCol
    Dim oProjectedData      As IJJigProjectedData
    
    Dim iCnt                As Integer
    
    
    Set oMfgParent = oPinJig
    
    Set oColl = oMfgParent.GetChildren()
    
    
    For iCnt = 1 To oColl.Count
        If TypeOf oColl.Item(iCnt) Is IJJigPart3D Then
            Set GetPart3DFromPinJig = oColl.Item(iCnt)
        End If
    Next
    
    Set oMfgParent = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function


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

Private Sub InitUOMService()
    Const METHOD As String = "InitUOMService"
    On Error GoTo ErrorHandler
    
'    Dim oTrader As New Trader
'
'    Set m_oUnitsOfMeasure = oTrader.Service(TKUnitsOfMeasure, "")
'
'    m_oUnitsOfMeasure.GetDefaultUnits UNIT_DISTANCE, m_CurrentDistanceUnits
'    m_oUnitsOfMeasure.GetDefaultUnits UNIT_ANGLE, m_CurrentAngleUnits
'
'    Set oTrader = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
End Sub


'Get all positions in the port
Private Function GetAllPositionsInPort(oPort As IJPort) As IJElements
    Const METHOD = "GetAllPositionsInPort"
    On Error GoTo ErrorHandler
    
    Dim oWireBody               As IJWireBody
    Dim oSGOWireBodyUtilities   As New GSCADShipGeomOps.SGOWireBodyUtilities
    Dim oCollPnts               As Collection
    
    
    Set oWireBody = oPort.Geometry
    
    Set oCollPnts = oSGOWireBodyUtilities.GetCurveDefinitionPoints(oWireBody)

    
''    Dim oPortGeometryDisp       As Object
''    Dim oPointsGraphBody        As IJPointsGraphBody
''    Dim oUnknown                As IUnknown
''    Dim oPointGraphEnum         As IEnumVARIANT
''
''    Dim nCount                  As Integer
''
''
''    Set oPortGeometryDisp = oPort.Geometry
''    MsgBox 1
''    Set oUnknown = oPortGeometryDisp
''    MsgBox 2
''    Set oPointsGraphBody = oUnknown
''    MsgBox 3
''    oPointsGraphBody.EnumPositions oPointGraphEnum, nCount
''
''    MsgBox nCount
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

'Get all positions in the port
Private Function GetTheRealSeamName(aLongies() As String, strSeamName As String) As String
    Const METHOD = "GetTheRealSeamName"
    On Error GoTo ErrorHandler
    
    Dim iCnt As Integer
    
    For iCnt = 0 To UBound(aLongies())
        If strSeamName = VBA.Left(aLongies(iCnt), 9) Then
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
            GetTheRealButtName = aFrames(iCnt)
            Exit Function
        End If
    Next
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function MakePinNameWithRowColumn(row As Integer, column As Integer) As String
    Const METHOD As String = "MakePinNameWithRowColumn"
    On Error GoTo ErrorHandler
    
    row = row + 1
    If Mid(Str(column), 1, 1) = "0" Then
        column = CInt(Mid(Str(column), 2, Len(Str(column) - 1))) + 1
    End If
    
        
    Select Case column
        Case 0
            MakePinNameWithRowColumn = "A" + Str(row)
        Case 1
            MakePinNameWithRowColumn = "B" + Str(row)
        Case 2
            MakePinNameWithRowColumn = "C" + Str(row)
        Case 3
            MakePinNameWithRowColumn = "D" + Str(row)
        Case 4
            MakePinNameWithRowColumn = "E" + Str(row)
        Case 5
            MakePinNameWithRowColumn = "F" + Str(row)
        Case 6
            MakePinNameWithRowColumn = "G" + Str(row)
        Case 7
            MakePinNameWithRowColumn = "H" + Str(row)
        Case 8
            MakePinNameWithRowColumn = "I" + Str(row)
        Case 9
            MakePinNameWithRowColumn = "J" + Str(row)
        Case 10
            MakePinNameWithRowColumn = "K" + Str(row)
        Case 11
            MakePinNameWithRowColumn = "L" + Str(row)
        Case 12
            MakePinNameWithRowColumn = "M" + Str(row)
        Case 13
            MakePinNameWithRowColumn = "N" + Str(row)
        Case 14
            MakePinNameWithRowColumn = "O" + Str(row)
        Case 15
            MakePinNameWithRowColumn = "P" + Str(row)
        Case 16
            MakePinNameWithRowColumn = "Q" + Str(row)
        Case 17
            MakePinNameWithRowColumn = "R" + Str(row)
        Case 18
            MakePinNameWithRowColumn = "S" + Str(row)
        Case 19
            MakePinNameWithRowColumn = "T" + Str(row)
        Case 20
            MakePinNameWithRowColumn = "U" + Str(row)
        Case 21
            MakePinNameWithRowColumn = "V" + Str(row)
        Case 22
            MakePinNameWithRowColumn = "W" + Str(row)
        Case 23
            MakePinNameWithRowColumn = "X" + Str(row)
        Case 24
            MakePinNameWithRowColumn = "Y" + Str(row)
        Case 25
            MakePinNameWithRowColumn = "Z" + Str(row)
        Case 26
            MakePinNameWithRowColumn = "AA" + Str(row)
        Case 27
            MakePinNameWithRowColumn = "AB" + Str(row)
        Case 28
            MakePinNameWithRowColumn = "AC" + Str(row)
        Case 29
            MakePinNameWithRowColumn = "AD" + Str(row)
        Case 30
            MakePinNameWithRowColumn = "AE" + Str(row)
        Case 31
            MakePinNameWithRowColumn = "AF" + Str(row)
        Case 32
            MakePinNameWithRowColumn = "AG" + Str(row)
        Case Else
            MakePinNameWithRowColumn = ""
    End Select
    
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

'Private Function DbgPrintObjName(obj As Object) As String
'    Dim oName As IJNamedItem
'
'    If TypeOf obj Is IJNamedItem Then
'        Set oName = obj
'
'        DbgPrintObjName = oName.name
'        Debug.Print DbgPrintObjName
'    End If
'
'End Function

