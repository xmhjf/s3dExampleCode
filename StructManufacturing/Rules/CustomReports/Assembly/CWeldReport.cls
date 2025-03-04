VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWeldReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CWeldReport.cls
''
''  Description : Populates a xml file with welding geometry and Part information
''                      - Points on Welding Curve, Normal Vector of Parts, Bevel Information..
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

Private Const Module = "AssemblyCustomReports.CWeldReport"
Private m_dTolerance As Double
Private m_dMinLength As Double
Private m_oPOM As IJDPOM
Private PI As Double

Private Sub Class_Initialize()
    Set m_oPOM = GetPOM
    m_dTolerance = 0.02
    m_dMinLength = 0.05
    PI = 3.141592
End Sub

Private Sub Class_Terminate()
        Set m_oPOM = Nothing
End Sub



Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
            
    If pElements.Count > 0 Then
       
        Dim oSelectedObj As Object
        
        Dim strExtension As String
        Dim strBaseOutputName As String

        Dim lastPos As Long
        lastPos = InStrRev(strFileName, ".")

        strBaseOutputName = Left$(strFileName, lastPos - 1)
        strExtension = Right$(strFileName, Len(strFileName) - lastPos + 1)
        
        For Each oSelectedObj In pElements
              
            If TypeOf oSelectedObj Is IJAssemblyBase Then
            
                If m_oPOM Is Nothing Then Exit Sub
                
                Dim oMfgXMLHelper As IJMfgXMLHelper
                Set oMfgXMLHelper = New MfgXMLHelper

                Dim strOutputXML As String
                oMfgXMLHelper.GetSMSWeldElement Nothing, oSelectedObj, MfgWeldDrillDownInternal, strOutputXML
    
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Build output document.
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim oOutputXML As IXMLDOMDocument
                Set oOutputXML = New DOMDocument
                Dim bSucess As Boolean
                
                bSucess = oOutputXML.loadXML(strOutputXML)

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '1. Create SMS_WELDING Node
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim oSMSWeldingNode As IXMLDOMNode
                Set oSMSWeldingNode = oOutputXML.selectSingleNode("//SMS_WELDING")
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '2. Create SMS_ASSEMBLY Node
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                GenerateAssemblyNode oOutputXML, oSelectedObj, oSMSWeldingNode
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '3. Create SMS_CONNECTION Trace Node
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim oSMSConnectionNodeList As IXMLDOMNodeList
                Dim oSMSConnectionNode As IXMLDOMNode
                Set oSMSConnectionNodeList = oOutputXML.selectNodes("//SMS_CONNECTION")
            
                For Each oSMSConnectionNode In oSMSConnectionNodeList ' For each SMS_CONNECTION
                    GenerateJointNode oOutputXML, oSMSConnectionNode
                Next oSMSConnectionNode

                Dim strGuid As String
                strGuid = m_oPOM.DbIdentifierFromMoniker(m_oPOM.GetObjectMoniker(oSelectedObj))

                Dim strOutputDocumentName As String
                strOutputDocumentName = strBaseOutputName + "_" + strGuid + strExtension
            
                If bSucess = True Then
                    oOutputXML.Save strOutputDocumentName
                End If
            End If
        Next
        
        eCustomReportStatus = StrMfgProcessFinished
    End If
            
End Sub

Private Sub GenerateAssemblyNode(ByVal oOutputXML As IXMLDOMDocument, ByVal oAssembly As Object, ByRef oParentNode As IXMLDOMNode)
Const METHOD = "GenerateAssemblyNode"
On Error GoTo ErrorHandler

    Dim oAssemblyNode As IXMLDOMNode
    Dim oAssemblyElem As IXMLDOMElement
    
    Set oAssemblyNode = oOutputXML.createNode(NODE_ELEMENT, "SMS_ASSEMBLY", "")
    Set oAssemblyElem = oAssemblyNode
    
    ' Set Assembly Attributes
    Dim oNamedItem As IJNamedItem
    Dim oAssemblyBase As IJAssemblyBase
    Dim oAssemeblyChild As IJAssemblyChild
    Dim oAssemblySequence As IJAssemblySequence
        
    Set oNamedItem = oAssembly
    Set oAssemblyBase = oAssembly
    Set oAssemeblyChild = oAssembly
    
    On Error Resume Next
    Set oAssemblySequence = oAssemeblyChild.Parent
    On Error GoTo ErrorHandler
    
    oAssemblyElem.setAttribute "NAME", oNamedItem.Name
    oAssemblyElem.setAttribute "GUID", CleanGUID(m_oPOM.DbIdentifierFromMoniker(m_oPOM.GetObjectMoniker(oAssembly)))
    oAssemblyElem.setAttribute "TYPE", oAssemblyBase.Type
    oAssemblyElem.setAttribute "STAGE", oAssemblyBase.Stage
    If oAssemblySequence Is Nothing Then
        oAssemblyElem.setAttribute "SEQUENCE", "1"
    Else
        oAssemblyElem.setAttribute "SEQUENCE", CStr(oAssemblySequence.ChildIndex(oAssembly))
    End If

    
    Set oNamedItem = Nothing
    Set oAssemblyBase = Nothing
    Set oAssemblySequence = Nothing
    
    ' Create and load the Assembly Orientation
    
    Dim oCVGMatrixNode As IXMLDOMNode
    Set oCVGMatrixNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_MATRIX", "")
    
    Dim oCVGMatrixElem As IXMLDOMElement
    Set oCVGMatrixElem = oCVGMatrixNode
    
    Dim oViewMatrix As IJDT4x4
    Dim oAssemblyOrientation As IJAssemblyOrientation
    Set oAssemblyOrientation = oAssembly
    
    Set oViewMatrix = oAssemblyOrientation.ViewMatrix
                    
    oCVGMatrixElem.setAttribute "M00", Format(oViewMatrix.IndexValue(0), "0.000")
    oCVGMatrixElem.setAttribute "M01", Format(oViewMatrix.IndexValue(1), "0.000")
    oCVGMatrixElem.setAttribute "M02", Format(oViewMatrix.IndexValue(2), "0.000")
    oCVGMatrixElem.setAttribute "M03", Format(oViewMatrix.IndexValue(3), "0.000")
    
    oCVGMatrixElem.setAttribute "M10", Format(oViewMatrix.IndexValue(4), "0.000")
    oCVGMatrixElem.setAttribute "M11", Format(oViewMatrix.IndexValue(5), "0.000")
    oCVGMatrixElem.setAttribute "M12", Format(oViewMatrix.IndexValue(6), "0.000")
    oCVGMatrixElem.setAttribute "M13", Format(oViewMatrix.IndexValue(7), "0.000")
    
    oCVGMatrixElem.setAttribute "M20", Format(oViewMatrix.IndexValue(8), "0.000")
    oCVGMatrixElem.setAttribute "M21", Format(oViewMatrix.IndexValue(9), "0.000")
    oCVGMatrixElem.setAttribute "M22", Format(oViewMatrix.IndexValue(10), "0.000")
    oCVGMatrixElem.setAttribute "M23", Format(oViewMatrix.IndexValue(11), "0.000")
    
    oCVGMatrixElem.setAttribute "M30", Format(oViewMatrix.IndexValue(12), "0.000")
    oCVGMatrixElem.setAttribute "M31", Format(oViewMatrix.IndexValue(13), "0.000")
    oCVGMatrixElem.setAttribute "M32", Format(oViewMatrix.IndexValue(14), "0.000")
    oCVGMatrixElem.setAttribute "M33", Format(oViewMatrix.IndexValue(15), "0.000")
    
    oAssemblyNode.appendChild oCVGMatrixNode
    oParentNode.appendChild oAssemblyNode
    
    Set oAssemblyOrientation = Nothing
    Set oViewMatrix = Nothing
    
CleanUp:
    Exit Sub

ErrorHandler:
    'Err.Raise Err.Number
    GoTo CleanUp
End Sub

Private Function GetShortDescription(strTableName As String, valueID As Long) As String
    On Error GoTo ErrorHandler
    Const METHOD = "GetShortDescription"
    
    Dim oIJDCodeListMetaData As IJDCodeListMetaData
    Set oIJDCodeListMetaData = m_oPOM
    
    If Not oIJDCodeListMetaData Is Nothing Then
        GetShortDescription = oIJDCodeListMetaData.ShortStringValue(strTableName, valueID)
    End If
    
    Exit Function
ErrorHandler:
    Resume Next
    
End Function


Private Sub GenerateJointNode(ByVal oOutputXML As IXMLDOMDocument, ByRef oParentNode As IXMLDOMNode)
Const METHOD = "GenerateJointNode"
On Error GoTo ErrorHandler

    Dim oPlate As IJPlate
    Dim oStiffener As IJStiffener
    Dim oBean As IJBeam
    Dim oNamedItem As IJNamedItem
    Dim oStructMaterial As IJStructureMaterial
    
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
    
    Dim strPhysicalConnGuid As String
    Dim strRefPartGuid As String
    Dim strNonRefPartGuid As String
    Dim FirstGeometry As Boolean
    
    strPhysicalConnGuid = oParentNode.selectSingleNode("./@CONN_GUID").nodeValue
    strRefPartGuid = oParentNode.selectSingleNode("./@REFERENCE_PART").nodeValue
    strNonRefPartGuid = oParentNode.selectSingleNode("./@NON_REFERENCE_PART").nodeValue
        
    Dim oPhysicalConn As IJStructPhysicalConnection
    Dim oRefPart As Object
    Dim oNonRefPart As Object
    
    
    Set oPhysicalConn = m_oPOM.GetObject(m_oPOM.MonikerFromDbIdentifier("{" + strPhysicalConnGuid + "}"))
    Set oRefPart = m_oPOM.GetObject(m_oPOM.MonikerFromDbIdentifier("{" + strRefPartGuid + "}"))
    Set oNonRefPart = m_oPOM.GetObject(m_oPOM.MonikerFromDbIdentifier("{" + strNonRefPartGuid + "}"))
    
    Dim oPCConnection As IJAppConnection
    Set oPCConnection = oPhysicalConn
    
    Dim oPCGeometryColl As Collection
    Dim oRefVecColl As Collection
    Dim oRefBevelColl As Collection
    Dim oNonRefVecColl As Collection
    Dim oNonRefBevelColl As Collection
    
    Dim oEntityHelper As MfgEntityHelper
    Set oEntityHelper = New MfgEntityHelper

    oEntityHelper.GetPCWeldingInfo oPCConnection, oPCGeometryColl, oRefVecColl, oNonRefVecColl, oRefBevelColl, oNonRefBevelColl
    
    Dim oPhyConCurve As IJCurve
    Dim oSMSConnectionNode As IXMLDOMNode
    Dim oSMSConnectionElem As IXMLDOMElement
     
    Dim i As Integer
     For i = 1 To oPCGeometryColl.Count
        Set oPhyConCurve = oPCGeometryColl.Item(i)
      
        Dim oSMSJointNode As IXMLDOMNode
        Set oSMSJointNode = oOutputXML.createNode(NODE_ELEMENT, "SMS_JOINT", "")
                
        Dim oSMSJointElem As IXMLDOMElement
        Set oSMSJointElem = oSMSJointNode
        
        oSMSJointElem.setAttribute "DIRECTION", "FLAT"
        oSMSJointElem.setAttribute "PASSES", "UNKNOWN"
        oSMSJointElem.setAttribute "DEFINITION", "TYPE_123"
        
        GenerateGeometryNode oOutputXML, oSMSJointNode, oPhyConCurve

        If oRefBevelColl.Count >= i Then
            GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oRefPart, "REFPART", oRefVecColl.Item(i), oRefBevelColl(i)
        Else
            GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oRefPart, "REFPART", oRefVecColl.Item(i), Nothing
        End If
        
        If oNonRefBevelColl.Count >= i Then
            GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oNonRefPart, "NONREFPART", oNonRefVecColl.Item(i), oNonRefBevelColl(i)
        Else
            GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oNonRefPart, "NONREFPART", oNonRefVecColl.Item(i), Nothing
        End If
        

        oParentNode.appendChild oSMSJointNode

    Next
    
CleanUp:
    Set oSDPhysicalConn = Nothing
    Set oRefPart = Nothing
    Set oNonRefPart = Nothing
    Exit Sub

ErrorHandler:
    'Err.Raise Err.Number
    GoTo CleanUp
End Sub

Private Function GeneratePartInfoNode(ByVal oOutputXML As IXMLDOMDocument, ByRef oParentNode As IXMLDOMNode, ByVal oPhysicalConn As Object, ByVal oPart As Object, ByVal oReference As String, ByVal oNormalVec As IJDVector, ByVal oBevelParamCol As Collection)
Const METHOD = "GeneratePartInfoNode"
On Error GoTo ErrorHandler

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Part (Reference or NonReference)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim dPartThickness As Double
    Dim oPlate As IJPlate
    Dim oStiffener As IJStiffener
    Dim oBean As IJBeam
    Dim oNamedItem As IJNamedItem
    Dim oStructMaterial As IJStructureMaterial
        
    Dim oSDProfile As StructDetailObjects.ProfilePart
    Set oSDProfile = New StructDetailObjects.ProfilePart
    
    Dim strPhysicalConnGuid As String

    Dim strPartGuid As String
    Dim strPartMaterial As String
    Dim strPartGrade As String
    Dim strPartCategory As String

    If TypeOf oPart Is IJPlate Then
        Set oPlate = oPart
        dPartThickness = oPlate.thickness
        strPartCategory = GetShortDescription("StructPlateType", oPlate.plateType)
        Set oPlate = Nothing
    Else
        Set oSDProfile.object = oPart
        dPartThickness = oSDProfile.WebThickness
        
        If TypeOf oPart Is IJStiffener Then
            Set oStiffener = oPart
            strPartCategory = GetShortDescription("StructProfileType", oStiffener.pType)
            Set oStiffener = Nothing
        End If
    End If
    
    Set oStructMaterial = oPart
    If Not oStructMaterial Is Nothing Then
        strPartMaterial = oStructMaterial.MaterialName
        strPartGrade = oStructMaterial.MaterialGrade
    End If
    Set oStructMaterial = Nothing

    Dim oPartInfoElem As IXMLDOMElement
    Dim oPartInfoNode As IXMLDOMNode
    Dim oCVGVectorElem As IXMLDOMElement
    Dim oCVGVectorNode As IXMLDOMNode
    Dim oSMSBevelSideElem As IXMLDOMElement
    Dim oSMSBevelSideNode As IXMLDOMNode
    
    Set oPartInfoNode = oOutputXML.createNode(NODE_ELEMENT, "SMS_PART_INFO", "")
    Set oPartInfoElem = oPartInfoNode
    
    Set oNamedItem = oPart
    oPartInfoElem.setAttribute "NAME", oNamedItem.Name
    oPartInfoElem.setAttribute "GUID", CleanGUID(m_oPOM.DbIdentifierFromMoniker(m_oPOM.GetObjectMoniker(oPart)))
    
    'Convert it to MM
    dPartThickness = dPartThickness * 1000
    If TypeOf oPart Is IJPlate Then
        oPartInfoElem.setAttribute "PART_TYPE", "PLATE"
        oPartInfoElem.setAttribute "THICKNESS", CStr(Format(dPartThickness, "0.00"))
    Else
        oPartInfoElem.setAttribute "PART_TYPE", "PROFILE"
        oPartInfoElem.setAttribute "THICKNESS", CStr(Format(dPartThickness, "0.00"))
    End If

    oPartInfoElem.setAttribute "CATERGORY", strPartCategory
    oPartInfoElem.setAttribute "MATERIAL", strPartMaterial
    oPartInfoElem.setAttribute "GRADE", strPartGrade
    oPartInfoElem.setAttribute "REFERENCE", oReference

    Set oCVGVectorNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_VECTOR", "")
    Set oCVGVectorElem = oCVGVectorNode

    Dim dVectorX As Double, dVectorY As Double, dVectorZ As Double
    dVectorX = 0#
    dVectorY = 0#
    dVectorZ = 0#

    If Not oNormalVec Is Nothing Then
        dVectorX = oNormalVec.x
        dVectorY = oNormalVec.y
        dVectorZ = oNormalVec.z
    End If

    oCVGVectorElem.setAttribute "PURPOSE", "PART"
    oCVGVectorElem.setAttribute "X", CStr(Format(dVectorX, "0.00"))
    oCVGVectorElem.setAttribute "Y", CStr(Format(dVectorY, "0.00"))
    oCVGVectorElem.setAttribute "Z", CStr(Format(dVectorZ, "0.00"))

    oPartInfoNode.appendChild oCVGVectorNode

    'SMS_BEVEL
    Set oSMSBevelSideNode = oOutputXML.createNode(NODE_ELEMENT, "SMS_BEVEL_SIDE", "")
    Set oSMSBevelSideElem = oSMSBevelSideNode

    'Collection has set of  BevelDepth, BevelAngle, ChamferFlag
    'Order is Chamfer, C, A, N Bevel
    Dim nCount As Long
    Dim dChamfer As Double: dChamfer = 0#
    Dim dChamfer_Angle As Double: dChamfer_Angle = 0#
    Dim dA As Double: dA = 0#
    Dim dA_Angle As Double: dA_Angle = 0#
    Dim dC As Double: dC = 0#
    Dim dC_Angle As Double: dC_Angle = 0#
    Dim dN As Double: dN = 0#
    Dim dN_Angle As Double: dN_Angle = 0#

    If Not oBevelParamCol Is Nothing Then
        nCount = oBevelParamCol.Count
        Dim i As Long
        For i = 1 To nCount Step 3
            Dim dDepth As Double
            Dim dAngle As Double
            Dim bChamfer As Boolean

            'Convert it to MM
            dDepth = oBevelParamCol.Item(i) * 1000
            'Convert it to Degree
            dAngle = oBevelParamCol.Item(i + 1) * 180 / PI
            bChamfer = oBevelParamCol.Item(i + 2)

            If bChamfer = True Then
                dChamfer = dDepth
                dChamfer_Angle = dAngle
            Else
                If (nCount - i + 1) Mod 9 = 0 Then ' C
                    dC = dDepth
                    dC_Angle = dAngle
                ElseIf (nCount - i + 1) Mod 6 = 0 Then 'A
                    dA = dDepth
                    dA_Angle = dAngle
                ElseIf (nCount - i + 1) Mod 3 = 0 Then 'N
                    dN = dDepth
                    dN_Angle = dAngle
                Else
                    'No
                End If
            End If


        Next i
    End If

    oSMSBevelSideElem.setAttribute "PURPOSE", "MANUFACTURING"
    oSMSBevelSideElem.setAttribute "CHAMFER", CStr(Format(dChamfer, "0.00"))
    oSMSBevelSideElem.setAttribute "CHAMFER_ANGLE", CStr(Format(dChamfer_Angle, "0.000"))
    oSMSBevelSideElem.setAttribute "A", CStr(Format(dA, "0.00"))
    oSMSBevelSideElem.setAttribute "A_ANGLE", CStr(Format(dA_Angle, "0.000"))
    oSMSBevelSideElem.setAttribute "C", CStr(Format(dC, "0.00"))
    oSMSBevelSideElem.setAttribute "C_ANGLE", CStr(Format(dC_Angle, "0.000"))
    oSMSBevelSideElem.setAttribute "N", CStr(Format(dN, "0.00"))
    oSMSBevelSideElem.setAttribute "N_ANGLE", CStr(Format(dN_Angle, "0.000"))

    oPartInfoNode.appendChild oSMSBevelSideNode

    oParentNode.appendChild oPartInfoNode

    oPartInfoElem = Nothing
    oPartInfoNode = Nothing
    oCVGVectorElem = Nothing
    oCVGVectorNode = Nothing
    oSMSBevelSideElem = Nothing
    oSMSBevelSideNode = Nothing
    oNamedItem = Nothing
    oSDProfile = Nothing
    oPlate = Nothing

ErrorHandler:
    'Err.Raise Err.Number

End Function

Private Sub GenerateGeometryNode(ByVal oOutputXML As IXMLDOMDocument, ByRef oParentNode As IXMLDOMNode, ByRef oPhyConCurve As IJCurve)
    Const METHOD = "GenerateGeometryNode"
    On Error GoTo ErrorHandler

    ' Initial version, we only create the end-points as reference
   
    If Not oPhyConCurve Is Nothing Then
        
        Dim dStartX As Double, dStartY As Double, dStartZ As Double
        Dim dEndX As Double, dEndY As Double, dEndZ As Double
        oPhyConCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ


        Dim oSMSEdgeElem As IXMLDOMElement
        Dim oSMSEdgeNode As IXMLDOMNode
        Set oSMSEdgeNode = oOutputXML.createNode(NODE_ELEMENT, "SMS_EDGE", "")
        Set oSMSEdgeElem = oSMSEdgeNode

        Dim oCVGCurveElem As IXMLDOMElement
        Dim oCVGCurveNode As IXMLDOMNode
        Set oCVGCurveNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_CURVE", "")
        Set oCVGCurveElem = oCVGCurveNode

        Dim oCVGVertexElem As IXMLDOMElement
        Dim oCVGVertexNode As IXMLDOMNode
        Set oCVGVertexNode = oOutputXML.createNode(NODE_ELEMENT, "CVG_VERTEX", "")
        Set oCVGVertexElem = oCVGVertexNode

        'Convert it to MM
        dStartX = dStartX * 1000
        dStartY = dStartY * 1000
        dStartZ = dStartZ * 1000

        oCVGVertexElem.setAttribute "POINT_CODE", "s_point"
        oCVGVertexElem.setAttribute "SEG_TYPE", "line"
        oCVGVertexElem.setAttribute "SX", CStr(Format(dStartX, "0.00"))
        oCVGVertexElem.setAttribute "SY", CStr(Format(dStartY, "0.00"))
        oCVGVertexElem.setAttribute "SZ", CStr(Format(dStartZ, "0.00"))

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
        oCVGVertexElem.setAttribute "SEG_TYPE", "line"
        oCVGVertexElem.setAttribute "X", CStr(Format(dEndX, "0.00"))
        oCVGVertexElem.setAttribute "Y", CStr(Format(dEndY, "0.00"))
        oCVGVertexElem.setAttribute "Z", CStr(Format(dEndZ, "0.00"))

        oCVGCurveNode.appendChild oCVGVertexNode
        
        oSMSEdgeNode.appendChild oCVGCurveNode
        oParentNode.appendChild oSMSEdgeNode

    End If

    Exit Sub
ErrorHandler:
    Resume Next

End Sub

Private Function CleanGUID(sGUID As String) As String
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
