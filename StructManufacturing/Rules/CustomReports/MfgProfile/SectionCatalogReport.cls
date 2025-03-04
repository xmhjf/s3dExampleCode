VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SectionCatXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2003,2008 Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        : SectionCatXML.cls
''
''  Description : Populates the XML DOM Document with section information of profiles
''  and members and with plate thickness.
''
''
''  Authors     : Brandon Affenzeller
''
''  Notes       :
''
''**************************************************************************************
Implements IJDCustomReport

Private Const CATALOGDATABASE = "Catalog"
Private Const MODELDATABASE = "Model"
Private Const CONNECTMIDDLE = "ConnectMiddle"
Private Const DBTYPECONFIG = "DBTypeConfiguration"

Private m_strConnType As String

Private Const MODULE = "MfgProfileCustomReports.SectionCatXML"
Private Const m_xmlns = "http://www.ControlledAutomation.com/ProductionControl/CAShapes.xsd"

Private m_objXMLDoc As DOMDocument

Private Sub Class_Initialize()
'    InitUOMService
'    Set m_oPOM = GetPOM
End Sub

Private Sub Class_Terminate()
    Set m_objXMLDoc = Nothing
    'Set m_oPOM = Nothing
End Sub

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
    
    Dim oATeckCatalogMapDOM As New DOMDocument
    Dim strCatalogMapLoc As String
    Dim strSymbolShare As String
    Dim oContext As IJContext
    
    Set oContext = GetJContext()
    strSymbolShare = oContext.GetVariable("OLE_SERVER")
    
    strCatalogMapLoc = strSymbolShare & "\StructManufacturing\DSTV_OUTPUT\ATeK_OUTPUT\CATALOG_MAP.xml"
    
    If Not oATeckCatalogMapDOM.Load(strCatalogMapLoc) Then
        eCustomReportStatus = StrMfgErrorUnknown
        Exit Sub
    End If
    
    'Get reference to file
    Set m_objXMLDoc = New DOMDocument
    
    Dim oRootNode As IXMLDOMNode
    Set oRootNode = m_objXMLDoc.createNode(1, "CAShapes", m_xmlns)
    m_objXMLDoc.appendChild oRootNode
    
    'Query for profile and memeber section information
    Dim strTestQuery As String

    Dim oQueryCatalog As IJQueryCatalog
    Set oQueryCatalog = GetQueryCatalog()
    
     If StrComp(m_strConnType, "ORACLE", vbTextCompare) = 0 Then
             strTestQuery = "SELECT DISTINCT JDPC.Name As PartClassName, JSCS.SectionName, JSXS.SectionTypeName, RDCST.name As CrossSectionType, JRS.Name As ReferenceStandard, XSW.WebLength, XSW.WebThickness, XSF.FlangeLength, XSF.FlangeThickness, JSFBG.gf As FlangeGage,  JSCSD.Depth As OuterDiameter, JSCSD.Area * JDM.Density As WeightPerUnitLength, JSS.tnom As NominalThickness " & _
                    "FROM JDPartClass JDPC " & _
                    "INNER JOIN XReferenceStdHasPartClasses XRSHPC ON XRSHPC.OidDestination = JDPC.Oid " & _
                    "INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin " & _
                    "INNER JOIN JDCrossSection JDCS ON JDCS.Type = JDPC.Name " & _
                    "INNER JOIN JSTRUCTCrossSectionDimensions JSCSD ON JSCSD.Oid = JDCS.Oid " & _
                    "INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = JDCS.Oid " & _
                    "INNER JOIN JStructXSection JSXS ON JSXS.oid = JDCS.oid " & _
                    "LEFT JOIN XCrossSectionClassToType XCSCT ON XCSCT.oidOrigin = JDPC.oid " & _
                    "LEFT JOIN REFDATCrossSectionType RDCST ON RDCST.oid = XCSCT.oidDestination " & _
                    "LEFT JOIN JUAHSS JSS ON JSS.oid = JDCS.oid " & _
                    "LEFT JOIN JUAHSSC JSSC ON JSSC.oid = JDCS.oid " & _
                    "LEFT JOIN JSTRUCTFlangedBoltGage JSFBG ON JSFBG.oid = JDCS.oid " & _
                    "LEFT JOIN JUAXSectionWeb XSW ON XSW.Oid = JDCS.Oid " & _
                    "LEFT JOIN JUAXSectionFlange XSF ON XSF.Oid = JDCS.Oid, " & _
                    "(SELECT Density FROM JDMaterial M WHERE M.MaterialType = 'Steel - Carbon' AND M.MaterialGrade = 'A') As JDM " & _
                    "WHERE NOT EXISTS(SELECT * FROM COREBoolAttribute CBA WHERE CBA.iid =HEXTORAW('0742B82CEABE494B829AE66C2A6D1CEE') AND JDPC.oid = CBA.oid) " & _
                    "ORDER BY JDPC.Name, JSCS.SectionName"
    Else
             strTestQuery = "SELECT DISTINCT JDPC.Name As PartClassName, JSCS.SectionName, JSXS.SectionTypeName, RDCST.name As CrossSectionType, JRS.Name As ReferenceStandard, XSW.WebLength, XSW.WebThickness, XSF.FlangeLength, XSF.FlangeThickness, JSFBG.gf As FlangeGage,  JSCSD.Depth As OuterDiameter, JSCSD.Area * JDM.Density As WeightPerUnitLength, JSS.tnom As NominalThickness " & _
                    "FROM JDPartClass JDPC " & _
                    "INNER JOIN XReferenceStdHasPartClasses XRSHPC ON XRSHPC.OidDestination = JDPC.Oid " & _
                    "INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin " & _
                    "INNER JOIN JDCrossSection JDCS ON JDCS.Type = JDPC.Name " & _
                    "INNER JOIN JSTRUCTCrossSectionDimensions JSCSD ON JSCSD.Oid = JDCS.Oid " & _
                    "INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = JDCS.Oid " & _
                    "INNER JOIN JStructXSection JSXS ON JSXS.oid = JDCS.oid " & _
                    "LEFT JOIN XCrossSectionClassToType XCSCT ON XCSCT.oidOrigin = JDPC.oid " & _
                    "LEFT JOIN REFDATCrossSectionType RDCST ON RDCST.oid = XCSCT.oidDestination " & _
                    "LEFT JOIN JUAHSS JSS ON JSS.oid = JDCS.oid " & _
                    "LEFT JOIN JUAHSSC JSSC ON JSSC.oid = JDCS.oid " & _
                    "LEFT JOIN JSTRUCTFlangedBoltGage JSFBG ON JSFBG.oid = JDCS.oid " & _
                    "LEFT JOIN JUAXSectionWeb XSW ON XSW.Oid = JDCS.Oid " & _
                    "LEFT JOIN JUAXSectionFlange XSF ON XSF.Oid = JDCS.Oid, " & _
                    "(SELECT Density FROM JDMaterial M WHERE M.MaterialType = 'Steel - Carbon' AND M.MaterialGrade = 'A') As JDM " & _
                    "WHERE NOT EXISTS(SELECT * FROM COREBoolAttribute CBA WHERE CBA.iid ='0742B82C-EABE-494B-829A-E66C2A6D1CEE' AND JDPC.oid = CBA.oid) " & _
                    "ORDER BY JDPC.Name, JSCS.SectionName"
    End If


    Dim varRS As Variant
    oQueryCatalog.ExecuteQuery strTestQuery, , varRS
    
    Dim oRecordset As ADODB.Recordset
    Set oRecordset = varRS
    
    Dim oSectionNode As IXMLDOMNode
    Dim oShapeNode As IXMLDOMNode
    

    Dim oUoM            As IJUomVBInterface
    Set oUoM = New UnitsOfMeasureServicesLib.UomVBInterface
    
    'For each row in profile and member table
    While Not oRecordset.EOF
    
        Dim strSectionClass As String
        Dim strSectionName As String
        Dim strSectionTypeName As String
        Dim strCrossSectionType As String
        Dim iMaterialType As Integer
        Dim dWebLength As Double
        Dim dWebThickness As Double
        Dim dFlangeLength As Double
        Dim dFlangeThickness As Double
        Dim dFlangeGage As Double
        Dim dWeightPerLength As Double
        Dim iMeasuringSystem As Integer
        Dim dOuterDiameter As Double
        Dim dInnerDiameter As Double
        Dim dTubeThickness As Double
        Dim dTubeDiameter As Double
        Dim oMapElement As IXMLDOMElement
        Dim oSectionMapNameNodeList As IXMLDOMNodeList
        
        strSectionClass = oRecordset.fields.Item("PartClassName").Value
        strSectionName = oRecordset.fields.Item("SectionName").Value
        strSectionTypeName = oRecordset.fields.Item("SectionTypeName").Value
        
        If Not IsNull(oRecordset.fields.Item("CrossSectionType").Value) Then
            strCrossSectionType = oRecordset.fields.Item("CrossSectionType").Value
        Else
            strCrossSectionType = "-1"
        End If
        
        Dim oMapNode As IXMLDOMNode
        Set oMapNode = oATeckCatalogMapDOM.selectSingleNode("/SMS_CATALOG_MAP/MATERIAL_TYPE_MAPPING/XSECTION_TYPE[@S3D_NAME='" & strSectionTypeName & "']")
        
        If Not oMapNode Is Nothing Then
            Set oMapElement = oMapNode
            iMaterialType = CInt(oMapElement.getAttribute("ATek_NAME"))
            Set oMapElement = Nothing
        Else
            iMaterialType = -1
            GoTo Continue
        End If
        
        If Not IsNull(oRecordset.fields.Item("WebLength").Value) Then
            dWebLength = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("WebLength").Value, DISTANCE_INCH)
        Else
            dWebLength = 0
        End If
        
        If Not IsNull(oRecordset.fields.Item("WebThickness").Value) Then
            dWebThickness = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("WebThickness").Value, DISTANCE_INCH)
        Else
            dWebThickness = 0
        End If
        
        If Not IsNull(oRecordset.fields.Item("FlangeLength").Value) Then
            dFlangeLength = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("FlangeLength").Value, DISTANCE_INCH)
        Else
            dFlangeLength = 0
        End If
        
        If Not IsNull(oRecordset.fields.Item("FlangeThickness").Value) Then
            dFlangeThickness = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("FlangeThickness").Value, DISTANCE_INCH)
        Else
            dFlangeThickness = 0
        End If
        
        If Not IsNull(oRecordset.fields.Item("FlangeGage").Value) Then
            dFlangeGage = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("FlangeGage").Value, DISTANCE_INCH)
        Else
            dFlangeGage = 0
        End If
        
        iMeasuringSystem = 2
        dWeightPerLength = oRecordset.fields.Item("WeightPerUnitLength").Value
        
        If Not strCrossSectionType <> "CircTube" And Not IsNull(oRecordset.fields.Item("OuterDiameter").Value) Then
            dOuterDiameter = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("OuterDiameter").Value, DISTANCE_INCH)
        Else
            dOuterDiameter = 0
        End If
        
        If Not IsNull(oRecordset.fields.Item("NominalThickness").Value) Then
            dTubeThickness = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("NominalThickness").Value, DISTANCE_INCH)
        Else
            dTubeThickness = 0
        End If
        
        If dTubeThickness > 0 Then
            dInnerDiameter = dOuterDiameter - 2 * dTubeThickness
        Else
            dInnerDiameter = 0
        End If
        
'        If Not IsNull(oRecordset.fields.Item("Diameter").Value) Then
'            dTubeDiameter = FormatNumber(oUOM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("Diameter").Value, DISTANCE_INCH), 3)
'        Else
            dTubeDiameter = 0
        'End If
        
        Set oSectionNode = m_objXMLDoc.selectSingleNode("//Section[@Section='" & strSectionClass & "' and @MaterialType='" & CStr(iMaterialType) & "']")
        
        If oSectionNode Is Nothing Then
        
            Set oSectionMapNameNodeList = Nothing
            Set oSectionMapNameNodeList = oATeckCatalogMapDOM.selectNodes("/SMS_CATALOG_MAP/SECTION_MAP_NAME_MAPPING/XSECTION_TYPE[@S3D_NAME='" & strSectionTypeName & "']/XSECTION_MAP_NAME")
    
            Set oSectionNode = CreateSectionNode(strSectionClass, iMaterialType, oSectionMapNameNodeList)
            oRootNode.appendChild oSectionNode
        
        End If
        
        'Create <ShapeData> for the part
        Set oShapeNode = CreateShapeDataNode(strSectionName, strSectionName, dWebLength, dWebThickness, _
                                                dFlangeLength, dFlangeThickness, dFlangeGage, dWeightPerLength, _
                                                iMeasuringSystem, strSectionName, strSectionName, dOuterDiameter, _
                                                dInnerDiameter, dTubeThickness, dTubeDiameter)
                                                
        oSectionNode.appendChild oShapeNode
        
Continue:
        oRecordset.MoveNext
    
    Wend
    
    Set oSectionNode = Nothing
    Set oShapeNode = Nothing
    
    strTestQuery = "SELECT JPD.Thickness FROM JDPlateDimensions JPD"
    varRS = Empty
    oQueryCatalog.ExecuteQuery strTestQuery, , varRS
    Set oRecordset = varRS
    
    While Not oRecordset.EOF
    
        Dim dPlateThickness As Double
        dPlateThickness = oUoM.ConvertDbuToUnit(UNIT_DISTANCE, oRecordset.fields.Item("Thickness").Value, DISTANCE_INCH)
    
        Set oSectionNode = m_objXMLDoc.selectSingleNode("//Section[@Section='" & "PL" & "' and @MaterialType='4']")
        
        If oSectionNode Is Nothing Then
        
            Set oSectionMapNameNodeList = Nothing
            Set oSectionMapNameNodeList = oATeckCatalogMapDOM.selectNodes("/SMS_CATALOG_MAP/SECTION_MAP_NAME_MAPPING/XSECTION_TYPE[@S3D_NAME='Plate']/XSECTION_MAP_NAME")
            
            Set oSectionNode = CreateSectionNode("PL", 4, oSectionMapNameNodeList)
            oRootNode.appendChild oSectionNode
        
        End If
        
        'Create <ShapeData> for the part
        Set oShapeNode = CreateShapeDataNode(CStr(dPlateThickness), CStr(dPlateThickness), 0, dPlateThickness, _
                                                0, 0, 0, 0, 2, "", "", 0, 0, 0, 0)
                                                
        oSectionNode.appendChild oShapeNode
        
        oRecordset.MoveNext
        
    Wend
    
    m_objXMLDoc.save strFileName
    
    eCustomReportStatus = StrMfgProcessFinished
    
Wrapup:
    Set oQueryCatalog = Nothing
    Set oMapElement = Nothing
    Set oSectionMapNameNodeList = Nothing
    Exit Sub
ErrorHandler:
    Set oQueryCatalog = Nothing
    Set oMapElement = Nothing
    Set oSectionMapNameNodeList = Nothing
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function CreateSectionNode(ByVal strPartClass As String, ByVal iMaterialType As Integer, ByVal oSectionMapNameNodeList As IXMLDOMNodeList) As IXMLDOMNode
    Const METHOD = "CreateSectionNode"
    On Error GoTo ErrorHandler
    
    Dim oSectionNode As IXMLDOMNode
    Dim oSectionElement As IXMLDOMElement
    Dim oSectionMappingNode As IXMLDOMNode
    Dim oSectionMappingNodeChildNode As IXMLDOMNode
    Dim oSectionMappingNodeChildElement As IXMLDOMElement
    Dim oSectionMapNameNode As IXMLDOMNode
    Dim strSectionMapName As String
    
    Set oSectionNode = m_objXMLDoc.createNode(1, "Section", m_xmlns)
    
    Set oSectionElement = oSectionNode
    oSectionElement.setAttribute "Section", strPartClass
    oSectionElement.setAttribute "MaterialType", CStr(iMaterialType)
    
    Set oSectionMappingNode = m_objXMLDoc.createNode(1, "Section_Mapping", m_xmlns)
    oSectionElement.appendChild oSectionMappingNode
    
    If Not oSectionMapNameNodeList Is Nothing Then
        For Each oSectionMappingNodeChildNode In oSectionMapNameNodeList
            If Not oSectionMappingNodeChildNode Is Nothing Then
                Set oSectionMappingNodeChildElement = oSectionMappingNodeChildNode
                strSectionMapName = oSectionMappingNodeChildElement.getAttribute("NAME")
                
                Set oSectionMapNameNode = Nothing
                Set oSectionMapNameNode = m_objXMLDoc.createNode(1, "Section_Map_Name", m_xmlns)
                oSectionMapNameNode.nodeTypedValue = strSectionMapName
                oSectionMappingNode.appendChild oSectionMapNameNode
            End If
        Next
    End If
    
    Set CreateSectionNode = oSectionElement

Wrapup:
    Set oSectionNode = Nothing
    Set oSectionElement = Nothing
    Set oSectionMappingNode = Nothing
    Set oSectionMappingNodeChildNode = Nothing
    Set oSectionMappingNodeChildElement = Nothing
    Set oSectionMapNameNode = Nothing
    Exit Function
ErrorHandler:
    Set oSectionNode = Nothing
    Set oSectionElement = Nothing
    Set oSectionMappingNode = Nothing
    Set oSectionMappingNodeChildNode = Nothing
    Set oSectionMappingNodeChildElement = Nothing
    Set oSectionMapNameNode = Nothing
    Err.Raise Err.Number, , Err.Description
End Function

Private Function CreateShapeDataNode(ByVal strSectionName As String, _
                                        ByVal strDisplayString As String, _
                                        ByVal dWebLength As Double, _
                                        ByVal dWebThickness As Double, _
                                        ByVal dFlangeLength As Double, _
                                        ByVal dFlangeThickness As Double, _
                                        ByVal dFlangeGage As Double, _
                                        ByVal dWeightPerUnitLength As Double, _
                                        ByVal iMeasuringSystem As Integer, _
                                        ByVal strWinCADName As String, _
                                        ByVal strCADCADName As String, _
                                        ByVal dOuterDiameter As Double, _
                                        ByVal dInnerDiameter As Double, _
                                        ByVal dNominalThickness As Double, _
                                        ByVal dNominalDiameter As Double) As IXMLDOMNode
    Const METHOD = "CreateShapeDataNode"
    On Error GoTo ErrorHandler

    Dim oShapeRootNode As IXMLDOMNode
    Set oShapeRootNode = m_objXMLDoc.createNode(1, "ShapeData", m_xmlns)

    Dim oShapeNode As IXMLDOMNode
    Set oShapeNode = m_objXMLDoc.createNode(1, "Shape", m_xmlns)
    oShapeNode.nodeTypedValue = strSectionName
    oShapeRootNode.appendChild oShapeNode

    Dim oDisplayStringNode As IXMLDOMNode
    Set oDisplayStringNode = m_objXMLDoc.createNode(1, "FormalDisplayString", m_xmlns)
    oDisplayStringNode.nodeTypedValue = strDisplayString
    oShapeRootNode.appendChild oDisplayStringNode

    Dim oWebDepthNode As IXMLDOMNode
    Set oWebDepthNode = m_objXMLDoc.createNode(1, "WebDepth", m_xmlns)
    oWebDepthNode.nodeTypedValue = CStr(dWebLength)
    oShapeRootNode.appendChild oWebDepthNode

    Dim oWebThicknessNode As IXMLDOMNode
    Set oWebThicknessNode = m_objXMLDoc.createNode(1, "WebThickness", m_xmlns)
    oWebThicknessNode.nodeTypedValue = CStr(dWebThickness)
    oShapeRootNode.appendChild oWebThicknessNode

    Dim oFlangeBreathNode As IXMLDOMNode
    Set oFlangeBreathNode = m_objXMLDoc.createNode(1, "BreathOfFlange", m_xmlns)
    oFlangeBreathNode.nodeTypedValue = CStr(dFlangeLength)
    oShapeRootNode.appendChild oFlangeBreathNode
    
    Dim oFlangeThicknessNode As IXMLDOMNode
    Set oFlangeThicknessNode = m_objXMLDoc.createNode(1, "FlangeThickness", m_xmlns)
    oFlangeThicknessNode.nodeTypedValue = CStr(dFlangeThickness)
    oShapeRootNode.appendChild oFlangeThicknessNode
    
    Dim oFlangeGageNode As IXMLDOMNode
    Set oFlangeGageNode = m_objXMLDoc.createNode(1, "StandardFlangeGage", m_xmlns)
    oFlangeGageNode.nodeTypedValue = CStr(dFlangeGage)
    oShapeRootNode.appendChild oFlangeGageNode
    
    Dim oWeightPerLengthNode As IXMLDOMNode
    Set oWeightPerLengthNode = m_objXMLDoc.createNode(1, "WeightPerFoot", m_xmlns)
    oWeightPerLengthNode.nodeTypedValue = CStr(dWeightPerUnitLength)
    oShapeRootNode.appendChild oWeightPerLengthNode
    
    Dim oMeasuringSystemNode As IXMLDOMNode
    Set oMeasuringSystemNode = m_objXMLDoc.createNode(1, "MeasuringSystem", m_xmlns)
    oMeasuringSystemNode.nodeTypedValue = CStr(iMeasuringSystem)
    oShapeRootNode.appendChild oMeasuringSystemNode
    
    Dim oWinCADNode As IXMLDOMNode
    Set oWinCADNode = m_objXMLDoc.createNode(1, "WinCAd_Name", m_xmlns)
    oWinCADNode.nodeTypedValue = strWinCADName
    oShapeRootNode.appendChild oWinCADNode
    
    Dim oCADCADNode As IXMLDOMNode
    Set oCADCADNode = m_objXMLDoc.createNode(1, "CADCAD_Name", m_xmlns)
    oCADCADNode.nodeTypedValue = strCADCADName
    oShapeRootNode.appendChild oCADCADNode
    
    Dim oOuterDiameterNode As IXMLDOMNode
    Set oOuterDiameterNode = m_objXMLDoc.createNode(1, "OD", m_xmlns)
    oOuterDiameterNode.nodeTypedValue = CStr(dOuterDiameter)
    oShapeRootNode.appendChild oOuterDiameterNode
    
    Dim oInnerDiameterNode As IXMLDOMNode
    Set oInnerDiameterNode = m_objXMLDoc.createNode(1, "ID", m_xmlns)
    oInnerDiameterNode.nodeTypedValue = CStr(dInnerDiameter)
    oShapeRootNode.appendChild oInnerDiameterNode
    
    Dim oTubeThicknessNode As IXMLDOMNode
    Set oTubeThicknessNode = m_objXMLDoc.createNode(1, "Nominal_Thickness", m_xmlns)
    oTubeThicknessNode.nodeTypedValue = CStr(dNominalThickness)
    oShapeRootNode.appendChild oTubeThicknessNode
    
    Dim oTubeDiameterNode As IXMLDOMNode
    Set oTubeDiameterNode = m_objXMLDoc.createNode(1, "Nominal_Diameter", m_xmlns)
    oTubeDiameterNode.nodeTypedValue = CStr(dNominalDiameter)
    oShapeRootNode.appendChild oTubeDiameterNode
    
    Dim oShapeMappingNode As IXMLDOMNode
    Set oShapeMappingNode = m_objXMLDoc.createNode(1, "Shape_Mapping", m_xmlns)
    oShapeRootNode.appendChild oShapeMappingNode
    
    Set CreateShapeDataNode = oShapeRootNode

Wrapup:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function GetQueryCatalog() As Object
    Const METHOD = "GetQueryCatalog"
    On Error GoTo ErrorHandler

    Dim jContext As IJContext
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim oCatalogResourceMgr As IUnknown
    Dim oCatalogPOM As IJDPOM
    Dim strCatalogDBID As String
    
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService(DBTYPECONFIG)
    Set oConnectMiddle = jContext.GetService(CONNECTMIDDLE)

    strCatalogDBID = oDBTypeConfig.get_DataBaseFromDBType(CATALOGDATABASE)
    
    Set oCatalogResourceMgr = oConnectMiddle.GetResourceManager(strCatalogDBID)
    Set oCatalogPOM = oCatalogResourceMgr
    
    m_strConnType = oCatalogPOM.ConnectionType

    Set GetQueryCatalog = oCatalogPOM.GetObject("RefDataCacheManager")
    
    
Wrapup:
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oCatalogResourceMgr = Nothing
    Set oCatalogPOM = Nothing
    Exit Function
ErrorHandler:
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oCatalogResourceMgr = Nothing
    Set oCatalogPOM = Nothing
    Err.Raise Err.Number, , Err.Description
End Function
