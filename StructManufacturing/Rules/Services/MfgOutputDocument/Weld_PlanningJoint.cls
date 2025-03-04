VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Weld_PlanningJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJMfgOutputDocument


Private m_oControlInputDoc As IXMLDOMDocument
Private m_bstrOutputOption As String
Private m_oPOM As IJDPOM
Private PI As Double
Private Const IID_IJStructGeometry As String = "{6034AD40-FA0B-11D1-B2FD-080036024603}"
Private Const IID_IJPlnJoint As String = "{70FA6719-EAA3-4f49-9E59-0D79C852CE23}"
Private Const strPlanningJointColName As String = "PlnJointPhysConn_ORIG"


Private Sub Class_Initialize()
    Set m_oControlInputDoc = Nothing
    Set m_oPOM = GetResourceManager
    PI = 3.141592
End Sub

Private Sub Class_Terminate()
    Set m_oControlInputDoc = Nothing
    Set m_oPOM = Nothing
    
End Sub

Private Function IJMfgOutputDocument_Generate(ByVal pObject As Object, ByVal pOutputXMLObject As Object, ByVal bstrBaseOutputName As String) As Object
Const METHOD = "IJMfgOutputDocument_Generate"
On Error GoTo ErrorHandler

    If m_oPOM Is Nothing Then GoTo CleanUp
     
    Dim oMfgXMLHelper As IJMfgXMLHelper
    Set oMfgXMLHelper = New MfgXMLHelper
    
    Dim strOutputXML As String
    oMfgXMLHelper.GetSMSWeldElement Nothing, pObject, MfgWeldPlanningJoint, strOutputXML
    
    Dim oPlanningHelper As IJDPlnIntHelper
    Set oPlanningHelper = New CPlnIntHelper
    
    Dim oPlanningJoints As IJElements
    Set oPlanningJoints = oPlanningHelper.GetPlnJointsFromAssembly(pObject)
    
    '1. Get Associated Physical Connection from PlanningJoint
    
    Dim oPhysicalConns As IJElements
    Set oPhysicalConns = GetPhysicalConnection(oPlanningJoints)
    
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
    GenerateAssemblyNode oOutputXML, pObject, oSMSWeldingNode
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '3. Create SMS_CONNECTION Trace Node
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oSMSConnectionNodeList As IXMLDOMNodeList
    Dim oSMSConnectionNode As IXMLDOMNode
    Set oSMSConnectionNodeList = oOutputXML.selectNodes("//SMS_CONNECTION")
       
    For Each oSMSConnectionNode In oSMSConnectionNodeList ' For each SMS_CONNECTION
        GenerateJointNode oOutputXML, oSMSConnectionNode, oPlanningJoints
    Next oSMSConnectionNode
    
    Dim strGuid As String
    strGuid = m_oPOM.DbIdentifierFromMoniker(m_oPOM.GetObjectMoniker(pObject))
     
    Dim strExtension As String
    strExtension = GetRuleValue(m_bstrOutputOption, "DocumentType")
    
    Dim strOutputDocumentName As String
    strOutputDocumentName = bstrBaseOutputName + "_Welds_" + strGuid + strExtension
         
    If bSucess = True Then
        oOutputXML.save strOutputDocumentName
    End If
     
    'Generate Output ReportDocument
    Dim oOutputReportDoc As New DOMDocument
    Dim oOutputReportNode As IXMLDOMNode
    Set oOutputReportNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputReport", "")
    oOutputReportDoc.appendChild oOutputReportNode
    
    Dim oOutputDocumentNode As IXMLDOMNode
    Set oOutputDocumentNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputDocument", "")
    
    Dim oOutputDocumentElem As IXMLDOMElement
    Set oOutputDocumentElem = oOutputDocumentNode
    
    oOutputDocumentElem.setAttribute "ID", strGuid
    
    Dim lastPos As Long
    lastPos = InStrRev(strOutputDocumentName, "\")
    Dim strFileName As String
    strFileName = Trim(Right$(strOutputDocumentName, Len(strOutputDocumentName) - lastPos))
    oOutputDocumentElem.setAttribute "NAME", strFileName
    
    oOutputReportNode.appendChild oOutputDocumentNode
    Set IJMfgOutputDocument_Generate = oOutputReportDoc

CleanUp:
    Set m_oControlInputDoc = Nothing
    Set oOutputXML = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Sub IJMfgOutputDocument_Initialize(ByVal pControlXMLObject As Object, ByVal bstrOutputOption As String, ByVal bstrSharedContentsPath As String)
Const METHOD = "IJMfgOutputDocument_Initialize"
On Error GoTo ErrorHandler

    m_bstrOutputOption = bstrOutputOption
    If Not pControlXMLObject Is Nothing Then
        If TypeOf pControlXMLObject Is IXMLDOMDocument Then
            Set m_oControlInputDoc = pControlXMLObject
        End If
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    'Err.Raise Err.Number
    GoTo CleanUp
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
        oAssemblyElem.setAttribute "SEQUENCE", Str(oAssemblySequence.ChildIndex(oAssembly))
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
                    
    
    oCVGMatrixElem.setAttribute "M00", ConvertNumberToENLocale(oViewMatrix.IndexValue(0))
    oCVGMatrixElem.setAttribute "M01", ConvertNumberToENLocale(oViewMatrix.IndexValue(1))
    oCVGMatrixElem.setAttribute "M02", ConvertNumberToENLocale(oViewMatrix.IndexValue(2))
    oCVGMatrixElem.setAttribute "M03", ConvertNumberToENLocale(oViewMatrix.IndexValue(3))
    
    oCVGMatrixElem.setAttribute "M10", ConvertNumberToENLocale(oViewMatrix.IndexValue(4))
    oCVGMatrixElem.setAttribute "M11", ConvertNumberToENLocale(oViewMatrix.IndexValue(5))
    oCVGMatrixElem.setAttribute "M12", ConvertNumberToENLocale(oViewMatrix.IndexValue(6))
    oCVGMatrixElem.setAttribute "M13", ConvertNumberToENLocale(oViewMatrix.IndexValue(7))
    
    oCVGMatrixElem.setAttribute "M20", ConvertNumberToENLocale(oViewMatrix.IndexValue(8))
    oCVGMatrixElem.setAttribute "M21", ConvertNumberToENLocale(oViewMatrix.IndexValue(9))
    oCVGMatrixElem.setAttribute "M22", ConvertNumberToENLocale(oViewMatrix.IndexValue(10))
    oCVGMatrixElem.setAttribute "M23", ConvertNumberToENLocale(oViewMatrix.IndexValue(11))
    
    oCVGMatrixElem.setAttribute "M30", ConvertNumberToENLocale(oViewMatrix.IndexValue(12))
    oCVGMatrixElem.setAttribute "M31", ConvertNumberToENLocale(oViewMatrix.IndexValue(13))
    oCVGMatrixElem.setAttribute "M32", ConvertNumberToENLocale(oViewMatrix.IndexValue(14))
    oCVGMatrixElem.setAttribute "M33", ConvertNumberToENLocale(oViewMatrix.IndexValue(15))
    
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


Private Sub GenerateJointNode(ByVal oOutputXML As IXMLDOMDocument, ByRef oParentNode As IXMLDOMNode, ByVal oPlanningJointColl As IJElements)
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
    If IsRootPC(oPhysicalConn) Then
    
        On Error Resume Next
        Dim oSMSWeldNode As IXMLDOMNode
        Set oSMSWeldNode = oParentNode.parentNode
        
        If Not oSMSWeldNode Is Nothing Then
            oSMSWeldNode.removeChild oParentNode
            GoTo CleanUp
        End If
    End If
    
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
    
    
    Set oNamedItem = oPhysicalConn
    oEntityHelper.GetPCWeldingInfo oPCConnection, oPCGeometryColl, oRefVecColl, oNonRefVecColl, oRefBevelColl, oNonRefBevelColl
    
 '''''''''' Dump Geometry
'    Dim FileName As String
'    Dim k As Integer
'
'        Dim oPCGeomModelBody As IJDModelBody
'        On Error Resume Next
'        Set oPCGeomModelBody = oPCConnection
'        If Not oPCGeomModelBody Is Nothing Then
'
'            FileName = Environ("TEMP")
'            If FileName = "" Or FileName = vbNullString Then
'                FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
'            End If
'
'            FileName = FileName & "\" + oNamedItem.Name + "_" + CStr(oPlanningJointColl.Count) + ".sat"
'
'            oPCGeomModelBody.DebugToSATFile FileName
'            MsgBox FileName + " dumped"
'            Set oPCGeomModelBody = Nothing
'        End If
'        On Error GoTo ErrorHandler
      

    
    'Dim oPhyConCurve As IJCurve
    Dim oSMSConnectionNode As IXMLDOMNode
    Dim oSMSConnectionElem As IXMLDOMElement
    
    Dim debugFileLog As String
    debugFileLog = Environ("TEMP")
    If debugFileLog = "" Or debugFileLog = vbNullString Then
        debugFileLog = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
    End If
    debugFileLog = debugFileLog & "\WeldDebugInfo_PJ.log"
    
        
    Open debugFileLog For Append As #1
    
    Print #1, ""
    Print #1, ""
    Print #1, "Start - " & CStr(Now)
    Print #1, oNamedItem.Name + "    " + CStr(oPlanningJointColl.Count) + " Planning Joints "
    Print #1, ""
    
    
    
    Dim i As Integer
    For i = 1 To oPlanningJointColl.Count
    
        Dim oPlanningJoint As IJPlnJoint
        Dim oPlanningJointProps As IJPlnJointProdProps
        Dim oPlanningJointProperties As IJPlnJointProps
                   
        On Error Resume Next
        Set oPlanningJoint = oPlanningJointColl.Item(i)
        
        Dim oTempPhysicalConn As IJStructPhysicalConnection
        Set oTempPhysicalConn = oPlanningJoint.GetPhysicalConnection
        
        If oPhysicalConn Is oTempPhysicalConn Then
    
            Dim oPJNamedItem As IJNamedItem
            Set oPJNamedItem = oPlanningJoint
            Print #1, " "
            Print #1, oPJNamedItem.Name
            
            Set oPlanningJointProps = oPlanningJoint
            Set oPlanningJointProperties = oPlanningJoint
            
            On Error GoTo ErrorHandler
            
            If Not oPlanningJointProps Is Nothing And Not oPlanningJointProperties Is Nothing Then
              
                Dim oSMSJointNode As IXMLDOMNode
                Set oSMSJointNode = oOutputXML.createNode(NODE_ELEMENT, "SMS_JOINT", "")
                        
                Dim oSMSJointElem As IXMLDOMElement
                Set oSMSJointElem = oSMSJointNode
                
                oSMSJointElem.setAttribute "NAME", oPJNamedItem.Name
                
                'Weld Position
                Dim weldPosition As JointWeldPosition
                weldPosition = oPlanningJointProperties.weldPosition
                Dim strWeldPosition As String: strWeldPosition = "NA"
                
                Select Case weldPosition
                    Case PLN_WeldPosition_Flat
                        strWeldPosition = "Flat"
                    Case PLN_WeldPosition_Horizontal
                        strWeldPosition = "Horizontal"
                    Case PLN_WeldPosition_Vertical
                        strWeldPosition = "Vertical"
                    Case PLN_WeldPosition_Overhead
                        strWeldPosition = "Overhead"
                    Case PLN_WeldPosition_NotApplicable
                        strWeldPosition = "NA"
                End Select
                    
                'Weld Direction
                Dim weldDirection As JointWeldDirection
                weldDirection = oPlanningJointProps.weldDirection
                Dim strWeldDirection As String: strWeldDirection = "NA"
                
                Select Case weldDirection
                    Case PLN_WeldDirection_Down
                        strWeldDirection = "Down"
                    
                    Case PLN_WeldDirection_Up
                        strWeldDirection = "Up"
                    
                    Case PLN_WeldDirection_NA
                        strWeldDirection = "NA"
                End Select
                
                oSMSJointElem.setAttribute "POSITION", strWeldPosition
                oSMSJointElem.setAttribute "DIRECTION", strWeldDirection
                oSMSJointElem.setAttribute "PASSES", CStr(oPlanningJointProps.WeldPasses)
                oSMSJointElem.setAttribute "DEFINITION", oPlanningJointProps.WeldClassification
                oSMSJointElem.setAttribute "WELD_MATERIAL_NAME", oPlanningJointProps.WeldMaterialName
                oSMSJointElem.setAttribute "WELD_MATERIAL_GRADE", oPlanningJointProps.WeldMaterialGrade
                oSMSJointElem.setAttribute "WELD_TESTING_TYPE", CStr(oPlanningJointProps.WeldTestingType)
                
                Dim oPlanningJointGeometry As IJPlnJointGeometry
                Set oPlanningJointGeometry = oPlanningJoint
                
                Dim oPlanningJointCurve As IJCurve
                Set oPlanningJointCurve = oPlanningJointGeometry.Curve
                    
                GenerateGeometryNode oOutputXML, oSMSJointNode, oPlanningJointCurve
       
                
    ''Dump Output
    '            On Error Resume Next
    '            Dim oModelBody As IJDModelBody
    '            Set oModelBody = oPlanningJointCurve
    '            If Not oModelBody Is Nothing Then
    '                FileName = Environ("TEMP")
    '                If FileName = "" Or FileName = vbNullString Then
    '                    FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
    '                End If
    '
    '                FileName = FileName & "\" + oNamedItem.Name + "_" + oPJNamedItem + ".sat"
    '
    '                oModelBody.DebugToSATFile FileName
    '                 MsgBox FileName + " dumped"
    '                Set oModelBody = Nothing
    '            End If
    '            On Error GoTo ErrorHandler
            
                
                Dim midX As Double, midY As Double, midZ As Double
                Dim startParam As Double, endParam As Double
                oPlanningJointCurve.ParamRange startParam, endParam
                oPlanningJointCurve.Position (startParam + endParam) / 2#, midX, midY, midZ
                Dim oMidPoint As New DPosition
                oMidPoint.Set midX, midY, midZ
                
                Print #1, "MiddlePoint : " & CStr(Format(midX, "0.00")) + ", " + CStr(Format(midY, "0.00")) + ", " + CStr(Format(midZ, "0.00"))
                
                'Find matching Pysical Connection Geometry
                Dim minDistTol As Double: minDistTol = 0.001
                Dim nMatchingPhysicalGeometryIndex As Integer
                Dim bOverlap As Boolean: bOverlap = False
                Dim j As Integer
                Dim distance As Double
                Dim srcX As Double, srcY As Double, srcZ As Double, inX As Double, inY As Double, inZ As Double
                
                For j = 1 To oPCGeometryColl.Count
                    Dim oPCGeometry As IJCurve
                    Set oPCGeometry = oPCGeometryColl.Item(j)
    
                    Print #1, CStr(j) + "/" + CStr(oPCGeometryColl.Count) + " PCGeometry"
    
                    Call oPCGeometry.DistanceBetween(oMidPoint, distance, srcX, srcY, srcZ, inX, inY, inZ)
                    
                    Print #1, "MinimumDistance : " & CStr(Format(distance, "0.0000"))
    
                    If distance < minDistTol Then
                        nMatchingPhysicalGeometryIndex = j
                        bOverlap = True
                        Print #1, "Matching PC Collection Index " + CStr(nMatchingPhysicalGeometryIndex)
                        Exit For
                    End If
        
                Next j
                    
                If bOverlap = False Then
                    
                    Print #1, "Try to check the middle point from PCGeometry to PlanningJoint : "
                    For j = 1 To oPCGeometryColl.Count
                        Dim oPCGeometry1 As IJCurve
                        Set oPCGeometry1 = oPCGeometryColl.Item(j)
                        
                        oPCGeometry1.ParamRange startParam, endParam
                        oPCGeometry1.Position (startParam + endParam) / 2#, midX, midY, midZ
                        
                        Dim oPCMidPoint As New DPosition
                        oPCMidPoint.Set midX, midY, midZ
        
                        Print #1, CStr(j) + "/" + CStr(oPCGeometryColl.Count) + " PCGeometry"
        
                        Call oPlanningJointCurve.DistanceBetween(oPCMidPoint, distance, srcX, srcY, srcZ, inX, inY, inZ)
                            
                        Print #1, "MinimumDistance : " & CStr(Format(distance, "0.00"))
    
                        If distance < minDistTol Then
                            nMatchingPhysicalGeometryIndex = j
                            bOverlap = True
                            Print #1, "Matching PC Collection Index " + CStr(nMatchingPhysicalGeometryIndex)
                            Exit For
                        End If
    
                    Next j
    
                End If
        
        
                If bOverlap Then
                    If oRefBevelColl.Count >= nMatchingPhysicalGeometryIndex Then
                        GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oRefPart, "REFPART", oRefVecColl.Item(nMatchingPhysicalGeometryIndex), oRefBevelColl.Item(nMatchingPhysicalGeometryIndex)
                    Else
                        GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oRefPart, "REFPART", oRefVecColl.Item(nMatchingPhysicalGeometryIndex), Nothing
                    End If
                    
                    If oNonRefBevelColl.Count >= nMatchingPhysicalGeometryIndex Then
                        GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oNonRefPart, "NONREFPART", oNonRefVecColl.Item(nMatchingPhysicalGeometryIndex), oNonRefBevelColl.Item(nMatchingPhysicalGeometryIndex)
                    Else
                        GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oNonRefPart, "NONREFPART", oNonRefVecColl.Item(nMatchingPhysicalGeometryIndex), Nothing
                    End If
                Else
                    Dim oPlanningJointNormalVec As IJDVector
                    Set oPlanningJointNormalVec = GetPlanningJointNormalVector(oPlanningJoint)
                    GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oRefPart, "REFPART", oPlanningJointNormalVec, Nothing
                    GeneratePartInfoNode oOutputXML, oSMSJointNode, oPhysicalConn, oNonRefPart, "NONREFPART", oNonRefVecColl.Item(1), Nothing
                End If
    
                oParentNode.appendChild oSMSJointNode
            End If
        End If
    Next
    
CleanUp:
    
    Set oSDPhysicalConn = Nothing
    Set oRefPart = Nothing
    Set oNonRefPart = Nothing
    Print #1, "End"
    Print #1, ""
    Close #1
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
    Dim oMemberPart As ISPSMemberPartPrismatic
    Dim oBean As IJBeam
    Dim oNamedItem As IJNamedItem
    Dim oStructMaterial As IJStructureMaterial
        
    Dim oSDProfile As StructDetailObjects.ProfilePart
    Set oSDProfile = New StructDetailObjects.ProfilePart
    
    Dim oSDMemberPart As StructDetailObjects.MemberPart
    Set oSDMemberPart = New StructDetailObjects.MemberPart
    
    
    Dim strPhysicalConnGuid As String

    Dim strPartGuid As String
    Dim strPartMaterial As String: strPartMaterial = ""
    Dim strPartGrade As String: strPartGrade = ""
    Dim strPartCategory As String: strPartCategory = ""

    If TypeOf oPart Is IJPlate Then
        Set oPlate = oPart
        dPartThickness = oPlate.thickness
        strPartCategory = GetShortDescription("StructPlateType", oPlate.plateType)
        Set oPlate = Nothing
    Else
    
        If TypeOf oPart Is IJStiffener Then
        Set oSDProfile.object = oPart
        dPartThickness = oSDProfile.WebThickness
        
            Set oStiffener = oPart
            strPartCategory = GetShortDescription("StructProfileType", oStiffener.pType)
            Set oStiffener = Nothing
            
        ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
            Set oSDMemberPart.object = oPart
            dPartThickness = oSDMemberPart.WebThickness
 
            Set oMemberPart = oPart
            strPartCategory = GetShortDescription("StructuralMemberType", oMemberPart.MemberType.Type)
            
            strPartMaterial = oMemberPart.MaterialDefinition.MaterialType
            strPartGrade = oMemberPart.MaterialDefinition.MaterialGrade

            Set oMemberPart = Nothing
    
        Else
            strPartCategory = ""
            dPartThickness = 0#
        End If
    End If
    
    On Error Resume Next
    Set oStructMaterial = oPart
    If Not oStructMaterial Is Nothing Then
        strPartMaterial = oStructMaterial.MaterialName
        strPartGrade = oStructMaterial.MaterialGrade
    End If
    Set oStructMaterial = Nothing

    On Error GoTo ErrorHandler

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
    ElseIf TypeOf oPart Is IJStiffener Then
        oPartInfoElem.setAttribute "PART_TYPE", "PROFILE"
    ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
        oPartInfoElem.setAttribute "PART_TYPE", "MEMBER"
    End If

    oPartInfoElem.setAttribute "THICKNESS", ConvertNumberToENLocale(dPartThickness)
    oPartInfoElem.setAttribute "CATEGORY", strPartCategory
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
    oCVGVectorElem.setAttribute "X", ConvertNumberToENLocale(dVectorX)
    oCVGVectorElem.setAttribute "Y", ConvertNumberToENLocale(dVectorY)
    oCVGVectorElem.setAttribute "Z", ConvertNumberToENLocale(dVectorZ)

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
        Dim NumOfBevel As Long: NumOfBevel = nCount
        If nCount = 12 Then
            NumOfBevel = nCount - 3 'Remove Chamfer
        End If
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
                If (NumOfBevel - i + 1) Mod 9 = 0 Then ' C
                    dC = dDepth
                    dC_Angle = dAngle
                ElseIf (NumOfBevel - i + 1) Mod 6 = 0 Then 'A
                    dA = dDepth
                    dA_Angle = dAngle
                ElseIf (NumOfBevel - i + 1) Mod 3 = 0 Then 'N
                    dN = dDepth
                    dN_Angle = dAngle
                Else
                    'No
                End If
            End If


        Next i
    End If
    
    oSMSBevelSideElem.setAttribute "PURPOSE", "MANUFACTURING"
    oSMSBevelSideElem.setAttribute "CHAMFER", ConvertNumberToENLocale(dChamfer)
    oSMSBevelSideElem.setAttribute "CHAMFER_ANGLE", ConvertNumberToENLocale(dChamfer_Angle)
    oSMSBevelSideElem.setAttribute "C", ConvertNumberToENLocale(dC)
    oSMSBevelSideElem.setAttribute "C_ANGLE", ConvertNumberToENLocale(dC_Angle)
    oSMSBevelSideElem.setAttribute "A", ConvertNumberToENLocale(dA)
    oSMSBevelSideElem.setAttribute "A_ANGLE", ConvertNumberToENLocale(dA_Angle)
    oSMSBevelSideElem.setAttribute "N", ConvertNumberToENLocale(dN)
    oSMSBevelSideElem.setAttribute "N_ANGLE", ConvertNumberToENLocale(dN_Angle)

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



Private Sub ChangePrimaryAndSecondarySide(ByRef oOutputXML As IXMLDOMDocument, strReferenceSide As String, ByRef oSMSConnectionNode As IXMLDOMNode)
Const METHOD = "ChangePrimaryAndSecondarySide"
On Error GoTo ErrorHandler
            
    Dim oSMSConnectionElement As IXMLDOMElement
    Set oSMSConnectionElement = oSMSConnectionNode
       
    'If referenceSide is molded, primary is referenceSide
    Dim bReferenceSideAsPrimarySide As Boolean
    If LCase(strReferenceSide) = "molded" Then
        bReferenceSideAsPrimarySide = True
     Else
        bReferenceSideAsPrimarySide = False
     End If
    
    
    'Primary Side
    Dim dPrimaryThroatThickness As Double
    Dim strPrimaryGrooveType As String
    Dim dPrimaryGrooveSize As Double
    Dim dPrimaryGrooveAngle As Double
    Dim dPrimarySideLength As Double
       
    dPrimaryThroatThickness = Val(oSMSConnectionNode.selectSingleNode("./@PRIMARY_THROAT_THICKNESS").Text)
    strPrimaryGrooveType = oSMSConnectionNode.selectSingleNode("./@PRIMARY_GROOVE_TYPE").Text
    dPrimaryGrooveSize = Val(oSMSConnectionNode.selectSingleNode("./@PRIMARY_GROOVE_SIZE").Text)
    dPrimaryGrooveAngle = Val(oSMSConnectionNode.selectSingleNode("./@PRIMARY_GROOVE_ANGLE").Text)
    dPrimarySideLength = Val(oSMSConnectionNode.selectSingleNode("./@PRIMARY_SIDE_LENGTH").Text)

    'Secondary Side
    Dim dSecondaryThroatThickness As Double
    Dim strSecondaryGrooveType As String
    Dim dSecondaryGrooveSize As Double
    Dim dSecondaryGrooveAngle As Double
    Dim dSecondarySideLength As Double
    
    dSecondaryThroatThickness = Val(oSMSConnectionNode.selectSingleNode("./@SECONDARY_THROAT_THICKNESS").Text)
    strSecondaryGrooveType = Val(oSMSConnectionNode.selectSingleNode("./@SECONDARY_GROOVE_TYPE").Text)
    dSecondaryGrooveSize = Val(oSMSConnectionNode.selectSingleNode("./@SECONDARY_GROOVE_SIZE").Text)
    dSecondaryGrooveAngle = Val(oSMSConnectionNode.selectSingleNode("./@SECONDARY_GROOVE_ANGLE").Text)
    dSecondarySideLength = Val(oSMSConnectionNode.selectSingleNode("./@SECONDARY_SIDE_LENGTH").Text)
   
    'THROAT_THICKNESS
    If bReferenceSideAsPrimarySide = True Then
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_THROAT_THICKNESS", "REF_THROAT_THICKNESS", ConvertNumberToENLocale(dPrimaryThroatThickness)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_THROAT_THICKNESS", "ANTIREF_THROAT_THICKNESS", ConvertNumberToENLocale(dSecondaryThroatThickness)
    Else
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_THROAT_THICKNESS", "REF_THROAT_THICKNESS", ConvertNumberToENLocale(dSecondaryThroatThickness)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_THROAT_THICKNESS", "ANTIREF_THROAT_THICKNESS", ConvertNumberToENLocale(dPrimaryThroatThickness)
        
    End If
    
    'GROOVE_TYPE
     If bReferenceSideAsPrimarySide = True Then
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_GROOVE_TYPE", "REF_GROOVE_TYPE", strPrimaryGrooveType
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_GROOVE_TYPE", "ANTIREF_GROOVE_TYPE", strSecondaryGrooveType
    Else
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_GROOVE_TYPE", "REF_GROOVE_TYPE", strSecondaryGrooveType
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_GROOVE_TYPE", "ANTIREF_GROOVE_TYPE", strPrimaryGrooveType
    End If
        
    'GROOVE_SIZE
     If bReferenceSideAsPrimarySide = True Then
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_GROOVE_SIZE", "REF_GROOVE_SIZE", ConvertNumberToENLocale(dPrimaryGrooveSize)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_GROOVE_SIZE", "ANTIREF_GROOVE_SIZE", ConvertNumberToENLocale(dSecondaryGrooveSize)
    Else
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_GROOVE_SIZE", "REF_GROOVE_SIZE", ConvertNumberToENLocale(dSecondaryGrooveSize)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_GROOVE_SIZE", "ANTIREF_GROOVE_SIZE", ConvertNumberToENLocale(dPrimaryGrooveSize)
    End If
    
    'GROOVE_ANGLE
     If bReferenceSideAsPrimarySide = True Then
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_GROOVE_ANGLE", "REF_GROOVE_ANGLE", ConvertNumberToENLocale(dPrimaryGrooveAngle)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_GROOVE_ANGLE", "ANTIREF_GROOVE_ANGLE", ConvertNumberToENLocale(dSecondaryGrooveAngle)
        
    Else
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_GROOVE_ANGLE", "REF_GROOVE_ANGLE", ConvertNumberToENLocale(dSecondaryGrooveAngle)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_GROOVE_ANGLE", "ANTIREF_GROOVE_ANGLE", ConvertNumberToENLocale(dPrimaryGrooveAngle)
    End If

    'SIDE_LENGTH
     If bReferenceSideAsPrimarySide = True Then
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_SIDE_LENGTH", "REF_SIDE_LENGTH", ConvertNumberToENLocale(dPrimarySideLength)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_SIDE_LENGTH", "ANTIREF_SIDE_LENGTH", ConvertNumberToENLocale(dSecondarySideLength)
    Else
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "PRIMARY_SIDE_LENGTH", "REF_SIDE_LENGTH", ConvertNumberToENLocale(dSecondarySideLength)
        ReplaceAttributeNode oOutputXML, oSMSConnectionNode, "SECONDARY_SIDE_LENGTH", "ANTIREF_SIDE_LENGTH", ConvertNumberToENLocale(dPrimarySideLength)
    End If

    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Private Sub ReplaceAttributeNode(ByRef oOutputXML As IXMLDOMDocument, ByRef parentElement As IXMLDOMElement, strOldAttributeName As String, strNewAttributeName As String, varNewValue As Variant)
Const METHOD = "ReplaceAttributeNode"
On Error GoTo ErrorHandler

    parentElement.removeAttribute (strOldAttributeName)
    parentElement.setAttribute strNewAttributeName, varNewValue
    
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Private Function ConvertNumberToENLocale(ByVal value As Double) As String
Const METHOD = "ConvertNumberToEngishLocale"
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


Private Function GetPhysicalConnection(ByVal oPlanningJoints As IJElements) As IJElements
Const METHOD = "GetPhysicalConnection"
On Error GoTo ErrorHandler

    Dim oPhysicalConnections As IJElements
    Set oPhysicalConnections = New JObjectCollection
    
    Dim i As Integer
    For i = 1 To oPlanningJoints.Count
    
        Dim oPlanningJoint As IJPlnJoint
        Dim oPhysicalConnection As Object
        
        On Error Resume Next
        Set oPlanningJoint = oPlanningJoints.Item(i)
        If Not oPlanningJoint Is Nothing Then
            Set oPhysicalConnection = oPlanningJoint.GetPhysicalConnection
            
            If oPhysicalConnections.Contains(oPhysicalConnection) = False Then
                oPhysicalConnections.Add oPhysicalConnection
            End If
        End If
        
    Next i
    
    Set GetPhysicalConnection = oPhysicalConnections

Exit Function
ErrorHandler:
    Resume Next
End Function


Private Function GetPlanningJointNormalVector(oPlnJoint As IJPlnJoint) As IJDVector
    Const METHOD = "GetPlanningJointNormalVector"
    On Error GoTo ErrorHandler

    Dim oPlnRelations As ShipService.PlanningRelations
    Dim oPlnJointColl As IJDObjectCollection
    Dim oCurve As IJCurve
    Dim oMidPoint As IJDPosition  'Mid Point of the selected Planning Joint
    Dim oTanVector As IJDVector   'Tangent vector at the Mid Point
    Dim NormalX As Double
    Dim NormalY As Double
    Dim NormalZ As Double
    Dim oNormalVector As IJDVector
    Dim oLine As IJLine
    Dim oPlane As IJPlane
    Dim oMidPoint1 As IJDPosition
    Dim Code As Geom3dIntersectConstants
    Dim lNoOfIntersects As Long
    Dim Points() As Double
    Dim lNoOfOverlaps As Long
    Dim i As Long, lCount As Long
    Dim dDist1 As Double
    Dim dDist2 As Double
    Dim lX As Double
    Dim lY As Double
    Dim lZ As Double
    
    Set oPlnRelations = New ShipService.PlanningRelations
    
    Dim oConnection As Object
    Set oConnection = oPlnRelations.GetFirstTargetObject(oPlnJoint, IID_IJPlnJoint, "PlnJointPhysConn_DEST")
    
    ' When PC(or related seam) is deleted with PJ in read only, read only PJs will be left with no PC.
    ' In the above case GetTargetObjectCollection raises error, so check for connection
    If Not oConnection Is Nothing Then
        Dim oTargetObjectColl As IJDTargetObjectCol
        Set oTargetObjectColl = oPlnRelations.GetTargetObjectCollection(oConnection, _
                                                            IID_IJStructGeometry, _
                                                            "PlnJointPhysConn_ORIG")
                                                        
        If oTargetObjectColl Is Nothing Then GoTo CleanUp
    Else
        GoTo CleanUp
    End If
    
    'Get all(also the Split Planning Joints) the PlnJoints associated with the connection
    Set oPlnJointColl = New JObjectCollection
    
    For i = 1 To oTargetObjectColl.Count
        oPlnJointColl.Add oTargetObjectColl.Item(i)
    Next
    
    'Get mid point and normal to he selected planning joint
    Set oCurve = oPlnJoint
    Set oTanVector = New DVector
    Set oMidPoint = New DPosition
    
    Call GetMidPointAndTangentVectorOfCurve(oCurve, oMidPoint, oTanVector)
    
    Set oLine = New Line3d
    Set oPlane = New Plane3d
    
    'Create a plane with the mid point of the planning joint as the Root Point and the tangent as the normal to the plane.
    oPlane.DefineByPointNormal oMidPoint.x, oMidPoint.y, oMidPoint.z, oTanVector.x, oTanVector.y, oTanVector.z
        
    Set oMidPoint1 = New DPosition
    lNoOfIntersects = 0
    lCount = 1
    
    For Each oPlnJoint In oPlnJointColl
        Set oCurve = oPlnJoint
        Call GetMidPointAndTangentVectorOfCurve(oCurve, oMidPoint1, oTanVector)
        'verify that the Planning Joint is not the one selected
        If oMidPoint.x <> oMidPoint1.x Or oMidPoint.y <> oMidPoint1.y Or oMidPoint.z <> oMidPoint1.z Then
            'Check if the plane intersects one of the planning joints
            oCurve.Intersect oPlane, lNoOfIntersects, Points, lNoOfOverlaps, Code
            If lNoOfIntersects > 0 Then
                dDist2 = Sqr((oMidPoint.x - Points(0)) * (oMidPoint.x - Points(0)) + (oMidPoint.y - Points(1)) * (oMidPoint.y - Points(1)) + (oMidPoint.z - Points(2)) * (oMidPoint.z - Points(2)))
                'Get the Planning JOint which is closest to the selected Planning Joint
                If dDist2 < dDist1 Or lCount = 1 Then
                    dDist1 = dDist2
                    lX = Points(0)
                    lY = Points(1)
                    lZ = Points(2)
                    lCount = lCount + 1
                End If
            End If

        End If
        Set oCurve = Nothing
    Next
    
    'Get the direction of the line joining the mid point of the selected
    'Planning Joint and the closest point on the nearest planning joint
    oLine.DefineBy2Points oMidPoint.x, oMidPoint.y, oMidPoint.z, lX, lY, lZ
    oLine.GetDirection NormalX, NormalY, NormalZ
    
    'Display the Direction of the Vector in a direction opposite to that of the line
    Set oNormalVector = New DVector
    oNormalVector.x = -NormalX
    oNormalVector.y = -NormalY
    oNormalVector.z = -NormalZ
    

    Set GetPlanningJointNormalVector = oNormalVector
           
CleanUp:
    Set oMidPoint = Nothing
    Set oMidPoint1 = Nothing
    Set oTanVector = Nothing
    Set oNormalVector = Nothing
    Set oPlnRelations = Nothing
    Set oLine = Nothing
    Set oPlane = Nothing
    
    Exit Function
    
ErrorHandler:
    GoTo CleanUp
End Function


Private Sub GetMidPointAndTangentVectorOfCurve(oCurve As IJCurve, oMidPoint As IJDPosition, oTangentVector As IJDVector)
    Const METHOD = "GetTangentVectorAtMidPointOfCurve"
    On Error GoTo ErrorHandler
    
    Dim dStartPar As Double
    Dim dEndPar As Double
    Dim dMidPar As Double
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    Dim dTanx As Double
    Dim dTany As Double
    Dim dTanz As Double
    Dim dTan2x As Double
    Dim dTan2y As Double
    Dim dTan2z As Double
    
    'Get the Start and End Parameters of Curve
    oCurve.ParamRange dStartPar, dEndPar
       
    'Get the Parameter for the MidPoint
    dMidPar = (dStartPar + dEndPar) / 2
    
    'Get the MidPoint and Tangent at the Midpoint on the Curve
    oCurve.Evaluate dMidPar, dX, dY, dZ, dTanx, dTany, dTanz, dTan2x, dTan2y, dTan2z
    
    'Set the MidPoint
    oMidPoint.Set dX, dY, dZ
    'Set the Tangent Vector at the Midpoint of curve
    oTangentVector.Set dTanx, dTany, dTanz
    
    Exit Sub
ErrorHandler:
    Resume Next
End Sub



