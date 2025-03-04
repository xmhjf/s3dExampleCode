VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Generic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const PI As Double = 3.14159265358979
Implements IJMfgXMLData
Implements IJMfgXMLReport
Private Const ROUND_CONST = 1
'Private Const IID_IJMfgOutput As String = "{08E0F9B1-41F0-4279-8EBA-AE37DD7389CD}"
Private Const E_ACCESSDENIED As Long = -2147467259


Private Const TKWorkingSet = "WorkingSet"

Private Sub IJMfgXMLReport_GenerateReport(ByVal pMfgObj As Object, ByVal pXMLDOMDoc As Object, ByVal lFaceId As Long, pReportDoc As Object)
    Const METHOD = "IJMfgXMLData_GenerateReport"
    On Error GoTo ErrorHandler
    
    Dim oMfgProfileReport As IJMfgProfileReport
    Set oMfgProfileReport = pMfgObj
    If oMfgProfileReport Is Nothing Then
        Exit Sub
    End If
    Dim oReportXML As DOMDocument
    Dim lCounter As Long
    Dim iCount               As Long
    Dim strMsg As String
    Dim dProfLength As Double
    Set oReportXML = New DOMDocument
    Dim oPositionElement As IXMLDOMElement
    Dim oReportElement As IXMLDOMElement
    Dim oReportParent As IXMLDOMElement
    Dim dHeight As Double
    Dim dGirth As Double
    Dim oMfgProfilePart As IJMfgProfilePart
    
    '' *************************************************************************
    ''  TEST : GetCurveHeightByOffset -- STRMFG_BENDING_LINE
    '' *************************************************************************
    Dim Offset As Double
    Dim dMargin As Double
    Dim bReverseDir As Boolean
    Set oMfgProfilePart = pMfgObj
    
    If Not oMfgProfilePart Is Nothing Then
        Offset = oMfgProfilePart.IntervalDistance '0.5
        dMargin = oMfgProfilePart.BottomMargin
        bReverseDir = oMfgProfilePart.ReverseIBLDirection
    Else
        Offset = 0.5
    End If
    Dim nIBLCurveCount As Long
    Set oReportParent = oReportXML.createElement("SMS_REPORT")
    oReportXML.appendChild oReportParent
    dProfLength = oMfgProfileReport.GetProcessingLength(STRMFG_AfterFeaturesLength, lFaceId)
    nIBLCurveCount = oMfgProfileReport.GetCurveCount(STRMFG_BENDING_LINE, lFaceId)
    

    If nIBLCurveCount > 0 Then
        lCounter = 0
        Dim nIndex As Long
        
        'write each curve in one report element
        For nIndex = 1 To nIBLCurveCount
            Set oReportElement = oReportXML.createElement("SMS_REPORT_GEOMETRY")
            oReportElement.setAttribute "TYPE", "strmfg_bending_line"
            oReportElement.setAttribute "PROCESS", "2D"
            oReportElement.setAttribute "GEOM2D_GUID", ""   'MISSING
            oReportElement.setAttribute "CURVE_TYPE", "curve"
            oReportElement.setAttribute "LENGTH", ""        'MISSING
            Select Case lFaceId
                Case JXSEC_WEB_LEFT
                    oReportElement.setAttribute "FACE", "webleft"
                Case JXSEC_WEB_RIGHT
                    oReportElement.setAttribute "FACE", "webright"
                Case JXSEC_TOP
                    oReportElement.setAttribute "FACE", "top_flange"
                Case JXSEC_BOTTOM
                    oReportElement.setAttribute "FACE", "bottom_flange"
            End Select
            oReportElement.setAttribute "FACE_HEIGHT", ""   'MISSING
            
            oReportParent.appendChild oReportElement
            strMsg = nIndex & vbTab
            lCounter = 0
            
            Offset = oMfgProfileReport.GetMarkingCurveAttribute(STRMFG_BENDING_LINE, lFaceId, nIndex, Nothing, Offset, "INTERVAL_DISTANCE")
            
            Dim dStartX As Double, dStartY As Double, dEndX As Double, dEndY As Double
            oMfgProfileReport.GetCurveHeightAtEnds STRMFG_BENDING_LINE, lFaceId, nIndex, dStartX, dStartY, dEndX, dEndY
            Dim dStartXOut As Double, dStartYOut As Double, dEndXOut As Double, dEndYOut As Double
            If bReverseDir Then
                dStartXOut = dEndX
                dEndXOut = dStartX
                dStartYOut = dEndY
                dEndYOut = dStartY
            Else
                dStartXOut = dStartX
                dEndXOut = dEndX
                dStartYOut = dStartY
                dEndYOut = dEndY
            End If
            Set oPositionElement = oReportXML.createElement("SMS_POSITION")
            oPositionElement.setAttribute "X", Round(1000 * dStartXOut, ROUND_CONST)
            oPositionElement.setAttribute "Y", Round(1000 * dStartYOut, ROUND_CONST)
            oPositionElement.setAttribute "Z", 0
            oPositionElement.setAttribute "UNIT", "mm"
            On Error Resume Next
            dGirth = oMfgProfileReport.GetSupportValue(STRMFG_PROFILE_LENGTH, lFaceId, dStartXOut)
            On Error GoTo ErrorHandler
            oPositionElement.setAttribute "GIRTH_LENGTH", Round(1000 * dGirth, ROUND_CONST)
            oPositionElement.setAttribute "REF_LENGTH", ""      'MISSING
            oPositionElement.setAttribute "REF_HEIGHT", Round(1000 * dStartYOut, ROUND_CONST)
            oPositionElement.setAttribute "POSITION", "start_point"
            oReportElement.appendChild oPositionElement
            
            lCounter = lCounter + 1
            
            Do
                Dim dXPos As Double
                If bReverseDir Then
                    dXPos = (dEndX - lCounter * Offset)
                Else
                    dXPos = lCounter * Offset
                End If
                dHeight = Round(1000 * oMfgProfileReport.GetCurveHeightByOffset(STRMFG_BENDING_LINE, lFaceId, nIndex, dXPos), ROUND_CONST)
                If (dHeight > dMargin * 1000) And (dXPos > 0#) Then
                    Set oPositionElement = oReportXML.createElement("SMS_POSITION")
                    oPositionElement.setAttribute "X", Round(1000 * dXPos, ROUND_CONST)
                    oPositionElement.setAttribute "Y", dHeight
                    oPositionElement.setAttribute "Z", 0
                    oPositionElement.setAttribute "UNIT", "mm"
                    dGirth = oMfgProfileReport.GetSupportValue(STRMFG_PROFILE_LENGTH, lFaceId, dXPos)
                    oPositionElement.setAttribute "GIRTH_LENGTH", Round(1000 * dGirth, ROUND_CONST)
                    oPositionElement.setAttribute "REF_LENGTH", ""      'MISSING
                    oPositionElement.setAttribute "REF_HEIGHT", dHeight
                    oPositionElement.setAttribute "POSITION", "internal"
                    oReportElement.appendChild oPositionElement
                End If
                dGirth = 0
                lCounter = lCounter + 1
                Dim bContinue As Boolean
                If bReverseDir Then
                    bContinue = dXPos > dStartX
                Else
                    bContinue = dXPos < dEndX
                End If
            Loop While bContinue
            
            If (dEndXOut < 0) Then dEndXOut = 0#
        
            Set oPositionElement = oReportXML.createElement("SMS_POSITION")
            oPositionElement.setAttribute "X", Round(1000 * dEndXOut, ROUND_CONST)
            oPositionElement.setAttribute "Y", Round(1000 * dEndYOut, ROUND_CONST)
            oPositionElement.setAttribute "Z", 0
            oPositionElement.setAttribute "UNIT", "mm"
            On Error Resume Next
            dGirth = oMfgProfileReport.GetSupportValue(STRMFG_PROFILE_LENGTH, lFaceId, dEndXOut)
            On Error GoTo ErrorHandler
            oPositionElement.setAttribute "GIRTH_LENGTH", Round(1000 * dGirth, ROUND_CONST)
            oPositionElement.setAttribute "REF_LENGTH", ""      'MISSING
            oPositionElement.setAttribute "REF_HEIGHT", Round(1000 * dEndYOut, ROUND_CONST)
            oPositionElement.setAttribute "POSITION", "end_point"
            oReportElement.appendChild oPositionElement
        
        
        Next
    End If
        'Dim FileName As String
        'FileName = Environ("TEMP")
        'If FileName = "" Or FileName = vbNullString Then
        '       FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
        'End If
        'FileName = FileName & "\profilereportxml.xml"
    'oReportXML.save FileName
    Set pReportDoc = oReportXML
Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints
' Purpose   : Given SMS_EDGE node with geometry, return new SMS_EDGE node with Points
'---------------------------------------------------------------------------------------
'
Private Function ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oInputEdgeNode As IXMLDOMNode) As IXMLDOMNode
    Const METHOD = "ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints"
    On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler

    Dim oReturnEdgeNode As IXMLDOMNode
    Set oReturnEdgeNode = oInputEdgeNode.cloneNode(True)

    Set ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints = oReturnEdgeNode

    Dim oMfgXMLhelper As New MfgXMLHelper

    Dim oPointColl As IJElements
    ' Chord Height of 0.1 mm is recommended.  If customer's "30mm" is absolute must,
    ' replace 2nd and 3rd ags with MfgStroke_ArcLength and 0.03 respectively
    Set oPointColl = oMfgXMLhelper.StrokeEdgeNodeIntoPoints(oInputEdgeNode, _
                                                            MfgStroke_ChordHeight, _
                                                            0.0001)

    If oPointColl Is Nothing Then Exit Function
    If oPointColl.Count = 0 Then Exit Function

    Dim oChildCVGnode As IXMLDOMNode
    For Each oChildCVGnode In oReturnEdgeNode.childNodes
        If oChildCVGnode.baseName = "CVG_CURVE" Then
            oReturnEdgeNode.removeChild oChildCVGnode
        End If
    Next

    Dim oCreator As IXMLDOMDocument
    Set oCreator = oInputEdgeNode.ownerDocument

    Dim oNewElem As IXMLDOMElement
    Set oNewElem = oCreator.createElement("CVG_CURVE")

    Set oChildCVGnode = oReturnEdgeNode.appendChild(oNewElem)

    Dim i As Long
    For i = 1 To oPointColl.Count
        Dim oPos As IJDPosition
        Set oPos = oPointColl.Item(i)

        Set oNewElem = oCreator.createElement("CVG_VERTEX")
        If i = 1 Then
            oNewElem.setAttribute "POINT_CODE", "s_point"
        ElseIf i = oPointColl.Count Then
            oNewElem.setAttribute "POINT_CODE", "e_point"
        Else
            oNewElem.setAttribute "POINT_CODE", "node"
        End If
        oNewElem.setAttribute "SEG_TYPE", "point"

        oNewElem.setAttribute "SX", oPos.x * 1000#
        oNewElem.setAttribute "SY", oPos.y * 1000#

        oNewElem.setAttribute "A", 0#
        oNewElem.setAttribute "B", 0#
        oNewElem.setAttribute "C", 0#

        oChildCVGnode.appendChild oNewElem
    Next

    Set ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints = oReturnEdgeNode

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, METHOD
End Function


Private Sub IJMfgXMLData_Update(ByVal pMfgObj As Object, ByVal pXMLDOMDoc As Object, ByVal strOutputOption As String)
    Const METHOD = "IJMfgXMLData_Update"
    On Error GoTo ErrorHandler
    
    Dim oIJDObject As IJDObject
    Set oIJDObject = pMfgObj
    Dim oPOM As IJDPOM
    Set oPOM = oIJDObject.ResourceManager
    
    Dim lCreateNestData As Long
    lCreateNestData = GetRuleValue(strOutputOption, "Long4")
    
    Dim oNodeList As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim oDomDoc As DOMDocument
    Set oDomDoc = pXMLDOMDoc
    If lCreateNestData > 0 Then
        Dim oOutput As IJMfgOutput
        Set oOutput = pMfgObj
        Set oNodeList = oDomDoc.selectNodes("//SMS_PLATE | //SMS_PROFILE")
        
        Dim oNestData As IJMfgNestData
        For Each oNode In oNodeList
            Dim strNestGuid As String
            strNestGuid = oNode.selectSingleNode(".//SMS_PART_INFO/@NEST_GUID").nodeValue
            If strNestGuid = "" Then
                Dim strPartType As String
                strPartType = GetPartTypeString(oNode)
                Dim strModelGuid As String
                Dim oDetailMoniker As IMoniker
                Dim oDetailPart As Object
                strModelGuid = oNode.selectSingleNode(".//SMS_PART_INFO/@MODEL_PART_GUID").nodeValue
                Set oDetailMoniker = oPOM.MonikerFromDbIdentifier("{" & strModelGuid & "}")
                Set oDetailPart = oPOM.GetObject(oDetailMoniker)
                Set oNestData = oOutput.GetMfgNestData(oDetailPart, strPartType, True, strOutputOption)
                Dim oNestMoniker As IMoniker
                Set oNestMoniker = oPOM.GetObjectMoniker(oNestData)
                strNestGuid = oPOM.DbIdentifierFromMoniker(oNestMoniker)
                strNestGuid = Replace(strNestGuid, "{", "")
                strNestGuid = Replace(strNestGuid, "}", "")
                oNode.selectSingleNode(".//SMS_PART_INFO/@NEST_GUID").nodeValue = strNestGuid
                oNestData.OutputType = strOutputOption
            End If
        Next oNode
    Else
        If Not oDomDoc Is Nothing Then
            Set oNodeList = oDomDoc.selectNodes("//SMS_PART_INFO")
            For Each oNode In oNodeList
                Dim oElement As IXMLDOMElement
                Set oElement = oNode
                oElement.setAttribute "SMS_OUTPUT_FORMAT", strOutputOption
            Next oNode
        End If
    End If

    ' CR 2944781: Stroke IBLs into points - uncomment below 5 lines and recompile if required
''    Set oNodeList = oDomDoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bending_line']/SMS_EDGE")
''    For Each oNode In oNodeList
''        oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode)
''        oNode.parentNode.removeChild oNode
''    Next
''    Set oNodeList = oDomDoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_topflange_bending_line']/SMS_EDGE")
''    For Each oNode In oNodeList
''        oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode)
''        oNode.parentNode.removeChild oNode
''    Next
''    Set oNodeList = oDomDoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bottomflange_bending_line']/SMS_EDGE")
''    For Each oNode In oNodeList
''        oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode)
''        oNode.parentNode.removeChild oNode
''    Next

Cleanup:
    Set oNestData = Nothing
    Set pMfgObj = Nothing
    Set oIJDObject = Nothing
    Set oPOM = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
End Sub

Private Function GetPartTypeString(oNode As IXMLDOMNode) As String
    Const METHOD = "GetPartTypeString"
    On Error GoTo ErrorHandler
    
    Dim strTempType As String
    strTempType = oNode.selectSingleNode(".//SMS_PART_INFO/@PART_TYPE").nodeValue
    If strTempType = "PLATE" Then
        GetPartTypeString = "PLATE"
    Else
        Dim strTempCellName As String
        Dim oCellNode As IXMLDOMNode
        Set oCellNode = oNode.selectSingleNode(".//SMS_PART_INFO/@CELL_NAME")
        If oCellNode Is Nothing Then
            If strTempType = "SEC" Then
                GetPartTypeString = "PROFILE,SEC"
            Else
                GetPartTypeString = "PLATE," & strTempType
            End If
        Else
            strTempCellName = oCellNode.nodeValue
            If strTempType = "SEC" Then
                GetPartTypeString = "PROFILE," & strTempCellName
            Else
                GetPartTypeString = "PLATE," & strTempCellName
            End If
        End If
    End If

Cleanup:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
End Function

Private Function GetRuleValue(strRuleName As String, strColumn As String) As Variant
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
Cleanup:
    Set oOutputHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
    GoTo Cleanup
End Function
