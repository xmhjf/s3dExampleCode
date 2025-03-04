VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DXF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJMfgOutputDocument

Private m_oControlInputDoc As IXMLDOMDocument
Private m_bstrOutputOption As String
Private m_strSharedContentPath As String

Private Sub Class_Initialize()
    Set m_oControlInputDoc = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_oControlInputDoc = Nothing
End Sub

Private Function IJMfgOutputDocument_Generate(ByVal pObject As Object, ByVal pOutputXMLObject As Object, ByVal bstrBaseOutputName As String) As Object
Const METHOD = "IJMfgOutputDocument_Generate"
On Error GoTo ErrorHandler

    Dim oDXFService As IJStrMfgDXFService
    Set oDXFService = New StrMfgDXFServiceLib.CStrMfgDXFService


    Dim oStyleService As IJStrMfgStyleService
    Set oStyleService = New StrMfgStyleServiceLib.CStrMfgStyleService
    
    'Initialize the StyleService
    Dim strStyleXMLPath As String
    strStyleXMLPath = m_strSharedContentPath + "\" + GetRuleValue(m_bstrOutputOption, "String3")
    oStyleService.Initialize strStyleXMLPath

    Dim strExtension As String
    strExtension = GetRuleValue(m_bstrOutputOption, "DocumentType")


    'Initialize Default StyleService values
    'MsgBox ("Set default style values")
    'oStyleService.SetDefaultStyle "LAYER", "Default"
    'oStyleService.SetDefaultStyle "COLOR", 0
    'oStyleService.SetDefaultStyle "LINETYPE", ""
    'oStyleService.SetDefaultStyle "LEVEL", 1
    'oStyleService.SetDefaultStyle "STYLE", 1
    'oStyleService.SetDefaultStyle "LINEWEIGHT", 1
    'oStyleService.SetDefaultStyle "TEXT_HEIGHT", 40#
    'oStyleService.SetDefaultStyle "TEXT_ANGLE", 0#
    'oStyleService.SetDefaultStyle "TEXT_JUSTIFY", 0
    'MsgBox ("Successfully set default style values")

    Dim oAttachmentXMLDoc As IXMLDOMDocument
    Dim oS3DOutputRootNode As IXMLDOMNode
    Dim oXMLDOC As IXMLDOMDocument


    Set oXMLDOC = pOutputXMLObject

    Dim strFullFileName As String
    Dim strPartID As String

    'Generate Output ReportDocument
    Dim oOutputReportDoc As New DOMDocument

    Dim oOutputReportNode As IXMLDOMNode

    Set oOutputReportNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputReport", "")

    oOutputReportDoc.appendChild oOutputReportNode

    If Not oXMLDOC Is Nothing Then

        Dim oPlateOrProfileNodeList As IXMLDOMNodeList
        Dim oPlateOrProfileNode As IXMLDOMNode
        Set oPlateOrProfileNodeList = oXMLDOC.selectNodes("//SMS_SMS/SMS_PLATES/SMS_PLATE|//SMS_SMS/SMS_PROFILES/SMS_PROFILE")
        


        For Each oPlateOrProfileNode In oPlateOrProfileNodeList
            Dim oNodeList As IXMLDOMNodeList
            Dim oNode As IXMLDOMNode


            GetDXFFileNameAndPartID pObject, oPlateOrProfileNode, bstrBaseOutputName, strExtension, strFullFileName, strPartID

            'Write the beginning of this Plate's/Profile's DXF file.
            oDXFService.WriteHeaderData oStyleService, strFullFileName


'            'Write the Product Information label.
            Set oNodeList = oStyleService.GetLabel(oPlateOrProfileNode)
            If Not oNodeList Is Nothing Then
                oDXFService.WritePartIdentificationData oNodeList
            End If
'
            'Obtain the contour, annotation, and marking nodes for the current plate or profile.
            If oPlateOrProfileNode.baseName = "SMS_PLATE" Then
                Set oNodeList = oPlateOrProfileNode.selectNodes("./*[not(@BLOCK_NAME)]")
            ElseIf oPlateOrProfileNode.baseName = "SMS_PROFILE" Then
                Set oNodeList = oPlateOrProfileNode.selectNodes("./SMS_PROFILE_2D/*")
            End If

            For Each oNode In oNodeList
                Dim oTypeNodes As IXMLDOMNodeList
                Dim oTypeNode As IXMLDOMNode

                'Handle a Profile's origin points seperately.
                If oNode.baseName = "SMS_ORIGIN_POINTS" Then
                    'Set oTypeNodes = oNode.childNodes
                    'Handle the SMS_ORIGIN_POINTs for Profiles...
                    Set oTypeNodes = oXMLDOC.selectNodes("//BLEHBLEHBLEH")
                Else
                    'Obtain contour, annotation, or marking nodes on a list.
                    If oNode.baseName <> "SMS_PLATE_CONTOUR" And oNode.baseName <> "SMS_FACE_CONTOUR" Then
                        Set oTypeNodes = oNode.childNodes
                    ElseIf oPlateOrProfileNode.baseName = "SMS_PLATE" Then
                        Set oTypeNodes = oPlateOrProfileNode.selectNodes(".//SMS_PLATE_CONTOUR")
                    Else
                        Set oTypeNodes = oPlateOrProfileNode.selectNodes(".//SMS_FACE_CONTOUR")
                    End If

                    'Write DXF for all CVG_TEXT and SMS_EDGE nodes for the current contour/annotation/marking node
                    For Each oTypeNode In oTypeNodes
                        Dim oTextNodeList As IXMLDOMNodeList
                        Dim oEdgeNodeList As IXMLDOMNodeList
                        Dim oTextNode As IXMLDOMNode
                        Dim oEdgeNode As IXMLDOMNode
                        Dim oStyleNode As IXMLDOMNode

                        'Obtain a list of all CVG_TEXT nodes and write dxf for each
                        Set oTextNodeList = oTypeNode.selectNodes("./CVG_TEXT")
                        For Each oTextNode In oTextNodeList
                            Set oStyleNode = oStyleService.GetStyle(oTextNode)
                            oDXFService.WriteCVGTextData oTextNode, oStyleNode
                        Next oTextNode

                        'Obtain a list of all SMS_EDGE nodes and write dxf for each
                        Set oEdgeNodeList = oTypeNode.selectNodes("./SMS_EDGE")
                        For Each oEdgeNode In oEdgeNodeList
                            Set oStyleNode = oStyleService.GetStyle(oEdgeNode)
                            oDXFService.WriteCVGCurveData oEdgeNode, oStyleNode

                            'If the SMS_EDGE node has an Annotation, write the DXF for its TEXT and EDGE nodes
                            Dim oAnnotationNodeList As IXMLDOMNodeList
                            Dim oEdgeAnnotNodeList As IXMLDOMNodeList
                            Dim oTextAnnotNodeList As IXMLDOMNodeList
                            Dim oAnnotationNode As IXMLDOMNode
                            Dim oEdgeAnnotNode As IXMLDOMNode
                            Dim oTextAnnotNode As IXMLDOMNode

                            Set oAnnotationNodeList = oEdgeNode.selectNodes("./SMS_ANNOTATION")
                            If Not oAnnotationNodeList Is Nothing Then
                                For Each oAnnotationNode In oAnnotationNodeList

                                    'Obtain and write dxf for all CVG_TEXT nodes
                                    Set oTextAnnotNodeList = oAnnotationNode.selectNodes("./CVG_TEXT")
                                    For Each oTextAnnotNode In oTextAnnotNodeList
                                        Set oStyleNode = oStyleService.GetStyle(oTextAnnotNode)
                                        oDXFService.WriteCVGTextData oTextAnnotNode, oStyleNode
                                    Next oTextAnnotNode

                                    'Obtain and write dxf for all SMS_EDGE nodes
                                    Set oEdgeAnnotNodeList = oAnnotationNode.selectNodes("./SMS_EDGE")
                                    For Each oEdgeAnnotNode In oEdgeAnnotNodeList
                                        Set oStyleNode = oStyleService.GetStyle(oEdgeAnnotNode)
                                        oDXFService.WriteCVGCurveData oEdgeAnnotNode, oStyleNode
                                    Next oEdgeAnnotNode

                                Next oAnnotationNode
                            End If
                        Next oEdgeNode
                    Next oTypeNode
                End If
            Next oNode

            'Write the end of this Plate's/Profile's DXF file and close file.
            oDXFService.WriteTrailerData

            Dim oOutputDocumentNode As IXMLDOMNode
            Set oOutputDocumentNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputDocument", "")

            Dim oOutputDocumentElem As IXMLDOMElement
            Set oOutputDocumentElem = oOutputDocumentNode
            oOutputDocumentElem.setAttribute "ID", strPartID

            Dim lastPos As Long
            lastPos = InStrRev(strFullFileName, "\")
            Dim strFileName As String
            strFileName = Trim(Right$(strFullFileName, Len(strFullFileName) - lastPos))
            oOutputDocumentElem.setAttribute "NAME", strFileName

            oOutputReportNode.appendChild oOutputDocumentNode

        Next oPlateOrProfileNode
    End If


    Set IJMfgOutputDocument_Generate = oOutputReportDoc

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function
Private Sub IJMfgOutputDocument_Initialize(ByVal pControlXMLObject As Object, ByVal bstrOutputOption As String, ByVal bstrSharedContentsPath As String)
Const METHOD = "IJMfgOutputDocument_Initialize"
On Error GoTo ErrorHandler

    m_bstrOutputOption = bstrOutputOption
    m_strSharedContentPath = bstrSharedContentsPath
    If Not pControlXMLObject Is Nothing Then
        If TypeOf pControlXMLObject Is IXMLDOMDocument Then
            Set m_oControlInputDoc = pControlXMLObject
        End If
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Sub



