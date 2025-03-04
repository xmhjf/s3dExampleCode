VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''***********************************************************************************
''  Copyright (C) 2013, Intergraph Corporation.  All rights reserved.
''
''  Project     : StrMfgDSTVOutput
''  File        : Customize.cls
''
''  Description :
''
''  Author      : Siva
''
''  History     :
''                10 Jan 13       Siva      - Creation
''***********************************************************************************
Option Explicit

Implements IJDMfgOutputCustomRule
Private SymPath As String

Private Const MODULE = "Customize"
Private Const dDistTolerance = 1
Private Const dAngleTolerance = 1

Private Sub IJDMfgOutputCustomRule_CustomAnnotation(ByVal oDispMfgPart As Object, ByVal oMfgGeom2d As GSCADMfgRulesDefinitions.IJMfgGeom2d, ByVal bstrFormat As String, ByVal oGeom2DNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomBevelInfo(ByVal oDispMfgPart As Object, ByVal oMfgGeom2d As GSCADMfgRulesDefinitions.IJMfgGeom2d, ByVal bstrFormat As String, ByVal oGeom2DNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler
    
    Dim strTempString As String
    Dim bInnerContour As Boolean
    bInnerContour = False
            
    If Not oMfgGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR Then
        If Not oMfgGeom2d.GetGeometryType = STRMFG_INNER_CONTOUR Then
            Exit Sub
        End If
        bInnerContour = True
    End If

    ' Get the bevel XML from Geom2D
    Dim xmlBevel As String
    oMfgGeom2d.GetBevelXML xmlBevel
    
    xmlBevel = "<SMS_DOC>" & xmlBevel & "</SMS_DOC>"
    Dim oBevelXmlDoc As DOMDocument
    Set oBevelXmlDoc = New DOMDocument
    
    ' Create Bevel XML doucment
    If Not oBevelXmlDoc.loadXML(xmlBevel) Then
        Exit Sub
    End If
    
    ' Get all the bevel nodes in the bevel xml
    Dim oBevelNodeList As IXMLDOMNodeList
    'Set oBevelNodeList = oBevelXmlDoc.getElementsByTagName("SMS_BEVEL")
    Set oBevelNodeList = oBevelXmlDoc.selectNodes("//SMS_BEVEL | //SMS_GRIND")
    
    Dim iCnt    As Long, jCnt    As Long
    Dim bBevelStart As Boolean
    Dim bBevelEnd As Boolean
    
    ' Iterate through all bevel nodes and add the bevel info to contour points
    For iCnt = 0 To oBevelNodeList.Length - 1
        Dim oElem   As IXMLDOMElement
        Set oElem = oBevelNodeList.Item(iCnt)
        
        Dim oPointList As IXMLDOMNodeList
        Set oPointList = oElem.getElementsByTagName("CVG_POINT")
        
        Dim oP0 As IXMLDOMNode
        Set oP0 = oPointList.Item(0)
        
        Dim dBevelStartX As Double
        Dim dBevelStartY As Double
        strTempString = oP0.Attributes.getNamedItem("X").nodeValue
        dBevelStartX = Val(strTempString)
        strTempString = oP0.Attributes.getNamedItem("Y").nodeValue
        dBevelStartY = Val(strTempString)
        
        Dim oP1 As IXMLDOMNode
        Set oP1 = oPointList.Item(1)
        
        Dim dBevelEndX As Double
        Dim dBevelEndY As Double
        strTempString = oP1.Attributes.getNamedItem("X").nodeValue
        dBevelEndX = Val(strTempString)
        strTempString = oP1.Attributes.getNamedItem("Y").nodeValue
        dBevelEndY = Val(strTempString)
        
        Dim oGeom2DElem As IXMLDOMElement
        Set oGeom2DElem = oGeom2DNode
        
        ' Get the points information from the Geom2D element node
        Dim oDataNodeList As IXMLDOMNodeList
        Set oDataNodeList = oGeom2DElem.getElementsByTagName("SMS_DATA")
        Dim oDataNode   As IXMLDOMNode
        Dim strDataNodeVal As String
        
        bBevelStart = False
        bBevelEnd = False
        
        Dim dADepth As Double
        Dim dAAngle As Double
        Dim dBDepth As Double
        Dim dBAngle As Double
        Dim dNoseDepth As Double
        Dim dNoseOrientationAngle As Double
        Dim dDDepth As Double
        Dim dDAngle As Double
        Dim dEDepth As Double
        Dim dEAngle As Double
        
        ' Get the Nose Depth
        Dim oNodeAttrib As IXMLDOMAttribute
        Set oNodeAttrib = oElem.getAttributeNode("DEPTH_A")
        strTempString = oNodeAttrib.Value
        dADepth = Val(strTempString)
        
        Set oNodeAttrib = oElem.getAttributeNode("ANGLE_A")
        strTempString = oNodeAttrib.Value
        dAAngle = Val(strTempString)
        
        ' Get the Nose Depth
        Set oNodeAttrib = oElem.getAttributeNode("DEPTH_B")
        strTempString = oNodeAttrib.Value
        dBDepth = Val(strTempString)
        
        Set oNodeAttrib = oElem.getAttributeNode("ANGLE_B")
        strTempString = oNodeAttrib.Value
        dBAngle = Val(strTempString)
        
        ' Get the Nose Depth
        Set oNodeAttrib = oElem.getAttributeNode("DEPTH_N")
        strTempString = oNodeAttrib.Value
        dNoseDepth = Val(strTempString)
        
        Set oNodeAttrib = oElem.getAttributeNode("ANGLE_N")
        strTempString = oNodeAttrib.Value
        dNoseOrientationAngle = Val(strTempString)
        
        ' Get the Nose Depth
        Set oNodeAttrib = oElem.getAttributeNode("DEPTH_D")
        strTempString = oNodeAttrib.Value
        dDDepth = Val(strTempString)
        
        Set oNodeAttrib = oElem.getAttributeNode("ANGLE_D")
        strTempString = oNodeAttrib.Value
        dDAngle = Val(strTempString)
        
        ' Get the Nose Depth
        Set oNodeAttrib = oElem.getAttributeNode("DEPTH_E")
        strTempString = oNodeAttrib.Value
        dEDepth = Val(strTempString)
        
        Set oNodeAttrib = oElem.getAttributeNode("ANGLE_E")
        strTempString = oNodeAttrib.Value
        dEAngle = Val(strTempString)
    
        Dim dMarkedSideDepth As Double
        Dim dMarkedSideAngle As Double
        
        Dim dUnMarkedSideDepth As Double
        Dim dUnMarkedSideAngle As Double
        
        ' Below we map the values given as input into values compatible with GNest
        ' The method is fairly simple:
        ' 1) Find out which case we are dealing with. Cases include:
        '       nose orientationangle =0, >0 and <0
        ' 2) Check which depths are present and make sure the most important ones gets mapped.
        '    Order of importance is: Nose, 1st bevel depth(B or D), 2nd bevel depth(A or E)
        ' 3) If only one value is present on a side of a bevel we move the value towards the middle
        ' 4) Put the new values into the XML.
        
        If (Abs(dNoseOrientationAngle) < dAngleTolerance) Then
            If (Abs(dBAngle) > dAngleTolerance) Then
                dMarkedSideDepth = dDDepth + dEDepth + dNoseDepth
                dMarkedSideAngle = dBAngle
            End If
            
            If (Abs(dDAngle) > dAngleTolerance) Then
                dUnMarkedSideDepth = dDDepth + dEDepth
                dUnMarkedSideAngle = -1 * dDAngle
            End If
        Else
            'nose angle is <> 0 and we have to shift the nose to one of the sides to be able to
            ' send the nose angle to GNEST. The cost is that we lose part of the bevel on one side.
            ' See if there is NoseDepth
            
            If dNoseOrientationAngle >= 0 Then
                dMarkedSideDepth = dDDepth + dEDepth
                dMarkedSideAngle = dNoseOrientationAngle
                
                If (Abs(dDAngle) > dAngleTolerance) Then
                    dUnMarkedSideDepth = dDDepth + dEDepth
                    dUnMarkedSideAngle = -1 * dDAngle
                End If
            Else
                If (Abs(dBAngle) > dAngleTolerance) Then
                    dMarkedSideDepth = dDDepth + dEDepth + dNoseDepth
                    dMarkedSideAngle = dBAngle
                End If
                
                dUnMarkedSideDepth = dDDepth + dEDepth + dNoseDepth
                dUnMarkedSideAngle = dNoseOrientationAngle
            End If
        End If
        
        If Abs(dMarkedSideAngle) >= 90 Then
            dMarkedSideAngle = 0#
        End If
        
        If Abs(dUnMarkedSideAngle) >= 90 Then
            dUnMarkedSideAngle = 0#
        End If
        
'        dMarkedSideDepth = Format(dMarkedSideDepth, "#0.00")
'        dMarkedSideAngle = Format(dMarkedSideAngle, "#0.00")
        
'        dUnMarkedSideDepth = Format(dUnMarkedSideDepth, "#0.00")
'        dUnMarkedSideAngle = Format(dUnMarkedSideAngle, "#0.00")
    
'        Dim oCousinNode As IXMLDOMNode
'        On Error Resume Next
'        Set oCousinNode = oGeom2DElem.previousSibling.lastChild.lastChild
'        On Error GoTo ErrorHandler
'        If Not oCousinNode Is Nothing Then
'            Set oDataNode = oCousinNode
'
'            Dim strNodeVal1 As String
'            strDataNodeVal = oDataNode.nodeTypedValue
'            strNodeVal1 = Trim(strDataNodeVal)
'
'            Dim lStartPos1    As Long
'            lStartPos1 = InStr(1, strNodeVal1, " ")
'            strNodeVal1 = Trim(Mid(strNodeVal1, lStartPos1 + 1))
'
'            Dim lEndPos1    As Long
'            lEndPos1 = InStr(1, strNodeVal1, " ")
'
'            Dim strX1    As String
'            strX1 = Mid(strNodeVal1, 1, lEndPos1)
'
'            strNodeVal1 = Trim(Replace(strNodeVal1, strX1, ""))
'            lEndPos1 = InStr(1, strNodeVal1, " ")
'
'            Dim strY1    As String
'            strY1 = Mid(strNodeVal1, 1, lEndPos1)
'
'            Dim dX1 As Double, dY1 As Double
'            dX1 = Val(strX1)
'            dY1 = Val(strY1)
'            If (Abs(dBevelStartX - dX1) < dDistTolerance) And (Abs(dBevelStartY - dY1) < dDistTolerance) Or _
'                (Abs(dBevelEndX - dX1) < dDistTolerance) And (Abs(dBevelEndY - dY1) < dDistTolerance) Then
'                If (Abs(dUnMarkedSideDepth) < dDistTolerance) Then
'                    strDataNodeVal = strDataNodeVal & "  " & Format(dMarkedSideAngle, "#0.00") & "  " & Format(dMarkedSideDepth, "#0.00")
'                Else
'                    strDataNodeVal = strDataNodeVal & "  " & Format(dMarkedSideAngle, "#0.00") & "  " & Format(dMarkedSideDepth, "#0.00") & "  " & Format(dUnMarkedSideAngle, "#0.00") & "  " & Format(dUnMarkedSideDepth, "#0.00")
'                End If
'
'                If (Abs(dBevelEndX - dX1) < dDistTolerance) And (Abs(dBevelEndY - dY1) < dDistTolerance) Then
'                    dBevelEndX = dBevelStartX
'                    dBevelEndY = dBevelStartY
'                End If
'
'                ' set the bevel on the data node.
'                oDataNode.nodeTypedValue = strDataNodeVal
'                bBevelStart = True
'            End If
'        End If

        bBevelStart = False
       ' Find the point to which bevel info need to be appended
        For jCnt = 0 To oDataNodeList.Length - 1
            Set oDataNode = Nothing
            Set oDataNode = oDataNodeList.Item(jCnt)
            
            Dim strNodeVal As String
            strDataNodeVal = oDataNode.nodeTypedValue
            strNodeVal = Trim(strDataNodeVal)
            
            Dim lStartPos    As Long
            lStartPos = InStr(1, strNodeVal, " ")
            strNodeVal = Trim(Mid(strNodeVal, lStartPos + 1))
            
            Dim lEndPos    As Long
            lEndPos = InStr(1, strNodeVal, " ")
            
            Dim strX    As String
            strX = Mid(strNodeVal, 1, lEndPos)
            
            strNodeVal = Trim(Replace(strNodeVal, strX, ""))
            lEndPos = InStr(1, strNodeVal, " ")
            
            Dim strY    As String
            strY = Mid(strNodeVal, 1, lEndPos)
            
            Dim dX As Double, dY As Double
            dX = Val(strX)
            dY = Val(strY)
            
            If bBevelStart = False Then
                If (Abs(dBevelStartX - dX) < dDistTolerance) And (Abs(dBevelStartY - dY) < dDistTolerance) Then
                    If Not bInnerContour Then
                        bBevelStart = True
                    Else
                        If (oDataNode Is oDataNodeList.Item(0)) Then
                            GoTo NextDataNode
                        Else
                            bBevelStart = True
                        End If
                    End If
                ElseIf (Abs(dBevelEndX - dX) < dDistTolerance) And (Abs(dBevelEndY - dY) < dDistTolerance) Then
                    dBevelEndX = dBevelStartX
                    dBevelEndY = dBevelStartY
                    bBevelStart = True
                Else
                    GoTo NextDataNode
                End If
            Else
                If (Abs(dBevelEndX - dX) < dDistTolerance) And (Abs(dBevelEndY - dY) < dDistTolerance) Then
                    Exit For ' End data node for bevel is reahed
                End If
            End If
            
            If bBevelStart Then
    '            If (Abs(dMarkedSideDepth) < dDistTolerance) Then
    '                strDataNodeVal = "  " & strDataNodeVal & "  " & Format(dUnMarkedSideAngle, "#0.00") & "  " & Format(dUnMarkedSideDepth, "#0.00")
                If (Abs(dUnMarkedSideDepth) < dDistTolerance) Then
                    strDataNodeVal = "  " & strDataNodeVal & "  " & FormatDouble(dMarkedSideAngle) & "  " & FormatDouble(dMarkedSideDepth)
                Else
                    strDataNodeVal = "  " & strDataNodeVal & "  " & FormatDouble(dMarkedSideAngle) & "  " & FormatDouble(dMarkedSideDepth) & "  " & FormatDouble(dUnMarkedSideAngle) & "  " & FormatDouble(dUnMarkedSideDepth)
                End If
            End If
            
            ' set the bevel on the data node.
            oDataNode.nodeTypedValue = strDataNodeVal
NextDataNode:
        Next
        dMarkedSideDepth = 0
        dMarkedSideAngle = 0
        dUnMarkedSideDepth = 0
        dUnMarkedSideAngle = 0
NextBevelNode:
    Next
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomGeometryInfo(ByVal oDispMfgPart As Object, ByVal oMfgGeom2d As GSCADMfgRulesDefinitions.IJMfgGeom2d, ByVal bstrFormat As String, ByVal oGeom2DNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomLabel(ByVal oDispMfgPart As Object, ByVal bstrXSectionType As String, ByVal oLabelNode As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler
        
    Dim bstrQuery     As String

    Dim oCtrlXMLDoc As IXMLDOMDocument
    Set oCtrlXMLDoc = New DOMDocument
    oCtrlXMLDoc.resolveExternals = True
    oCtrlXMLDoc.validateOnParse = False
    
    Dim strCtrlXMLPath As String
    strCtrlXMLPath = GetSymbolSharePath & "\StructManufacturing\DSTV_OUTPUT\CONTROL_XML.xml"
    
    Dim bLoad   As Boolean
    bLoad = oCtrlXMLDoc.Load(strCtrlXMLPath)
    
    If bLoad = False Then
        Exit Sub
    End If
    
            
    Dim oNamedItem As IJNamedItem
    Dim oMfgPart As IJMfgProfilePart
    Dim oMfgPlatePart As IJMfgPlatePart
    On Error Resume Next
    Set oMfgPart = oDispMfgPart
    Set oMfgPlatePart = oDispMfgPart
    On Error GoTo ErrorHandler
    If Not oMfgPart Is Nothing Then
        Set oNamedItem = oMfgPart.GetDetailedPart
    bstrQuery = "//DSTV_CONTROL/DSTV_XSECTION[@S3D_XSECTION_TYPE='" & bstrXSectionType & "']/DSTV_FACES/DSTV_FACE[@NAME='webleft']/@FACE_CHAR"
    ElseIf Not oMfgPlatePart Is Nothing Then
        oMfgPlatePart.GetDetailedPart oNamedItem
    bstrQuery = "//DSTV_CONTROL/DSTV_XSECTION[@S3D_XSECTION_TYPE='" & bstrXSectionType & "']/DSTV_FACES/DSTV_FACE[@NAME='base']/@FACE_CHAR"
    End If
    Dim oReturnNode     As IXMLDOMNode
    Set oReturnNode = oCtrlXMLDoc.selectSingleNode(bstrQuery)
    
    Dim bstrWebLeft As String
    If Not oReturnNode Is Nothing Then
        bstrWebLeft = oReturnNode.nodeValue
    End If
    Dim strPartName As String
    If Not oNamedItem Is Nothing Then
        strPartName = oNamedItem.Name
    Else
        strPartName = "PartName"
    End If
    
    Dim dProfileHeight As Double
    Dim dProfileLength As Double
    If Not oMfgPart Is Nothing Then
        Dim oProfilePartSupport As IJProfilePartSupport
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
        
        Set oPartSupport.Part = oMfgPart.GetDetailedPart
        Set oProfilePartSupport = oPartSupport
        
        oProfilePartSupport.GetWebDepth dProfileHeight
        dProfileLength = 0.25
    Else
        
        dProfileHeight = Val(oLabelNode.ownerDocument.selectSingleNode("//SMS_DATA[@TYPE='web_height']").Text)
        dProfileHeight = dProfileHeight * 0.0025
        dProfileLength = Val(oLabelNode.ownerDocument.selectSingleNode("//SMS_DATA[@TYPE='part_length']").Text)
        dProfileLength = dProfileLength / 2000
    End If
    Dim oDOMLabelNode   As IXMLDOMNode
    Set oDOMLabelNode = oLabelNode

    Dim oDataNode As IXMLDOMNode
    Set oDataNode = oDOMLabelNode.selectSingleNode("//SMS_DATA[@TYPE='text']")
    
    If Not oDataNode Is Nothing Then
        oDataNode.nodeTypedValue = "  " & bstrWebLeft & " " & FormatDouble(dProfileLength * 1000) & "u      " & FormatDouble(dProfileHeight * 0.2 * 1000) & "    00.00      25  " & strPartName
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDMfgOutputCustomRule_CustomPartInfo(ByVal oDispMfgPart As Object, ByVal bstrFormat As String, ByVal oDoc As Object)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler
    
    Dim oXMLDoc As IXMLDOMDocument
    Set oXMLDoc = oDoc
        
    ' set order identification
    SetNodeValue oXMLDoc, "//SMS_DATA[@TYPE='order_identification']", "  Training_Job"
    
    ' set drawing identification
    SetNodeValue oXMLDoc, "//SMS_DATA[@TYPE='drawing_identification']", "  0"
    
    ' set phase identification
    SetNodeValue oXMLDoc, "//SMS_DATA[@TYPE='phase_identification']", "  1B"
    
    ' set text info
    ' SetNodeValue oXMLDoc, "//SMS_DATA[@TYPE='text_info1']", "  sample"
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SetNodeValue(oXMLDoc As IXMLDOMDocument, strQuery As String, strNodeValue As String)
    Const METHOD = "IJDMfgOutputCustomRule_CustomPartInfo"
    On Error GoTo ErrorHandler
        
    Dim oDataNode As IXMLDOMNode
    Set oDataNode = oXMLDoc.selectSingleNode(strQuery)
    
    If Not oDataNode Is Nothing Then
        oDataNode.nodeTypedValue = strNodeValue
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetSymbolSharePath() As String
    Const METHOD = "GetSymbolLocationPath"
    On Error GoTo ErrorHandler
    
    Dim oContext As IJContext
    Dim strContextString As String
    Dim strSymbolShare As String
    
    strContextString = "OLE_SERVER"
    
    'Get IJContext
    Set oContext = GetJContext()
    
    If Not oContext Is Nothing Then
        strSymbolShare = oContext.GetVariable(strContextString)
    End If
    GetSymbolSharePath = strSymbolShare
    Set oContext = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function FormatDouble(dValue As Double) As String
    Const METHOD = "FormatDouble"
    On Error GoTo ErrorHandler

    Dim strDouble As String
    strDouble = Format(dValue, "#0.00")
    strDouble = Replace(strDouble, ",", ".")
    FormatDouble = strDouble
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function
