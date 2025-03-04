VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysGenXMLData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CPlateReport.cls
''
''  Description : Populates a text file with PinJig related data
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

Private Const MODULE = "CustomPlateReports.CSysGenXMLData"

'- MfgPlate, MfgProfile ----------------------------------------------------'
Private Const IJMfgPlatePart = "{BCA241EE-F5E1-47A8-90DA-17141F9D39BC}"
Private Const IJMFGPROFILEPARTGUID = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"

'Private Const XML_DATA_FORMAT = "CMfgPlateOutputCmd_DEFAULT"
Private Const XML_DATA_FORMAT = "CMfgPlateOutputCmd_Annotation"
Private Const PLATES_NODE_NAME = "SMS_PLATES"
Private Const PROFILES_NODE_NAME = "SMS_PROFILES"

Private m_oXMLOutput As IXMLDOMDocument
Private m_strSMS_SMS_OPENING_TAG As String
Private m_strSMS_SMS_CLOSING_TAG As String
Private m_strSMS_PLATES As String
Private m_strSMS_PROFILES As String

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
    
    Dim bRootNodeNeeded As Boolean
    bRootNodeNeeded = True
    
    'Set default xml root node to use in the output
    m_strSMS_SMS_OPENING_TAG = "<SMS_SMS>"
    m_strSMS_SMS_CLOSING_TAG = "</SMS_SMS>"
    m_strSMS_PLATES = ""
    m_strSMS_PROFILES = ""
    
    If pElements.Count > 0 Then
        Dim oSelectedObj As Object
        For Each oSelectedObj In pElements
        
            Dim oMfgOutput As IJMfgOutput
            Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
            Dim oColl As IJElements
            Dim oXMLObject As Object
            
            If TypeOf oSelectedObj Is IJPlatePart Then
            
                Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
                Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
                Set oPlateWrapper.object = oSelectedObj
                
                Dim oMfgPlatePart As IJMfgPlatePart
                If oPlateWrapper.PlateHasMfgPart(oMfgPlatePart) Then
                    Set oMfgOutput = oMfgPlatePart
                End If
                
                Set oPlateWrapper = Nothing
                Set oMfgPlatePart = Nothing
            
            ElseIf TypeOf oSelectedObj Is IJProfilePart Or TypeOf oSelectedObj Is IJStructProfilePart Then
            
                Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
                Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
                Set oProfileWrapper.object = oSelectedObj
                
                Dim oMfgProfilPart As IJMfgProfilePart
                If oProfileWrapper.ProfileHasMfgPart(oMfgProfilPart) Then
                    Set oMfgOutput = oMfgProfilPart
                End If
                
                Set oProfileWrapper = Nothing
                Set oMfgProfilPart = Nothing
                
            ElseIf TypeOf oSelectedObj Is IJMfgPlatePart Or TypeOf oSelectedObj Is IJMfgProfilePart Then
                Set oMfgOutput = oSelectedObj
            Else
            
            End If
            
'            If Not TypeOf oSelectedObj Is IJMfgPlatePart And Not TypeOf oSelectedObj Is IJMfgProfilePart Then
'                'This is not a mfg part. Get Mfg Part from object
'                'Assume oSelectedObj is a plate and attempt to get its MfgPlatePart
'                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJMfgPlatePart)
'
'                'If oColl is nothing, then oSelectedObj was not a plate and get the MfgprofilePart
'                If oColl Is Nothing Then
'                    Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJMFGPROFILEPARTGUID)
'                End If
'
'                Set oMfgOutput = oColl.Item(1) 'count will be one
'            Else
'                Set oMfgOutput = oSelectedObj
'            End If
            
            Set oXMLObject = oMfgOutput.GetOriginalData(XML_DATA_FORMAT)
            
            GroupXML oXMLObject
            
            If bRootNodeNeeded Then
                GetRootNode oXMLObject
                bRootNodeNeeded = False
            End If
            
            Set oMfgOutput = Nothing
            Set oStructMfgGlobals = Nothing
            Set oColl = Nothing
            Set oXMLObject = Nothing
            
        Next
        
        Set m_oXMLOutput = New DOMDocument
        m_oXMLOutput.loadXML MergeXML()
        
        If Not m_oXMLOutput Is Nothing Then
            m_oXMLOutput.Save strFileName
            Set m_oXMLOutput = Nothing
        End If
    
    End If
    
    eCustomReportStatus = StrMfgProcessFinished
 
Cleanup:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    eCustomReportStatus = StrMfgErrorUnknown
End Sub

Private Sub GetRootNode(oPartXML As Object)
    Const METHOD = "GroupXML"
    On Error GoTo ErrorHandler
    
    Dim oDOMDocTest As IXMLDOMDocument
    Set oDOMDocTest = oPartXML
    
    If Not oDOMDocTest Is Nothing Then
    
        Dim oRootNode As IXMLDOMNode
        Set oRootNode = oDOMDocTest.selectSingleNode("/SMS_SMS")
        
        Dim strXML As String
        strXML = oRootNode.xml
        
        m_strSMS_SMS_OPENING_TAG = Left(strXML, InStr(strXML, ">"))
        m_strSMS_SMS_CLOSING_TAG = "</" + oRootNode.nodeName + ">"
    
    End If
    
Cleanup:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub GroupXML(oPartXML As Object)
    Const METHOD = "GroupXML"
    On Error GoTo ErrorHandler

    Dim oDOMDoc As IXMLDOMDocument
    Set oDOMDoc = oPartXML
    
    If Not oDOMDoc Is Nothing Then
    
        Dim oNodeList As IXMLDOMNodeList
        Set oNodeList = oDOMDoc.selectNodes("/SMS_SMS/SMS_PLATES/SMS_PLATE|/SMS_SMS/SMS_PROFILES/SMS_PROFILE")
        
        Dim strNodeXML As String
        Dim oXMLNode As IXMLDOMNode
        For Each oXMLNode In oNodeList
            strNodeXML = oXMLNode.xml
            
            If Not oXMLNode.nodeName <> "SMS_PLATE" Then
                m_strSMS_PLATES = m_strSMS_PLATES + strNodeXML + vbCrLf
            ElseIf Not oXMLNode.nodeName <> "SMS_PROFILE" Then
                m_strSMS_PROFILES = m_strSMS_PROFILES + strNodeXML + vbCrLf
            Else
                MsgBox "Node Name: " + oXMLNode.nodeName + " not handled"
            End If
        Next
    
    End If

Cleanup:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function MergeXML() As String
    Const METHOD = "MergeXML"
    On Error GoTo ErrorHandler
    
    Dim strMergedXML As String
    strMergedXML = m_strSMS_SMS_OPENING_TAG + vbCrLf
    
    If Len(m_strSMS_PLATES) > 0 Then
        strMergedXML = strMergedXML + "<" + PLATES_NODE_NAME + ">" + vbCrLf + m_strSMS_PLATES + "</" + PLATES_NODE_NAME + ">" + vbCrLf
    End If
    
    If Len(m_strSMS_PROFILES) > 0 Then
        strMergedXML = strMergedXML + "<" + PROFILES_NODE_NAME + ">" + vbCrLf + m_strSMS_PROFILES + "</" + PROFILES_NODE_NAME + ">" + vbCrLf
        
    End If
    
    strMergedXML = strMergedXML + m_strSMS_SMS_CLOSING_TAG
    
    MergeXML = strMergedXML

Cleanup:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function
