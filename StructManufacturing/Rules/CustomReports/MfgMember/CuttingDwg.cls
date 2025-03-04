VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuttingDwg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        :   CuttingDwg.cls
''
''  Description :
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

Private Const MODULE = "MfgMemberCustomReports.CuttingDwg"
Private Const IJMfgProfilePart = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"
Dim m_strLogFile As String
Dim m_oFileSystmeObject As FileSystemObject
Dim m_oPOM As IJDPOM

Private Sub Class_Initialize()
    Set m_oFileSystmeObject = New FileSystemObject
    Set m_oPOM = GetPOM
    Dim sTemp As String
    sTemp = Environ("TEMP")

    m_strLogFile = sTemp & "\GenerateMemberCuttingDwg.log"
    
    Open m_strLogFile For Append As #1
    
    Print #1, "Start - " & CStr(Now)
    Print #1, ""
End Sub

Private Sub Class_Terminate()
    Set m_oFileSystmeObject = Nothing
    Set m_oPOM = Nothing
    Print #1, ""
    Print #1, "End - " & CStr(Now)
    Print #1, ""
    Close #1
End Sub

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
Const METHOD = "IJDCustomReport_Generate"
On Error GoTo ErrorHandler
    
'    Dim strExtension As String
'    Dim strBaseOutputName As String
'
'    Dim lastPos As Long
'    lastPos = InStrRev(strFileName, ".")
'
'    strBaseOutputName = Left$(strFileName, lastPos - 1)
'    strExtension = Right$(strFileName, Len(strFileName) - lastPos + 1)
'
'    Dim oMfgMemberColl As IJElements
'    Set oMfgMemberColl = New JObjectCollection
'
'    '1. Get all MfgMemberPart
'    If pElements.Count > 0 Then
'        Dim oSelectedObj As Object
'        For Each oSelectedObj In pElements
'            GetAllMfgMemberParts oSelectedObj, oMfgMemberColl
'        Next
'        Set oSelectedObj = Nothing
'    End If
'
'    '2. Generate MemberCuttingDrawings
'    Dim oMfgOutput As IJMfgOutput
'
'    Dim strMemberCuttingDwgProgID As String
'    strMemberCuttingDwgProgID = "MfgMemberCuttingDrawings.DXF"
'
'    On Error Resume Next
'    Dim oMemberCuttingDwg As IGenerateDocument
'    Set oMemberCuttingDwg = CreateObject(strMemberCuttingDwgProgID)
'
'    If Not oMemberCuttingDwg Is Nothing Then
'        Dim nCount As Integer
'        nCount = 0
'
'        For Each oMfgOutput In oMfgMemberColl
'            Dim oXMLDocument As IXMLDOMDocument
'            Set oXMLDocument = oMfgOutput.GetData("CMfgMemberOutputCmd2", strFileName)
'
'            Dim strName As String
'            Dim oNamedItem As IJNamedItem
'            Set oNamedItem = oMfgOutput
'            strName = oNamedItem.Name
'
'            nCount = nCount + 1
'            If Not oXMLDocument Is Nothing Then
'                'Validate Output SMS_MACROS node
'                Dim oSMS_MACROSNode As IXMLDOMNode
'                Set oSMS_MACROSNode = oXMLDocument.selectSingleNode("//SMS_MACROS")
'
'                If Not oSMS_MACROSNode Is Nothing Then
'
'                    Dim strSharedContentPath As String
'                    strSharedContentPath = GetSharedContentPath
'
'                    Dim strControlInputXML As String
'                    strControlInputXML = strSharedContentPath + "\" + "StructManufacturing\SMS_SCHEMA\ControlInput.xml"
'
'                    If m_oFileSystmeObject.FileExists(strControlInputXML) = False Then
'                        Print #1, strControlInputXML + " doesn't exist"
'                        GoTo Continue
'                    Else
'                        Dim oControlInputXMLDocument As IXMLDOMDocument
'                        Set oControlInputXMLDocument = New DOMDocument
'
'                        Dim bSucess As Boolean
'                        bSucess = oControlInputXMLDocument.Load(strControlInputXML)
'                        If bSucess = False Then
'                            Print #1, strControlInputXML + " is not valid XML"
'                            GoTo Continue
'                         End If
'
'                        Dim oCONTROLINPUTSNode As IXMLDOMNode
'                        Set oCONTROLINPUTSNode = oControlInputXMLDocument.selectSingleNode("//CONTROLINPUTS")
'
'                        If oCONTROLINPUTSNode Is Nothing Then
'                            Print #1, strControlInputXML + "," + " CONTROLINPUTSNode node doesn't exist in the xml"
'                            GoTo Continue
'                        Else
'                            Dim bFileExist As Boolean
'                            bFileExist = False
'                            'Validate lua script location
'                            Dim strLuaScriptLocation As String
'                            strLuaScriptLocation = oControlInputXMLDocument.selectSingleNode("//MACROSCRIPTS/@PATH").Text
'
'                            If m_oFileSystmeObject.FolderExists(strLuaScriptLocation) = True Then
'                               bFileExist = True
'                            ElseIf m_oFileSystmeObject.FolderExists(strSharedContentPath + "\" + strLuaScriptLocation) = True Then
'                                bFileExist = True
'                            End If
'
'                            If bFileExist = False Then
'                                Print #1, strLuaScriptLocation + ", " + " lua script location doesn't exist"
'                                GoTo Continue
'                            End If
'
'                            'Validate detail scallop dxf location
'                            bFileExist = False
'                            Dim strDetailScallopLocation As String
'                            strDetailScallopLocation = oControlInputXMLDocument.selectSingleNode("//DETAILVIEWTEMPLATE/@PATH").Text
'
'                            If m_oFileSystmeObject.FileExists(strDetailScallopLocation) = True Then
'                               bFileExist = True
'                            ElseIf m_oFileSystmeObject.FileExists(strSharedContentPath + "\" + strDetailScallopLocation) = True Then
'                                bFileExist = True
'                            End If
'
'                            If bFileExist = False Then
'                                Print #1, strDetailScallopLocation + ", " + " Detail scallop dxf template file doesn't exist"
'                                GoTo Continue
'                            End If
'                        End If
'                    End If
'
'                    oMemberCuttingDwg.Initialize strControlInputXML, strSharedContentPath
'
'                    Dim partPath() As String
'                    ReDim partPath(1)
'                    partPath(0) = oXMLDocument.xml
'
'                    Dim strGuid As String
'                    If Not m_oPOM Is Nothing Then
'                        strGuid = m_oPOM.DbIdentifierFromMoniker(m_oPOM.GetObjectMoniker(oMfgOutput))
'                    End If
'
'                    Dim strOutputFileName As String
'                    strOutputFileName = strBaseOutputName + "_" + strGuid + ".dxf"
'
'                    Dim strDrawingTemplate As String
'                    strDrawingTemplate = strSharedContentPath + "\" + "StructManufacturing\DrawingTemplate\drawing_template_mcp_single_member_beam_CommonQty.dxf"
'                    If m_oFileSystmeObject.FileExists(strDrawingTemplate) = False Then
'                         Print #1, strDrawingTemplate + " doesn't exist"
'                         GoTo Continue
'                     End If
'
'                    Dim strReportName As String
'                    strReportName = oMemberCuttingDwg.Generate(partPath, strDrawingTemplate, False, strOutputFileName)
'
'                     'Delete log files
'                    If m_oFileSystmeObject.FileExists(strReportName) = True Then
'                         m_oFileSystmeObject.DeleteFile strReportName
'                    End If
'
'                    Print #1, CStr(nCount) + ". " + strName + ", " + partPath(0) + "," & strOutputFileName & ".dxf File Is generated"
'                Else
'                    Print #1, CStr(nCount) + ". " + strName + ", " + "SMS_MACROS node doesn't exist in the xml"
'                End If
'            Else
'                eCustomReportStatus = StrMfgNoObjectFound
'                Print #1, CStr(nCount) + ". " + strName + ", " + "No XML"
'            End If
'Continue:
'            Set oMfgOutput = Nothing
'        Next
'
'        Set oMemberCuttingDwg = Nothing
'    Else
'        eCustomReportStatus = StrMfgNoObjectFound
'        Print #1, "MfgMemberCuttingDrawings.DXF is not registered or can't be created"
'    End If
    
    eCustomReportStatus = StrMfgProcessFinished

Wrapup:
    Exit Sub
    
ErrorHandler:

    Set m_oFileSystmeObject = Nothing
    Set m_oPOM = Nothing
    eCustomReportStatus = StrMfgErrorUnknown
    Close #1
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function IsMfgMemberPart(oMfgProfilePart As IJMfgProfilePart) As Boolean
Const METHOD = "IsMfgMemberPart"
On Error GoTo ErrorHandler

    IsMfgMemberPart = False
    
    If TypeOf oMfgProfilePart.GetDetailedPart Is ISPSMemberPartPrismatic Then
        IsMfgMemberPart = True
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function


Private Sub GetAllMfgMemberParts(ByRef oObject As Object, ByRef oMfgMemberColl As IJElements)
Const METHOD = "GetAllMfgMemberParts"
On Error GoTo ErrorHandler

    If TypeOf oObject Is IJMfgProfilePart Then
        If IsMfgMemberPart(oObject) = True Then
            If oMfgMemberColl.Contains(oObject) = False Then
                oMfgMemberColl.Add oObject
            End If
        End If
    ElseIf TypeOf oObject Is IJAssemblyBase Then
        
        Dim oAssemblyParent As IJAssembly
        Set oAssemblyParent = oObject
        
        Dim oChildren As IJDTargetObjectCol
        Set oChildren = oAssemblyParent.GetChildren
        
        Dim Index As Long
        Dim oItem As Object
                        
        'Get each assembly child and check for plate part
        For Index = 1 To oChildren.Count
            'Get next item
            Set oItem = oChildren.Item(Index)
            GetAllMfgMemberParts oItem, oMfgMemberColl
            Set oItem = Nothing
        Next
    ElseIf TypeOf oObject Is ISPSMemberPartPrismatic Then
        Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
        Dim oColl As IJElements
        Set oColl = oStructMfgGlobals.GetMfgPart(oObject, IJMfgProfilePart)
        
        Dim oMfgProfilePart As IJMfgProfilePart
        For Each oMfgProfilePart In oColl
            If oMfgMemberColl.Contains(oMfgProfilePart) = False Then
                oMfgMemberColl.Add oMfgProfilePart
            End If
            Set oMfgProfilePart = Nothing
        Next
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function GetSharedContentPath() As String
Const METHOD = "GetSharedContentPath"
On Error GoTo ErrorHandler

    Dim oContext As IJContext
    Set oContext = GetJContext()
    
    Dim strContextString As String
    strContextString = "OLE_SERVER"
    
    Dim strSymbolShare As String
    
    If Not oContext Is Nothing Then
        strSymbolShare = oContext.GetVariable(strContextString)
    End If
    
    GetSharedContentPath = strSymbolShare
    
    Set oContext = Nothing
    
Cleanup:
    Exit Function
ErrorHandler:
    GoTo Cleanup
End Function

