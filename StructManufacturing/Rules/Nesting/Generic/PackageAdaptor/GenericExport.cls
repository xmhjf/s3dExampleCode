VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------
' Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
' Project
'   ISMfgGenericExport
'
' File
'   GenericExport
'
' Description
'   This file contains the package and delivery functions for Generic nesting
'
' Author
'   Kristian Kamph
'
' History:
'   2006-07-07  K. Kamph     Creation date
'----------------------------------------------------------------------------

Option Explicit
Implements IJMfgNestDelivery

Private Const MODULE = "GenericExport"

Private m_strOutputPath As String
Private m_strBinPath As String
Private m_strNestingServerName As String
Private m_strNestingDBName As String

Private m_bIndividualFile As Boolean

'Will be used as parameter in CreateProcess()
Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

'Will be used as parameter in CreateProcess()
Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Const SYNCHRONIZE = 1048576
Const NORMAL_PRIORITY_CLASS = &H20&
Const CREATE_NO_WINDOW = &H8
Const INFINITE = -1&

Const SW_HIDE = &O0
Const STARTF_USESHOWWINDOW = &H1
Const SERVER_REGISTRY_KEY = "\HKEY_CURRENT_USER\Software\Intergraph\Applications\Environments\CommonApp\ProjectDB"
Const SERVER_REGISTRY_NAME = "PhysicalName"

'For executing GSCAD2GNEST.exe module
Private Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
   (ByVal lpApplicationName As String, _
   ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, _
   lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   lpEnvironment As Any, _
   ByVal lpCurrentDriectory As String, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long




Private Sub Class_Initialize()

ErrorHandler:
End Sub

Private Sub Class_Terminate()

ErrorHandler:
End Sub

Private Sub IJMfgNestDelivery_PackageAndDeliver(ByVal pCol As Object, ByVal bstrRuleName As String, ByVal bstrFileName As String, ByVal bIndividualFile As Boolean)
     Const sMETHOD As String = "IJMfgNestDelivery_PackageAndDeliver"
    On Error GoTo ErrorHandler
    
    m_bIndividualFile = bIndividualFile
    
    Dim pCollection As Collection
    m_strNestingServerName = GetRuleValue(bstrRuleName, "NestingServerName")
    m_strNestingDBName = GetRuleValue(bstrRuleName, "NestingDBName")
    m_strOutputPath = GetRuleValue(bstrRuleName, "NestingFilePath")
    m_strBinPath = GetRuleValue(bstrRuleName, "NestingExecutablePath")
    If m_strOutputPath = "" Then
        m_strOutputPath = GetTempDirectory
    End If
    Set pCollection = pCol
    Dim strFileNames() As String
    
    MergeDocuments pCollection, bIndividualFile
    strFileNames = SendDocumentsToNesting(pCollection, bstrFileName)
    Exit Sub
ErrorHandler:
End Sub

Private Function SendDocumentsToNesting(pCollection As Collection, ByVal bstrFileName As String) As String()
    Const sMETHOD As String = "SendDocumentsToNesting"
    On Error GoTo ErrorHandler
    
    Dim oXMLDom As DOMDocument
    Dim strSeqXMLFile As String
    Dim strLogFile As String
    Dim strReturn() As String
    ReDim strReturn(pCollection.count - 1)
    Dim count As Integer
    count = 0
    Dim strFile As String
    Dim strFileExt As String
    Dim lPos As Long

    lPos = InStrRev(bstrFileName, ".")
    strFile = Trim(Left$(bstrFileName, lPos - 1))
    strFileExt = Trim(Right$(bstrFileName, Len(bstrFileName) - lPos + 1))

    
    For Each oXMLDom In pCollection
        If m_bIndividualFile Then
            oXMLDom.save strFile & count & strFileExt
        Else
            oXMLDom.save bstrFileName
        End If
        strReturn(count) = strSeqXMLFile
        count = count + 1
    Next oXMLDom
    SendDocumentsToNesting = strReturn
    Exit Function
ErrorHandler:
End Function

Private Sub MergeDocuments(pCollection As Collection, bIndividualFile As Boolean)
    Const sMETHOD As String = "MergeDocuments"
    On Error GoTo ErrorHandler
    
    Dim oXMLDom                 As DOMDocument
    Dim oBoardsideNodeList      As IXMLDOMNodeList
    Dim oElement                As IXMLDOMElement
    Dim oUnknown                As IUnknown
    Dim oMfgPlateSym            As IJMfgPlatePart
    Dim oMfgPlate               As IJMfgPlatePart
    Dim rtStr                   As String
    Dim bSame                   As Boolean
    Dim strPartGUIDSym          As String
    Dim oFinalDom               As New DOMDocument
    Dim oSMS                    As IXMLDOMElement
    Dim regPath                 As String
    Dim ShipNumber              As String
    Dim ShipName                As String
    Dim oPlates                 As IXMLDOMElement
    Dim oProfiles               As IXMLDOMElement
    Dim oTemplateSets           As IXMLDOMElement
    Dim oBuiltUps               As IXMLDOMElement
    Dim oPlateDomColl           As New Collection
    Dim oChildDom               As New DOMDocument
    Dim oNodeList               As IXMLDOMNodeList
    Dim oAppendNode             As IXMLDOMNode
    Dim oComparisonRule         As IJMfgComparisonRule
    Dim oName                   As IJNamedItem
    Dim i As Integer
    Dim j As Integer
    
    Dim oCommonPartMember As IJCommonPartMember
    Dim oCmnAppUtil As IJDCmnAppGenericUtil
    Set oCmnAppUtil = New CmnAppGenericUtil
    Dim oAccessMiddle As IJDAccessMiddle
    oCmnAppUtil.GetActiveConnection oAccessMiddle
    
    Dim oPOM As IJDPOM
    Set oPOM = oAccessMiddle.GetResourceManagerFromType("Model")
    
    'Comparison rule
    Dim oRule               As IJSRDRule
    Dim oSRDQuery           As IJSRDQuery
    Dim oRuleQuery          As IJSRDRuleQuery
    Dim strIsIDRules()      As String
    Dim strRuleProgId       As String
    Set oSRDQuery = New SRDQuery
    Set oRuleQuery = oSRDQuery.GetRulesQuery
    
    strIsIDRules = oRuleQuery.GetRuleNames("Identical")
    ' Default rule
    Set oRule = oRuleQuery.GetRule(strIsIDRules(0))
        strRuleProgId = oRule.ProgId
    Set oRule = Nothing
    Set oSRDQuery = Nothing
    Set oRuleQuery = Nothing
    
    Set oComparisonRule = SP3DCreateObject(strRuleProgId)
    
    'For non corresponding mfg plates.
    'Close
    Dim FileNumber
    FileNumber = FreeFile
    Open m_strOutputPath & "\StatusOfMergingPart.log" For Output As #FileNumber
    Print #FileNumber, "Result of merging Part data"
    Print #FileNumber, "---------------------------"
    
    Close #FileNumber
    Open m_strOutputPath & "\StatusOfMergingPart.log" For Append As #FileNumber
    'IJComm
    For Each oXMLDom In pCollection
        If Not oXMLDom Is Nothing Then
            Set oBoardsideNodeList = oXMLDom.getElementsByTagName("SMS_PART_INFO")
            
            If oBoardsideNodeList.length > 0 Then
                Set oElement = oBoardsideNodeList.Item(0)
    
                If Not ((oElement.getAttributeNode("PART_SYM_GUID") Is Nothing)) Then
                    If oElement.getAttributeNode("PART_SYM_GUID").Text <> "" Then
                        'Get moniker of the symmetrical Mfg part
                        Dim oOrigObject As Object
                        Dim oSymObject As Object
                        Dim oOrigDetObject As Object
                        Dim oSymDetObject As Object
                        
                        Set oUnknown = oPOM.MonikerFromDbIdentifier("{" + Trim(oElement.getAttributeNode("PART_GUID").Text) + "}")
                        Set oOrigObject = oPOM.GetObject(oUnknown)
                        Set oUnknown = Nothing
                        
                        Set oUnknown = oPOM.MonikerFromDbIdentifier("{" + Trim(oElement.getAttributeNode("PART_SYM_GUID").Text) + "}")
                        Set oSymObject = oPOM.GetObject(oUnknown)
                        Set oUnknown = Nothing
                        
                        'Get the symmetrical Mfg part from moniker
                        Dim oMfgChild As IJMfgChild
                        Set oMfgChild = oOrigObject
                        Set oOrigDetObject = oMfgChild.getParent
                        
                        Set oMfgChild = oSymObject
                        Set oSymDetObject = oMfgChild.getParent
        
                        'Get First detailed part
                        Dim oOrigDetName As IJNamedItem
                        Dim oOrigMfgName As IJNamedItem
                        Set oOrigMfgName = oOrigObject
                        Set oOrigDetName = oOrigDetObject
                        
                        Dim sOrigBoard As String
                        Dim sOpoBoard As String
                        'Set origin mfg part's boardside information
                        sOrigBoard = oElement.getAttributeNode("PART_BOARDSIDE_PROCESSED").Text
                        If sOrigBoard = "P" Then
                            sOpoBoard = "S"
                        Else
                            sOpoBoard = "P"
                        End If
                        
                        If oSymObject Is Nothing Then
                            Set oName = oOrigDetObject
                            
                            Print #FileNumber, vbTab + ">>Detailed Part  : " + oOrigDetName.Name + " (" + sOrigBoard + ") " + oName.Name + " (" + sOpoBoard + ") " + vbCrLf + _
                                      vbTab + "  MFG Plate Part : " + oOrigMfgName.Name + " (" + sOrigBoard + ") " + "None" + " (" + sOpoBoard + ") " + vbCrLf + _
                                      vbTab + " The MfgPlate part was not sent due to the symmetrical MfgPlate part is not existing."
                        Else
                            'If two MFG plate parts are completely identical, give "B" value to <PART_BOARDSIDE> element.
                            Dim fSame As Boolean
                            Set oCommonPartMember = oOrigDetObject
                            fSame = False
                            If Not oCommonPartMember Is Nothing Then
                                fSame = oCommonPartMember.InSameGroup(oSymDetObject)
                            ElseIf TypeName(oOrigObject) = "IJMfgPlatePart" Then
                                fSame = oComparisonRule.CompareMfgPlate(oOrigObject, oSymObject, rtStr)
                            ElseIf TypeName(oOrigObject) = "IJMfgProfilePart" Then
                                fSame = oComparisonRule.CompareMfgProfile(oOrigObject, oSymObject, rtStr)
                            End If
                            If fSame = True Then
                                If sOrigBoard = "P" Then
                                    oElement.getAttributeNode("PART_BOARDSIDE").Text = "B"
                                    If TypeName(oOrigObject) = "IJMfgProfilePart" Then
                                        If oBoardsideNodeList.length > 1 Then
                                            Set oElement = oBoardsideNodeList.Item(1)
                                            oElement.getAttributeNode("PART_BOARDSIDE").Text = "B"
                                        End If
                                        If oBoardsideNodeList.length > 2 Then
                                            Set oElement = oBoardsideNodeList.Item(2)
                                            oElement.getAttributeNode("PART_BOARDSIDE").Text = "B"
                                        End If
                                    End If
                                    Dim oPartElement As IXMLDOMElement
                                    Set oName = oSymObject
                                    Set oPartElement = oXMLDom.getElementsByTagName("SMS_PART_INFO").Item(0)
                                    If oPartElement.getAttributeNode("PART_BOARDSIDE_PROCESSED").Text = "P" Then
                                        Print #FileNumber, vbTab + ">>" + oOrigMfgName.Name + " and " + oName.Name + _
                                        " were transferd to GNEST with name of " + oOrigMfgName.Name + vbCrLf
                                    End If
                                    oPlateDomColl.Add oXMLDom
                                End If
                            Else 'Parts are not symmetrical at Mfg Level
                                Set oName = oOrigDetObject
                                Print #FileNumber, vbTab + ">>Detailed Part  : " + oOrigDetName.Name + " (" + sOrigBoard + ") " + oName.Name + " (" + sOpoBoard + ") " + vbCrLf + _
                                    vbTab + "  MFG Plate Part : " + oOrigMfgName.Name + " (" + sOrigBoard + ") " + oName.Name + " (" + sOpoBoard + ") " + vbCrLf
                                oPlateDomColl.Add oXMLDom
                            End If
                        End If
                    Else
                        oPlateDomColl.Add oXMLDom
                    End If
                Else
                    oPlateDomColl.Add oXMLDom
                End If
            End If
        End If
    Next oXMLDom
    
    Close #FileNumber

    Dim strSchemaName As String

    Dim oProjectRoot As GSCADProjMgmt.IJProjectRoot
    Set oProjectRoot = GetActiveShipClass
    Dim strServerName As String
    Dim strSiteName As String
    Dim strTempString As String
    GetShipInfo ShipNumber, strSiteName, strServerName, strTempString
    If Not bIndividualFile Then
        'Add Plate xml nodes and profile xml nodes
        Dim oNode As IXMLDOMNode
        For i = 1 To oPlateDomColl.count
            oChildDom.loadXML oPlateDomColl.Item(i).xml
            If oSMS Is Nothing Then
                Set oSMS = oChildDom.documentElement
            Else
                Set oNodeList = oChildDom.selectNodes("/*/*")
                For j = 0 To oNodeList.length - 1
                    Set oAppendNode = oSMS.selectSingleNode("//" & oNodeList.Item(j).baseName)
                    If oAppendNode Is Nothing Then
                        oSMS.appendChild oNodeList.Item(j)
                    Else
                        For Each oNode In oNodeList.Item(j).childNodes
                            oAppendNode.appendChild oNode
                        Next oNode
                    End If
                Next
            End If
        Next
        oSMS.setAttribute "SHIP_NUMBER", ShipNumber
        oSMS.setAttribute "SHIP_NAME", ShipName
        oFinalDom.appendChild oSMS
        pCollection.Add oFinalDom
    Else
        Dim oTempCollection As New Collection
        For Each oXMLDom In oPlateDomColl
            If Not oXMLDom Is Nothing Then
                Set oFinalDom = New DOMDocument
                oFinalDom.async = True
                ShipName = ShipNumber
                
                If oSMS Is Nothing Then
                    Set oSMS = oXMLDom.documentElement
                Else
                    Set oNodeList = oXMLDom.selectNodes("/*/*")
                    For j = 0 To oNodeList.length - 1
                        Set oAppendNode = oSMS.selectSingleNode("//" & oNodeList.Item(j).baseName)
                        If oAppendNode Is Nothing Then
                            oSMS.appendChild oNodeList.Item(j)
                        Else
                            For Each oNode In oNodeList.Item(j).childNodes
                                oAppendNode.appendChild oNode
                            Next oNode
                        End If
                    Next
                End If
                
                oSMS.setAttribute "PROJECT_DB_SERVER_NAME", strServerName
                oSMS.setAttribute "PROJECT_DB_NAME", strSiteName
                
                oFinalDom.appendChild oSMS
                oTempCollection.Add oFinalDom
                Set oSMS = Nothing
            End If
        Next oXMLDom
        For i = 1 To oPlateDomColl.count
            oPlateDomColl.Remove 1
        Next i

        Set pCollection = oTempCollection
    End If
    
    Set oTempCollection = Nothing
    Set oChildDom = Nothing
    Set oNodeList = Nothing
    Set oAppendNode = Nothing
    Set oComparisonRule = Nothing
    Set oPOM = Nothing
    Exit Sub
ErrorHandler:
End Sub

Private Function GetRuleValue(strRuleName As String, strColumn As String) As Variant
Const METHOD = "GetRuleValue"
On Error GoTo ErrorHandler

    Dim oOutputHelper As IJMfgOutputHelper
    Set oOutputHelper = New MfgCatalogQueryHelper
    If Not oOutputHelper Is Nothing Then
        GetRuleValue = oOutputHelper.GetOutputRuleValue(strRuleName, strColumn)
    End If
    Set oOutputHelper = Nothing
Exit Function
ErrorHandler:
End Function
Private Function GetActiveShipClass() As IJProjectRoot
    Const METHOD = "GetActiveShipClass"
    On Error GoTo ErrHandler

    Dim oMfgEntHelper As IJMfgEntityHelper
    Set oMfgEntHelper = New MfgEntityHelper

    Dim oActiveProject As IJProjectRoot
    Set oActiveProject = oMfgEntHelper.GetConfigProjectRoot

    Set GetActiveShipClass = oActiveProject

    Exit Function
ErrHandler:
End Function

Private Function GetTempDirectory() As String
        GetTempDirectory = Environ("TEMP")
        If GetTempDirectory = "" Then
            GetTempDirectory = Environ("TMP")
        End If
End Function

'***********************************************************************************************
' Routine: GetShipInfo
'
' Description:  This routine gets the details of the selected ship
'************************************************************************************************
Private Sub GetShipInfo(strShip As String, strDBName As String, strServerName As String, strConnectionType As String)
    
    Const METHOD = "GetShipInfo"
    On Error GoTo ErrHndlr
    
    Const STRDM = ";"    'delimiters in the connection strings
    Const STRDM2 = "="
    Dim strPrefix As String
    Dim pos As Long
    Dim strDB As String
    Dim oShipClass As IJProjectRoot

    Set oShipClass = GetActiveShipClass
    strShip = oShipClass.Name
    Dim oRegistry As IJRegistry
    Set oRegistry = New Registry
    strDB = oRegistry.GetStringValue(SERVER_REGISTRY_KEY, SERVER_REGISTRY_NAME) 'oDatabase.Path
    
    'Get the Database's physical name.
    pos = InStr(1, strDB, STRDM)
    strServerName = Left$(strDB, (pos - 1))
    strDBName = Mid$(strDB, (pos + 1))
    pos = InStr(1, strDBName, STRDM2)
    strDBName = Mid$(strDBName, (pos + 1))
    'Get the server name
    pos = InStr(1, strServerName, STRDM2)
    strServerName = Mid$(strServerName, (pos + 1))
    Exit Sub
    
ErrHndlr:
End Sub

