VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPAdjustment"
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
'   2015-06-24  K. Kamph     Creation date
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

Private Sub IJMfgNestDelivery_PackageAndDeliver(ByVal AssemblyXML As Object, ByVal strTimeStamp As String, ByVal bstrFilePath As String, ByVal bFlag As Boolean)
     Const sMETHOD As String = "IJMfgNestDelivery_PackageAndDeliver"
    On Error GoTo ErrorHandler
        'MsgBox "TEST"
        m_bIndividualFile = bFlag
    
    'Copy Files to ALMA location
    SendDocumentsToNesting AssemblyXML, bstrFilePath, strTimeStamp
    Exit Sub
ErrorHandler:
End Sub

Private Function SendDocumentsToNesting(AssemblyXML As Object, strManagedExportPath As String, ByVal strTimeStamp As String) As String()
    Const sMETHOD As String = "SendDocumentsToNesting"
    On Error GoTo ErrorHandler
    
    Dim oExportStartDate As Date
    Dim oExportStartTime As Date
    strTimeStamp = Replace(strTimeStamp, "T", " ")
    strTimeStamp = Replace(strTimeStamp, "Z", "")
    oExportStartDate = DateValue(strTimeStamp)
    oExportStartTime = TimeValue(strTimeStamp)
    oExportStartDate = oExportStartDate + oExportStartTime
    'oExportStart strTimeStamp
    Dim oAssemblyXML As IXMLDOMDocument
    Set oAssemblyXML = AssemblyXML
    oAssemblyXML.preserveWhiteSpace = True
    Dim oAttachmentNodeList As IXMLDOMNodeList
    Dim oAttachmentNode As IXMLDOMNode
    Dim strModelGUID As String
    Dim strModelName As String

    Set oAttachmentNodeList = oAssemblyXML.selectNodes("//Attachment[../../RelatedPart/PartType = 'COMMON']")
    For Each oAttachmentNode In oAttachmentNodeList
        Dim oPartXML As IXMLDOMDocument
        Set oPartXML = New DOMDocument
        Dim strFileOrigin As String
        strFileOrigin = strManagedExportPath + "\" + oAttachmentNode.selectSingleNode("./Path").Text + "\" + oAttachmentNode.selectSingleNode("./FileName").Text
        'oPartXML.preserveWhiteSpace = True
        oPartXML.Load strFileOrigin
        Dim oModelPartGuid As IXMLDOMNode
        Set oModelPartGuid = oPartXML.selectSingleNode("//SMS_PART_INFO/@MODEL_PART_GUID")
        If Not oModelPartGuid Is Nothing Then
            strModelGUID = oAttachmentNode.selectSingleNode("../../PartID").Text
            strModelGUID = Replace(strModelGUID, "{", "")
            strModelGUID = Left(strModelGUID, InStr(strModelGUID, "}") - 1)
            oModelPartGuid.nodeValue = strModelGUID
        End If
        Set oModelPartGuid = oPartXML.selectSingleNode("//SMS_PART_INFO/@MODEL_PART_NAME")
        If Not oModelPartGuid Is Nothing Then
            strModelName = oAttachmentNode.selectSingleNode("../../Name").Text
            oModelPartGuid.nodeValue = strModelName
        End If
        Set oModelPartGuid = oPartXML.selectSingleNode("//SMS_PART_INFO/SMS_COMMON_PART_INFO/@PART_BLOCK")
        If Not oModelPartGuid Is Nothing Then
            oModelPartGuid.nodeValue = ""
        End If
        Set oModelPartGuid = oPartXML.selectSingleNode("//SMS_PART_INFO/SMS_COMMON_PART_INFO/@PART_GUID")
        If Not oModelPartGuid Is Nothing Then
            oModelPartGuid.nodeValue = ""
        End If
        Set oModelPartGuid = oPartXML.selectSingleNode("//SMS_PART_INFO/SMS_COMMON_PART_INFO/@COMMON_PART_NAME")
        If Not oModelPartGuid Is Nothing Then
            oModelPartGuid.nodeValue = ""
        End If
        Set oModelPartGuid = oPartXML.selectSingleNode("//SMS_PART_INFO/SMS_COMMON_PART_INFO/@QUANTITY")
        If Not oModelPartGuid Is Nothing Then
            oModelPartGuid.nodeValue = "1"
        End If
        Dim oCommonPartList As IXMLDOMNodeList
        Dim oCommonPart As IXMLDOMNode
        Set oCommonPartList = oPartXML.selectNodes("//SMS_COMMON_PART")
        For Each oCommonPart In oCommonPartList
            oCommonPart.parentNode.removeChild oCommonPart
        Next oCommonPart
        oPartXML.save strFileOrigin
    Next oAttachmentNode
    
    Set oAttachmentNodeList = oAssemblyXML.selectNodes("//RelatedPart")
    For Each oAttachmentNode In oAttachmentNodeList
        oAttachmentNode.parentNode.removeChild oAttachmentNode
    Next oAttachmentNode
    oAssemblyXML.save strManagedExportPath + "\SmartProductionFull.xml"
    
    
'    Dim oFSO As New FileSystemObject
'    Dim oFSOFile As File
'
'    For Each oAttachmentNode In oAttachmentNodeList
'        Dim bstrRuleName As String
'        bstrRuleName = oAttachmentNode.selectSingleNode("./AttachmentType").Text
'        m_strOutputPath = GetRuleValue(bstrRuleName, "NestingFilePath")
'        Dim strFileOrigin As String
'        Dim strFileDestination As String
'        strFileOrigin = strManagedExportPath + "\" + oAttachmentNode.selectSingleNode("./Path").Text + "\" + oAttachmentNode.selectSingleNode("./FileName").Text
'        strFileDestination = m_strOutputPath + "\" + oAttachmentNode.selectSingleNode("./FileName").Text
'        If oFSO.FileExists(strFileOrigin) Then
'            Dim FileDate As Date
'
'            Set oFSOFile = oFSO.GetFile(strFileOrigin)
'            FileDate = oFSOFile.DateLastModified
'            If (oExportStartDate) < FileDate Then
'                'oFSOFile.DateLastModified
'                oFSO.CopyFile strFileOrigin, strFileDestination, True
'            End If
'        End If
'        'MsgBox "strFileOrigin: " & strFileOrigin & vbNewLine & "strFileDestination: " & strFileDestination
'    Next oAttachmentNode
    
    Exit Function
ErrorHandler:
End Function

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


Private Function GetTempDirectory() As String
        GetTempDirectory = Environ("TEMP")
        If GetTempDirectory = "" Then
            GetTempDirectory = Environ("TMP")
        End If
End Function

