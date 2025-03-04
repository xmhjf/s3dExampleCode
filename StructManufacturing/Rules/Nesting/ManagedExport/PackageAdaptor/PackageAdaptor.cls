VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackageAdaptor"
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
   
Private Const REMOVE_PART_PROFILE = "SMS_SCHEMA_PROFILE"
Private Const REMOVE_PART_MEMBER = "SMS_SCHEMA_TEMPLATE"
Private Const REMOVE_PART_TEMPLATE = "SMS_SCHEMA_MEMBER"
   
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
    Dim oAssemblyXML As IXMLDOMDocument2
    Set oAssemblyXML = AssemblyXML
    Dim oAttachmentNodeList As IXMLDOMNodeList
    Dim oAttachmentNode As IXMLDOMNode

    oAssemblyXML.setProperty "SelectionLanguage", "XPath"
    Set oAttachmentNodeList = oAssemblyXML.selectNodes("//Attachment[./Definition = 'SMS_PLATE' or ./Definition ='SMS_PANEL']")

    Dim oFSO As New FileSystemObject
    Dim oFSOFile As File

    Dim strBlockName As String
    Dim strShipName As String
    
    strShipName = oAssemblyXML.selectSingleNode("/Root/ProjectNumber").Text
    strBlockName = oAssemblyXML.selectSingleNode("/Root/Assembly/Name").Text
    
    For Each oAttachmentNode In oAttachmentNodeList
    
        Dim strNestingFilePath As String: strNestingFilePath = ""
        Dim bstrRuleName As String
        
        bstrRuleName = oAttachmentNode.selectSingleNode("./AttachmentType").Text
        On Error Resume Next
            
        strNestingFilePath = GetRuleValue(bstrRuleName, "NestingFilePath")
    
        On Error GoTo ErrorHandler
        If strNestingFilePath = "" Then
            strNestingFilePath = strManagedExportPath
        End If
        
        If Not oFSO.FolderExists(strNestingFilePath) Then
            oFSO.CreateFolder strNestingFilePath
        End If
        
        strNestingFilePath = strNestingFilePath + "\" + strShipName
        If Not oFSO.FolderExists(strNestingFilePath) Then
            oFSO.CreateFolder strNestingFilePath
        End If
        
        strNestingFilePath = strNestingFilePath + "\" + strBlockName
        If Not oFSO.FolderExists(strNestingFilePath) Then
            oFSO.CreateFolder strNestingFilePath
        End If
        
        strNestingFilePath = strNestingFilePath + "\" + "import"
        If Not oFSO.FolderExists(strNestingFilePath) Then
            oFSO.CreateFolder strNestingFilePath
        End If
        
        Dim strFileOrigin As String
        Dim strFileDestination As String
        strFileOrigin = strManagedExportPath + "\" + oAttachmentNode.selectSingleNode("./Path").Text + "\" + oAttachmentNode.selectSingleNode("./FileName").Text
        strFileDestination = strNestingFilePath + "\" + oAttachmentNode.selectSingleNode("./FileName").Text
        If oFSO.FileExists(strFileOrigin) Then
            Dim FileDate As Date
            
            Set oFSOFile = oFSO.GetFile(strFileOrigin)
            FileDate = oFSOFile.DateLastModified
            If (oExportStartDate) < FileDate Then
                oFSO.CopyFile strFileOrigin, strFileDestination, True
            End If
        End If
    Next oAttachmentNode
    
    'Copy Revision file - Export Revision and Export Deletion
    Dim strRevisionFileOrigin As String
    strRevisionFileOrigin = strManagedExportPath + "\" + oAssemblyXML.selectSingleNode("//RevisionLog").Text
    
    Dim oRevisionDOM As IXMLDOMDocument
    Set oRevisionDOM = New DOMDocument60
    
    Dim oRevisionDOM2 As IXMLDOMDocument2
    Set oRevisionDOM2 = oRevisionDOM
    oRevisionDOM2.setProperty "SelectionLanguage", "XPath"
    Set oRevisionDOM2 = Nothing
    
    Dim bLoadXML As Boolean: bLoadXML = False
    If oRevisionDOM.Load(strRevisionFileOrigin) = True Then
    
    ' 1. Remove unnecessary PART - Profile, Member, and Template
        Dim oRemovePartNodes As IXMLDOMNodeList
        Dim strRemoveXPath As String
        strRemoveXPath = "//PART[@OUTPUT_TYPE =" & "'" & REMOVE_PART_PROFILE & "'" & "or"
        strRemoveXPath = strRemoveXPath & " @OUTPUT_TYPE =" & "'" & REMOVE_PART_MEMBER & "'" & "or"
        strRemoveXPath = strRemoveXPath & " @OUTPUT_TYPE =" & "'" & REMOVE_PART_TEMPLATE & "']"
    
        Set oRemovePartNodes = oRevisionDOM.selectNodes(strRemoveXPath)
    
        Dim oRemovePartNode As IXMLDOMNode
        For Each oRemovePartNode In oRemovePartNodes
            Dim oParentNode As IXMLDOMNode
            Set oParentNode = oRemovePartNode.parentNode
    
            If Not oParentNode Is Nothing Then
                oParentNode.removeChild oRemovePartNode
            End If
            Set oRemovePartNode = Nothing
        Next
    
    '2. Get All recent Export Part and Copy Revision log and deletion log file
    
        On Error Resume Next
        Dim oRevisionDeleteDOM As IXMLDOMDocument
        Set oRevisionDeleteDOM = GetDeleteRevisionDocument(oRevisionDOM)
        On Error GoTo ErrorHandler
    
        Dim oRecentExportPartList As IXMLDOMNodeList
        Set oRecentExportPartList = oRevisionDOM.selectNodes("//EXPORT[last()]//PART")
        
        Dim oPartNode As IXMLDOMNode
        Dim strExportedRevisionPath As String: strExportedRevisionPath = strManagedExportPath
        For Each oPartNode In oRecentExportPartList
        
            Dim oDataFormatNode As IXMLDOMNode
            Set oDataFormatNode = oPartNode.selectSingleNode("./@DATA_FORMAT")
            
            Dim strRevisionPath As String: strRevisionPath = strManagedExportPath
            If Not oDataFormatNode Is Nothing Then
                Dim strRevisionNestingFilePath As String: strRevisionNestingFilePath = ""
                strRevisionNestingFilePath = GetRuleValue(oDataFormatNode.nodeValue, "NestingFilePath")

                If strRevisionNestingFilePath <> "" Then
                    strRevisionPath = strRevisionNestingFilePath
                End If
            End If
          
            If Not oFSO.FolderExists(strRevisionPath) Then
                oFSO.CreateFolder strRevisionPath
            End If
            
            strRevisionPath = strRevisionPath + "\" + strShipName
            If Not oFSO.FolderExists(strRevisionPath) Then
                oFSO.CreateFolder strRevisionPath
            End If
            
            strRevisionPath = strRevisionPath + "\" + strBlockName
            If Not oFSO.FolderExists(strRevisionPath) Then
                oFSO.CreateFolder strRevisionPath
            End If
            
            strRevisionPath = strRevisionPath + "\" + "import"
            If Not oFSO.FolderExists(strRevisionPath) Then
                oFSO.CreateFolder strRevisionPath
            End If

            strRevisionPath = strRevisionPath + "\" + "Revision"
            If Not oFSO.FolderExists(strRevisionPath) Then
                oFSO.CreateFolder strRevisionPath
            End If
            strRevisionPath = strRevisionPath + "\" + strShipName + "_" + strBlockName

            If Not oFSO.FolderExists(strRevisionPath) Then
                oFSO.CreateFolder strRevisionPath
            End If
            
            If VBA.StrComp(strRevisionPath, strExportedRevisionPath, vbTextCompare) <> 0 Then
                strExportedRevisionPath = strRevisionPath
         
                oFSO.CopyFile strRevisionFileOrigin, strRevisionPath + "\" + "Export_Revision.xml", True
                If Not oRevisionDeleteDOM Is Nothing Then
                    oRevisionDeleteDOM.save strRevisionPath + "\" + "Export_Deletion.xml"
                End If
            End If
        Next
    End If

    Exit Function
ErrorHandler:
    'MsgBox "error handler"
End Function

Private Function GetDeleteRevisionDocument(ByVal oRevisionDOM As IXMLDOMDocument) As DOMDocument60
Const METHOD = "GetDeleteRevisionDocument"
On Error GoTo ErrorHandler

    Dim oRevisionDeleteDOM As IXMLDOMDocument
    Set oRevisionDeleteDOM = New DOMDocument60
    
    Dim oRevisonDeleteDOM2 As IXMLDOMDocument2
    Set oRevisonDeleteDOM2 = oRevisionDeleteDOM
    oRevisonDeleteDOM2.setProperty "SelectionLanguage", "XPath"
    Set oRevisonDeleteDOM2 = Nothing
    
    Dim oNotDeleteNodes As IXMLDOMNodeList
    Dim oNotDeleteNode As IXMLDOMNode
    oRevisionDeleteDOM.loadXML oRevisionDOM.xml
    Set oNotDeleteNodes = oRevisionDeleteDOM.selectNodes("//EXPORT[not(./ACTION[@TYPE='Deletion'])]")
    For Each oNotDeleteNode In oNotDeleteNodes
        oNotDeleteNode.parentNode.removeChild oNotDeleteNode
    Next oNotDeleteNode
    
    Set oNotDeleteNodes = oRevisionDeleteDOM.selectNodes("//ACTION[not(@TYPE='Deletion')]")
    For Each oNotDeleteNode In oNotDeleteNodes
        oNotDeleteNode.parentNode.removeChild oNotDeleteNode
    Next oNotDeleteNode
    
    Set GetDeleteRevisionDocument = oRevisionDeleteDOM

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

