VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMergeSymmetry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------
' Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
' Project
'   MfgManagedExportRule
'
' File
'   CMergeSymmetry
'
' Description
'   This file contains the package and delivery functions to handle symmetric parts
'
' Author
'   Nautilus-HSV
'
' History:
'   2016-03-16  Nautilus-HSV  Creation
'----------------------------------------------------------------------------

Option Explicit
Implements IJMfgNestDelivery

Private Const MODULE = "CMergeSymmetry"

Private m_strOutputPath As String
Private m_strBinPath As String
Private m_strNestingServerName As String
Private m_strNestingDBName As String


Private Sub Class_Initialize()

ErrorHandler:
End Sub

Private Sub Class_Terminate()

ErrorHandler:
End Sub

Private Sub IJMfgNestDelivery_PackageAndDeliver(ByVal AssemblyXML As Object, ByVal strTimeStamp As String, ByVal bstrFilePath As String, ByVal bFlag As Boolean)
     Const sMETHOD As String = "IJMfgNestDelivery_PackageAndDeliver"
    On Error GoTo ErrorHandler
        
        ProcessSymmetry AssemblyXML, bstrFilePath
        
    Exit Sub
ErrorHandler:
End Sub

Private Sub ProcessSymmetry(AssemblyXML As Object, bstrFilePath As String)
    Const sMETHOD As String = "ProcessSymmetry"
    On Error GoTo ErrorHandler
    
    Dim oAssemblyXML As IXMLDOMDocument
    Set oAssemblyXML = AssemblyXML
    
    'Find SymmetryPart NodeList
        'CheckPartType is "SYMMETRY" and MASTER is "YES"  and get MasterID and MemberID
        'Compare MasterID XML and  MemberID XML
        'If symmetry,
           ' Set PART_BOARDSIDE = "B"
            'Delete MemberID Attachment
            'Delete MemberID Attachment physical file
            'Update Master.s3d file????
                                    
        'If not symmetry, delete SymmeryPart Node for Master and Member both.
        
    
    Dim oSymmetryPartNodeList As IXMLDOMNodeList
    Set oSymmetryPartNodeList = oAssemblyXML.selectNodes("//SymmetryPart[./PartType = 'SYMMETRY' and ./Master ='YES']")
    
    Dim oRemovedSymmetryPartNodeCollection As New Collection

    
    Dim oFSObject As New FileSystemObject
    Dim oSymmetryPartNode As IXMLDOMNode
    Dim strRevisionLocation As String
    strRevisionLocation = bstrFilePath + "\" + oAssemblyXML.selectSingleNode("//RevisionLog").Text
    'Clean Unnecessary parts
    
    Dim oRevisionDOM As IXMLDOMDocument
    Set oRevisionDOM = New DOMDocument
    Dim bLoadXML As Boolean: bLoadXML = False
    bLoadXML = oRevisionDOM.Load(strRevisionLocation)

    For Each oSymmetryPartNode In oSymmetryPartNodeList
    
        Dim bstrMasterXMLID As String
        Dim bstrMemberXMLID As String
               
        bstrMasterXMLID = oSymmetryPartNode.selectSingleNode("./MasterID").Text
        bstrMemberXMLID = oSymmetryPartNode.selectSingleNode("./MemberID").Text
        
        
        Dim bstrMasterXML As String
        Dim bstrMasterXMLFile As String
        Dim bstrMasterXMLPath As String
        Dim bstrMemberXML As String
        Dim bstrMemberXMLPath As String
        Dim bstrCompareSchemaXML As String
        Dim bstrDataFormat As String
        Dim bstrXPath As String
        
        '''''''''''''''''''''''''''''''''''''''''''
        ' Get MasterXML and MemberXML
        '''''''''''''''''''''''''''''''''''''''''''
        Dim oMasterAttachmentNode As IXMLDOMNode
        Dim oMemberAttachmentNode As IXMLDOMNode
        Dim oMasterPartNode As IXMLDOMNode
        Dim oMasterSymmetryPartNode As IXMLDOMNode
        Dim oMemberPartNode As IXMLDOMNode
        Dim oMemberSymmetryPartNode As IXMLDOMNode
        
        bstrXPath = "//Part[./PartID ='" + bstrMasterXMLID + "']"
        Set oMasterPartNode = oAssemblyXML.selectSingleNode(bstrXPath)
        If Not oMasterPartNode Is Nothing Then
            Set oMasterAttachmentNode = oMasterPartNode.selectSingleNode("./Attachments/Attachment")
        End If
        
        bstrXPath = vbNullString
        bstrXPath = "//Part[./PartID ='" + bstrMemberXMLID + "']"
        Set oMemberPartNode = oAssemblyXML.selectSingleNode(bstrXPath)
        If Not oMemberPartNode Is Nothing Then
            Set oMemberAttachmentNode = oMemberPartNode.selectSingleNode("./Attachments/Attachment")
        End If
        
        If Not oMasterAttachmentNode Is Nothing And Not oMemberAttachmentNode Is Nothing Then
        'MasterXML
            bstrMasterXMLFile = oMasterAttachmentNode.selectSingleNode("./FileName").Text
            bstrMasterXMLPath = oMasterAttachmentNode.selectSingleNode("./Path").Text
            bstrMasterXML = bstrFilePath + "\" + bstrMasterXMLPath + "\" + bstrMasterXMLFile
        
        'MemberXML
            bstrMemberXML = oMemberAttachmentNode.selectSingleNode("./FileName").Text
            bstrMemberXMLPath = oMemberAttachmentNode.selectSingleNode("./Path").Text
            bstrMemberXML = bstrFilePath + "\" + bstrMemberXMLPath + "\" + bstrMemberXML
        
            bstrDataFormat = oMasterAttachmentNode.selectSingleNode("./AttachmentType").Text
        ''''''''''''''''''''''''''''''''''''''''
        ' CompareXML
        '''''''''''''''''''''''''''''''''''''''''
            bstrCompareSchemaXML = GetRuleValue(bstrDataFormat, "String8")
            bstrCompareSchemaXML = GetSharedContentPath() + "\" + bstrCompareSchemaXML
            
            Dim oMasterXMLDoc As New DOMDocument
            Dim oMemberXMLDoc As New DOMDocument
            Dim oCompareSchemaXMLDoc As New DOMDocument
            
            Dim bSucessLoadMasterXML As Boolean
            Dim bSucessLoadMemberXML As Boolean
            Dim bSucessLoadSchemaXML As Boolean
            bSucessLoadMasterXML = oMasterXMLDoc.Load(bstrMasterXML)
            bSucessLoadMemberXML = oMemberXMLDoc.Load(bstrMemberXML)
            bSucessLoadSchemaXML = oCompareSchemaXMLDoc.Load(bstrCompareSchemaXML)
            
            If bSucessLoadMasterXML = True And bSucessLoadMemberXML = True Then
            
                Dim oCompareService As IJCompareService
                Set oCompareService = New CXMLCompareService
                
                Dim lResult As Long
                Dim bstrResult As String
                lResult = oCompareService.CompareXML(oMasterXMLDoc.xml, oMemberXMLDoc.xml, oCompareSchemaXMLDoc.xml, bstrResult)

                Dim strCompareFileName As String
                Dim strMasterPartName As String
                Dim strMemberPartName As String
                If Not oMasterPartNode Is Nothing Then
                    strMasterPartName = oMasterPartNode.selectSingleNode("./Name").Text
                End If
                
                If Not oMemberPartNode Is Nothing Then
                    strMemberPartName = oMemberPartNode.selectSingleNode("./Name").Text
                End If
                
                strMasterPartName = SanitizeFileName(strMasterPartName)
                strMemberPartName = SanitizeFileName(strMemberPartName)
                strCompareFileName = GetTempDirectory() + "\" + "CompareResult_" + strMasterPartName + "_" + strMemberPartName + ".xml"

                Dim oXMLResultDoc As New DOMDocument
                Dim bSucessLoadCompareResult As Boolean
                bSucessLoadCompareResult = oXMLResultDoc.loadXML(bstrResult)
                
                If bSucessLoadCompareResult = True Then
                    Dim oComment As IXMLDOMComment
                    Dim strResult As String
                    strResult = " Comparison Result : " + CStr(lResult)
                    Set oComment = oXMLResultDoc.createComment(strResult)
                    
                    Dim oRootElement As IXMLDOMElement
                    Set oRootElement = oXMLResultDoc.documentElement
                    oXMLResultDoc.insertBefore oComment, oRootElement
                    oXMLResultDoc.save strCompareFileName
                End If
                              

                If lResult = 0 Then 'Symmetry
                    
                    'Set PartBoardSide to "B"
                    Dim oPartBoardSideAttribute As IXMLDOMAttribute
                    Set oPartBoardSideAttribute = oMasterXMLDoc.selectSingleNode("//SMS_PART_INFO/@PART_BOARDSIDE")
                    If Not oPartBoardSideAttribute Is Nothing Then
                        oPartBoardSideAttribute.Value = "B"
                        oMasterXMLDoc.save Replace(bstrMasterXML, ".xml", "_B.xml")
                    End If
                    
                    'Remove MemberXML attachment
                    Dim oMemberAttachmentParentNode As IXMLDOMNode
                    Set oMemberAttachmentParentNode = oMemberAttachmentNode.parentNode
                    
                    Dim strMemberAttachmentID As String
                    strMemberAttachmentID = oMemberAttachmentNode.selectSingleNode("./AttachmentID").Text

                    If Not oMemberAttachmentParentNode Is Nothing Then 'Attachments
                        oMemberAttachmentParentNode.removeChild oMemberAttachmentNode
                        If oMemberAttachmentParentNode.selectNodes("./Attachment").length = 0 Then
                            If Not oMemberPartNode Is Nothing Then
                                oMemberPartNode.removeChild oMemberAttachmentParentNode
                            End If
                        End If
                    End If
                    'Update MasterXML attachment
                    Dim oMasterAttachmentParentNode As IXMLDOMNode
                    Set oMasterAttachmentParentNode = oMasterAttachmentNode.parentNode
                    
                    If Not oMasterAttachmentNode Is Nothing Then 'Attachments
                        oMasterAttachmentNode.selectSingleNode("./FileName").Text = Replace(bstrMasterXMLFile, ".xml", "_B.xml")
                    End If
                    
                    
                    'Delete MemberXML
'                    Dim oFSObject As New FileSystemObject
'                    If oFSObject.FileExists(bstrMemberXML) = True Then
'                            oFSObject.DeleteFile bstrMemberXML
'                    End If

                    If bLoadXML = True Then
                        'Find nodes to remove from revision XML
                        'Only latest revsion export should be adjusted.
                        Dim oMemberRevisionPartNodes As IXMLDOMNodeList
                        
                        Set oMemberRevisionPartNodes = oRevisionDOM.selectNodes("//EXPORT[last()]//ACTION[not(@TYPE='Deletion')]//PART[@PART_GUID='" & strMemberAttachmentID & "']")

                        Dim oMemberRevisionPartNode As IXMLDOMNode
                        For Each oMemberRevisionPartNode In oMemberRevisionPartNodes
                        
                            If Not oMemberRevisionPartNode Is Nothing Then
                            
                                Dim oMemberRevisionPartNodeFileNameAttribute As IXMLDOMAttribute
                                Set oMemberRevisionPartNodeFileNameAttribute = oMemberRevisionPartNode.selectSingleNode("./@FILE_NAME")
                                oMemberRevisionPartNodeFileNameAttribute.Value = Replace(bstrMasterXMLFile, ".xml", "_B.xml")
                            End If

                        Next
                        
                        Dim oMasterRevisionPartNode As IXMLDOMNode
                        Set oMasterRevisionPartNode = oRevisionDOM.selectSingleNode("//EXPORT[last()]//ACTION[not(@TYPE='Deletion')]//PART[@FILE_NAME='" & bstrMasterXMLFile & "']")
                        
                        If Not oMasterRevisionPartNode Is Nothing Then
                        
                            Dim oMasterRevisionPartNodeFileNameAttribute As IXMLDOMAttribute
                            Set oMasterRevisionPartNodeFileNameAttribute = oMasterRevisionPartNode.selectSingleNode("./@FILE_NAME")
                            oMasterRevisionPartNodeFileNameAttribute.Value = Replace(bstrMasterXMLFile, ".xml", "_B.xml")
                        End If
                    End If
                Else 'Not Symmetry
                    'Remove SymmetryPart Node at Master and Member side (by adding them to collection of nodes to be removed.)

                    If Not oMasterPartNode Is Nothing Then
                        Set oMasterSymmetryPartNode = oMasterPartNode.selectSingleNode("./SymmetryPart")
                    
                        If Not oMasterSymmetryPartNode Is Nothing Then
                            oRemovedSymmetryPartNodeCollection.Add oMasterSymmetryPartNode
                            'oMasterPartNode.removeChild oMasterSymmetryPartNode
                        End If
                        If oFSObject.FileExists(Replace(bstrMasterXML, ".xml", "_B.xml")) = True Then
                            oFSObject.DeleteFile Replace(bstrMasterXML, ".xml", "_B.xml")
                        End If
                    End If

                    If Not oMemberPartNode Is Nothing Then
                        Set oMemberSymmetryPartNode = oMemberPartNode.selectSingleNode("./SymmetryPart")
                        If Not oMemberSymmetryPartNode Is Nothing Then
                            oRemovedSymmetryPartNodeCollection.Add oMemberSymmetryPartNode
                            'oMemberPartNode.removeChild oMemberSymmetryPartNode
                        End If
                    End If
                    
                End If
            Else 'If bSucessLoadMasterXML = True And bSucessLoadMemberXML = True Then
            
            End If
        Else 'If Not oMasterAttachmentNode Is Nothing And Not oMemberAttachmentNode Is Nothing Then
            
            If Not oMasterAttachmentNode Is Nothing Then
                 If Not oMasterPartNode Is Nothing Then
                   Set oMasterSymmetryPartNode = oMasterPartNode.selectSingleNode("./SymmetryPart")
                   If Not oMasterSymmetryPartNode Is Nothing Then
                        oRemovedSymmetryPartNodeCollection.Add oMasterSymmetryPartNode
                        'oMasterPartNode.removeChild oMasterSymmetryPartNode
                   End If
                End If
                If oFSObject.FileExists(Replace(bstrMasterXML, ".xml", "_B.xml")) = True Then
                        
                    oFSObject.DeleteFile Replace(bstrMasterXML, ".xml", "_B.xml")
                End If
            End If
                    
            If Not oMemberAttachmentNode Is Nothing Then
                If Not oMemberPartNode Is Nothing Then
                   Set oMemberSymmetryPartNode = oMemberPartNode.selectSingleNode("./SymmetryPart")
                   If Not oMemberSymmetryPartNode Is Nothing Then
                        oRemovedSymmetryPartNodeCollection.Add oMemberSymmetryPartNode
                        'oMemberPartNode.removeChild oMemberSymmetryPartNode
                   End If
                End If
            End If
        End If 'If Not oMasterAttachmentNode Is Nothing And Not oMemberAttachmentNode Is Nothing Then

    Next
    
    'Remove All SymmetryPartNodes
    Dim oRemovedSymmetryPartNode As IXMLDOMNode
    For Each oRemovedSymmetryPartNode In oRemovedSymmetryPartNodeCollection
        Dim oSymParentNode As IXMLDOMNode
        Set oSymParentNode = oRemovedSymmetryPartNode.parentNode
        If Not oSymParentNode Is Nothing Then
            oSymParentNode.removeChild oRemovedSymmetryPartNode
        End If
    Next
    'Save potentially edited revision log file
    If bLoadXML Then
        oRevisionDOM.save strRevisionLocation
    End If
    
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


Private Function GetTempDirectory() As String
        GetTempDirectory = Environ("TEMP")
        If GetTempDirectory = "" Then
            GetTempDirectory = Environ("TMP")
        End If
End Function


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



