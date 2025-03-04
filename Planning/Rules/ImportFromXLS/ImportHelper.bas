Attribute VB_Name = "ImportHelper"
'**************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'  Project     : Planning\Data\Rules\XLSImport\
'  File        : ImportHelper.bas
'
'  Description :
'
'  History     :
'   6th Feb 2012      Siva     Initial creation
'**************************************************************************************

Option Explicit

Private Const MODULE = "ImportHelper"
Private Const IID_IJAssemblyChild   As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const IID_IJSTXPartInfo     As String = "'3f005307-7b84-460b-a5d7-e1140040ff00'"

Private Const RELATION_DELETED = &H1000000

Private iRow                As Integer
Private iColumn             As Integer
Private giFilePointer       As Integer
Private iRowStart           As Long
Private iRowEnd             As Long
Private iColumnAttrType     As Long

Private m_oExcelApp     As Object
Private m_oWorkbook     As Object
Private m_oWorksheet    As Object

Private m_AssyHierarchyCol          As Collection
Private m_CheckDuplicatedSystemsCol As Collection
Private m_oPgWriteAccColl           As Collection
Private m_AttributeValColl          As Collection
Private m_AttributeNameColl         As Collection
Private m_oParent                   As GSCADAssembly.IJAssembly
Private m_RootAssemblyBlock         As GSCADAssembly.IJAssembly
Private m_ConfigProjectRoot         As IJConfigProjectRoot
    
Private Const ASSEMBLY_HIERARCHY As String = "ASSEMBLY_HIERARCHY"
Private Const SYSTEM1_COLUMN = 2
Private Const TYPE_KEY = "TYPE"
Private Const START_KEY = "Start"
Private Const END_KEY = "End"

Private Function IsPartAssignableToAssembly(oPart As IJAssemblyChild, oAssembly As GSCADAssembly.IJAssembly) As Boolean

    Dim oParent         As Object
    Dim oIJDObject      As IJDObject
    
    IsPartAssignableToAssembly = False
    
    Set oIJDObject = oPart
    'Check approval status and acess rights on the part and assembly
    If (oIJDObject.ApprovalStatus <> Working) Or (oIJDObject.AccessControl And acUpdate) <> acUpdate Then
        Exit Function
    Else
        Set oIJDObject = Nothing
        Set oIJDObject = oAssembly
        If (oIJDObject.ApprovalStatus <> Working) Or (oIJDObject.AccessControl And acUpdate) <> acUpdate Then
            Exit Function
        End If
    End If
    
    Set oParent = oPart.Parent
    Set oIJDObject = Nothing
    
    'Part has a parent
    If Not oParent Is Nothing Then
        
        'check approval status and acess rights on the part's parent
        Set oIJDObject = oParent
        If (oIJDObject.ApprovalStatus <> Working) Or (oIJDObject.AccessControl And acUpdate) <> acUpdate Then
            Exit Function
        'Part assignable only when its parent is Failed Parts folder or,
        'Unassigned parts folder or Project root.
        ElseIf TypeOf oParent Is IJPlnFailedParts Or TypeOf oParent Is IJPlnUnprocessedParts Or TypeOf oParent Is IJConfigProjectRoot Then
            IsPartAssignableToAssembly = True
        Else
            IsPartAssignableToAssembly = False
        End If
    Else
        'Part has no parent, hence assignable
        IsPartAssignableToAssembly = True
    End If
    
    Set oIJDObject = Nothing

End Function

'**************************************************************************************
' Description: ImportAssemblyHierarchy(DataFilePath As String)
'**************************************************************************************
Public Sub ImportAssemblyHierarchy(bstrWorkBook As String, bstrWorkSheet As String, bstrLogFile As String, bstrRootAssy As String, pRootAssembly As Object)
    Const MethodStr = "ImportAssemblyHierarchy"
    On Error GoTo ReadExcelError
    
    Dim AssyInfo()      As String
    Dim AssyPath        As String
    Dim AssemblyName    As String
    Dim AssemblyType    As String
    Dim strErrMsg       As String
    Dim oNewAssy        As GSCADAssembly.IJAssembly
    Dim tempParent      As GSCADAssembly.IJAssembly
    
    'Variables for iteration
    Dim i       As Integer
    Dim j       As Integer
    
    If giFilePointer <> 0 Then
        Close #giFilePointer
        giFilePointer = 0
    End If
    
    giFilePointer = FreeFile
    
    ' open log file for error reporting
    Open bstrLogFile For Append As #giFilePointer
    Print #giFilePointer, ""
    
    ' Read excel data from the work sheet
    ReadExcelData bstrWorkBook, bstrWorkSheet
        
    ' If no valid data in the  work sheet then exit
    If m_AssyHierarchyCol Is Nothing Then
        strErrMsg = "Assembly hierarchy coll is nothing"
        Print #giFilePointer, "Assembly hierarchy coll is nothing"
        GoTo ReadExcelError
    End If
    
    If Not pRootAssembly Is Nothing Then
        If TypeOf pRootAssembly Is GSCADAssembly.IJAssembly Then
            Set m_RootAssemblyBlock = pRootAssembly
        Else
            Print #giFilePointer, "Root assembly name specified is invalid"
            Exit Sub
        End If
    ElseIf Not bstrRootAssy = "" Then
        ' get the root assembly object from name.
        Dim oRootAssy As Object
        Set oRootAssy = GetObjectFromName(bstrRootAssy)
        
        If Not oRootAssy Is Nothing Then
            If TypeOf oRootAssy Is GSCADAssembly.IJAssembly Then
                Set m_RootAssemblyBlock = oRootAssy
            End If
        Else
            Print #giFilePointer, "Root assembly name specified is invalid"
            Exit Sub
        End If
    Else
        Print #giFilePointer, "No root assembly specified"
        Exit Sub
    End If
    
    ' Create Root BO if not exist.
    PlanningInitialize
    
    ' Get write access Permission Groups
    Set m_oPgWriteAccColl = GetAllPermissionGroupsOfThePlant()
    
    'Create System Hierarchy
    For i = 1 To m_AssyHierarchyCol.Count
        
        AssyInfo = m_AssyHierarchyCol.Item(i)
        AssyPath = vbNullString
        
        For j = 1 To UBound(AssyInfo) - 1
            AssemblyName = AssyInfo(j)

            AssyPath = AssyPath & "\" & AssemblyName
            
            Set m_oParent = Nothing
            
            'Set parent system to create system
            If tempParent Is Nothing Then
                Set m_oParent = m_RootAssemblyBlock
            Else
                Set m_oParent = tempParent
            End If
            
            'Check AssemblyName is existing system or not
            Set oNewAssy = CheckAssemblyExists(AssemblyName)
            
            If Not oNewAssy Is Nothing Then
                
                Print #giFilePointer, AssyPath & " already exist"
                Set tempParent = Nothing
                Set tempParent = oNewAssy
                Set oNewAssy = Nothing
            Else
                AssemblyType = AssyInfo(0)
                
                ' Create new assembly
                Set oNewAssy = CreateAssembly(AssemblyName)
                
                ' Set the attributes on the new assembly
                SetAttributesOnAssembly oNewAssy, m_AttributeValColl.Item(i)
                Print #giFilePointer, AssyPath & " created"
                
                Set tempParent = Nothing
            End If

            If (Not oNewAssy Is Nothing) Then
                'On Error Resume Next    ' in case system has already been added
                'oAlreadyDisplayedMsgForSystemsCol.Add oNewAssy
                'On Error GoTo ImportSystemsError
            End If
        Next j
        
        Set tempParent = Nothing
    Next i
    
cleanup:
    'Release all objects
    m_oExcelApp.Quit
    Set m_AssyHierarchyCol = Nothing
    Set m_RootAssemblyBlock = Nothing
    Set tempParent = Nothing
    Set m_oExcelApp = Nothing
    Set m_oWorkbook = Nothing
    Set m_oWorksheet = Nothing
    Set m_CheckDuplicatedSystemsCol = Nothing
    Set m_AttributeValColl = Nothing
    Set m_AttributeNameColl = Nothing
    Set m_oParent = Nothing
    Set m_ConfigProjectRoot = Nothing
    Set m_oPgWriteAccColl = Nothing
    Close #giFilePointer
    
    Exit Sub

ReadExcelError:
    Print #giFilePointer, strErrMsg
    Print #giFilePointer, AssyPath & " Import Assembly Hierarchy  failed"
    GoTo cleanup
    Exit Sub
End Sub

' Retrieves the object by name directly from the database.  Assumes unique names
' If name is not unique then the first object is retrieved
Private Function GetObjectFromName(strName As String) As Object
    On Error GoTo ErrorHandler
    
    Dim oCommand            As JMiddleCommand
    Dim oAllObjects         As IEnumMoniker
    Dim oPOM                As IJDPOM
   
    Set oCommand = New JMiddleCommand

    Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    ' Get all the objects matches with input name from the model through query
    With oCommand
        .Prepared = False
        .QueryLanguage = LANGUAGE_SQL
        .ActiveConnection = oPOM.DatabaseID
        .CommandText = "SELECT oid FROM CORENamedItem WHERE strName ='" & strName & "'"
        
        Set oAllObjects = .SelectObjects()
    End With
    
    Set oCommand = Nothing
    
    Dim oObjVariant     As Variant
    Dim oObj            As Object

    ' Get the first object in the collection.
    If Not oAllObjects Is Nothing Then
        For Each oObjVariant In oAllObjects
            
            On Error Resume Next
            
            Set oObj = oPOM.GetObject(oObjVariant)
            If Not oObj Is Nothing Then
               Set GetObjectFromName = oObj
               Exit For
            End If
            
            Set oObj = Nothing
        Next
    End If
    
    Set oObj = Nothing
    Set oAllObjects = Nothing
    Set oObjVariant = Nothing
    Set oPOM = Nothing

Exit Function
ErrorHandler:
  ' log the error and pass it along so the ATP can catch it
    Print #giFilePointer, "GetObjectFromName failed"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub SetAttributesOnAssembly(oAssyBase As IJAssemblyBase, oAttributeValColl As Collection)
    Const METHOD = "SetAttributesOnAssembly"
    On Error GoTo ErrorHandler
    
    ' Get collection of meta data attributes(on occurrence)
    Dim oAttrColl   As Collection
    Set oAttrColl = GetCollectionOfAttributesFromObject(oAssyBase)
    
    Dim lCodeList   As Long
    Dim iIndex      As Integer
    
    For iIndex = 1 To oAttributeValColl.Count
        If Not oAttributeValColl.Item(iIndex) = "" Then
            ' work center is not attribute on assembly so special case
            If UCase(m_AttributeNameColl.Item(iIndex)) = "WORKCENTER" Then  ' Work Center
            
                Dim oPlnAssociateCatalog        As IJPlnAssociateCatalog
                Set oPlnAssociateCatalog = oAssyBase
                oPlnAssociateCatalog.SetWorkcenter oAttributeValColl.Item(iIndex)
                
                Set oPlnAssociateCatalog = Nothing
                
            ' Permission group is not attribute on assembly so special case
            ElseIf UCase(m_AttributeNameColl.Item(iIndex)) = "PERMISSION_GROUP" Then    ' Permission group
                If Not oAttributeValColl.Item(iIndex) = "" Then
                    Dim lConditionID     As Long
                    If CheckInputPerGrpIsValid(oAttributeValColl.Item(iIndex), lConditionID) = True Then
                        Dim oIJDObject As IJDObject
                        Set oIJDObject = oAssyBase
                        oIJDObject.PermissionGroup = lConditionID
                        
                        Set oIJDObject = Nothing
                    Else
                        Print #giFilePointer, "Invalid Permission group " & oAttributeValColl.Item(iIndex)
                    End If
                End If
                
            Else    'Metadata attributes read and set on assembly
                
                Dim jIndex      As Integer
                For jIndex = 1 To oAttrColl.Count
                    
                    Dim oAttribute As IJDAttribute
                    Set oAttribute = oAttrColl.Item(jIndex)
                    
                    If UCase(oAttribute.AttributeInfo.Name) = UCase(m_AttributeNameColl.Item(iIndex)) Then
                        If oAttribute.AttributeInfo.ReadOnly = False Then ' check write access
                        
                            Dim strCodelListName As String
                            strCodelListName = oAttribute.AttributeInfo.CodeListTableName
                            If Not strCodelListName = "" Then
                                lCodeList = GetCodelistNumberFromLongValue(oAttributeValColl.Item(iIndex), strCodelListName)
                                
                                If Not lCodeList = -1 Then
                                    oAttribute.Value = lCodeList
                                End If
                            Else
                                '    Type values
                                '    Char = 1
                                '    LONG = 4
                                '    Short = 5
                                '    Float = 7
                                '    DOUBLE = 8
                                
                                Dim lAttrType As Long
                                lAttrType = oAttribute.AttributeInfo.Type
    
                                If lAttrType = 1 Then
                                    oAttribute.Value = oAttributeValColl.Item(iIndex)
                                ElseIf lAttrType = 4 Or lAttrType = 5 Then
                                    oAttribute.Value = CLng(oAttributeValColl.Item(iIndex))
                                ElseIf lAttrType = 7 Or lAttrType = 8 Then
                                        oAttribute.Value = CDbl(oAttributeValColl.Item(iIndex))
                                End If
                            End If
                            
                        End If
                    End If
                    
                    Set oAttribute = Nothing
                Next
            End If
        End If
    Next
    
    Set oAttrColl = Nothing
    Exit Sub
    
ErrorHandler:
    Print #giFilePointer, "SetAttributesOnAssembly failed"
    ' Err.Raise Err.Number, MODULE, METHOD
    Exit Sub
End Sub

''**************************************************************************************
'' Routine      : GetCodelistNumberFromPrimerCode
'' Abstract     : This Function gives the code list value of the Primer by taking table name and
''                Short description of the code list as input
''
''**************************************************************************************
Private Function GetCodelistNumberFromLongValue(ByVal strLongVal As String, ByVal strTableName As String) As Long
    Const METHOD = "GetCodelistNumberFromLongValue"
    On Error GoTo ErrorHandler
    
    ' If the code is null string then return it as undefined
    If strLongVal = vbNullString Then
        GetCodelistNumberFromLongValue = -1
        Exit Function
    End If
    
    Dim oCodeValue          As IJDCodeListValue
    Dim strValue            As String
    Dim lInfoCount          As Long
    Dim iIndex              As Integer
    Dim lCodelistValue      As Long
    
    lCodelistValue = -1     'Initialize with Unknown Value

    
    Dim oCodeListMetaData   As IJDCodeListMetaData
    
    Set oCodeListMetaData = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    If Not oCodeListMetaData Is Nothing Then
        Dim oInfoColl As IJDInfosCol
        Set oInfoColl = oCodeListMetaData.CodelistValueCollection(strTableName)
        
        If Not oInfoColl Is Nothing Then
            lInfoCount = oInfoColl.Count
            For iIndex = 1 To lInfoCount
                Set oCodeValue = oInfoColl.Item(iIndex)
                'Getting the Short string Value of the Code List
                strValue = oCodeValue.LongValue
                If StrComp(strLongVal, strValue) = 0 Then     'Checking whether both the strings are same
                    'Assigning the Corresponding Code List Value
                    lCodelistValue = oCodeValue.ValueID
                    Exit For
                 End If
            Next iIndex
        End If
    End If
    
    GetCodelistNumberFromLongValue = lCodelistValue
    
    Exit Function
    
ErrorHandler:
    Print #giFilePointer, "GetCodelistNumberFromLongValue failed"
    Err.Raise Err.Number, MODULE, METHOD
    Exit Function
End Function

Private Sub PlanningInitialize()
    Const METHOD = "PlanningInitialize"
    On Error GoTo ErrorHandler
    
    Dim bCommit                     As Boolean
    Dim oPlnIntHelper               As IJDPlnIntHelper
    Dim oBlock                      As GSCADBlock.IJBlock
    Dim oBlockFactory               As GSCADBlock.IJBlockFactory
    Dim objAssembly                 As GSCADAssembly.IJAssemblyBase
    Dim objSRDQuery                 As IJSRDQuery
    Dim ListOfWorkcenters()         As String
    Dim objWorkcenterQuery          As IJSRDWorkcenterQuery
    Dim oPlnAssociateCatalog        As IJPlnAssociateCatalog

    bCommit = False
   
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    If oResourceManager Is Nothing Then Exit Sub
    
    Set oPlnIntHelper = New CPlnIntHelper
    Set oBlock = oPlnIntHelper.GetTopLevelBlock(oResourceManager)
    
    If oBlock Is Nothing Then
        bCommit = True
        ' -------------------------------------------------
        ' If no assembly tree root exists, create a new one
        ' -------------------------------------------------
        
        Print #giFilePointer, "Top level bock doesn't exist, creating it now"
        
        ' Get active ship class from Project Management
        m_ConfigProjectRoot = GetConfigProjectRootFromDB()
        ' Get top-level block
        
        Set oBlockFactory = New GSCADBlock.BlockFactory
        Set oBlock = oBlockFactory.CreateTopLevelBlock(oResourceManager, m_ConfigProjectRoot)
        
         ' set Assembly type
        Set objAssembly = oBlock
        objAssembly.Type = 2
        
        Set oPlnAssociateCatalog = oBlock
        Set objSRDQuery = New SRDQuery
        
        ' Fetch the SRDWorkcenterQuery (persisted in the catalog db) to perform SQL queries on the catalog database
        Set objWorkcenterQuery = objSRDQuery.GetWorkcenterQuery()
        
        'get list of all workcenter
        objWorkcenterQuery.GetAllWorkcenters ListOfWorkcenters
        
        'check about the workcenter is bulkload
        If UBound(ListOfWorkcenters) - LBound(ListOfWorkcenters) >= 0 Then
            'add the default workcenter
            oPlnAssociateCatalog.SetWorkcenter ListOfWorkcenters(0)
        End If
        
        ' Clean up (before transaction ends)
        Set objSRDQuery = Nothing
        Set objWorkcenterQuery = Nothing
        Set oBlockFactory = Nothing
        Set objAssembly = Nothing
        Set oPlnAssociateCatalog = Nothing
        
    End If
    
    ' Clean up (before transaction ends)
    Set oPlnIntHelper = Nothing
    Set m_ConfigProjectRoot = Nothing
    Set oResourceManager = Nothing

    ' if root assembly is not not supplied
    If m_RootAssemblyBlock Is Nothing Then
        Set m_RootAssemblyBlock = oBlock
    End If

'    ' --------------
'    ' Commit Changes
'    ' --------------
'
'    If bCommit Then
'        Dim oTransactionMgr As IMSTransactionManager.IJTransactionMgr
'        Set oTransactionMgr = oTrader.Service(TKTransactionMgr, "")
'        oTransactionMgr.Commit "PlanningUEService.PlanningInitialize()"
'        Set oTransactionMgr = Nothing
'    End If
    
'    Set oPlnIntHelper = New CPlnIntHelper
'    Set m_RootAssemblyBlock = oPlnIntHelper.GetTopLevelBlock(oConnection.ResourceManager)
'
cleanup:
    Exit Sub
    
ErrorHandler:
    Print #giFilePointer, "PlanningInitialize failed"
    Err.Raise Err.Number, MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description: Read system Hierarchy from DataFile and
'             Create rsASSEMBLYHIERARCHY
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReadExcelData(bstrWorkBook As String, bstrWorkSheet As String)
    Const MethodStr = "ReadExcelData"
    On Error GoTo ReadExcelDataError
    
    Dim strDataType     As String
    Dim lSheetCount     As Long
    Dim Index           As Long
    
    Set m_oExcelApp = excel.Application
    
    ' Disable the display alerts, in order to avoid displaying "open in read-only mode" etc..
    m_oExcelApp.DisplayAlerts = False

    'Open Workbook named DataFile
    Set m_oWorkbook = m_oExcelApp.Workbooks.Open(bstrWorkBook)

    'Restore the display alerts flag
    m_oExcelApp.DisplayAlerts = True
    
    On Error Resume Next
    
    For lSheetCount = 1 To m_oWorkbook.Worksheets.Count
        Set m_oWorksheet = m_oWorkbook.Worksheets(lSheetCount)
        m_oWorksheet.Activate
        
        If m_oWorksheet.Name = bstrWorkSheet Then
            ' Read the assembly hierarchy sheet
            ReadASSEMBLYHIERARCHY
        End If
    Next
    
Exit Sub

ReadExcelDataError:
    
    Print #giFilePointer, "Error in reading the excel sheet data of " & bstrWorkSheet & " in " & bstrWorkBook
    
    'Release worksheet
    If Not m_oWorksheet Is Nothing Then Set m_oWorksheet = Nothing
    'Release workbook
    If Not m_oWorkbook Is Nothing Then
        Call m_oWorkbook.Close(False)
        Set m_oWorkbook = Nothing
    End If
    
End Sub

'**************************************************************************************
''  Sub             : ReadASSEMBLYHIERARCHY
''  Description     : Read system hierarchy from given data files
''
'**************************************************************************************
Private Sub ReadASSEMBLYHIERARCHY()
    Const MethodStr = "ReadASSEMBLYHIERARCHY"
    On Error GoTo ReadASSEMBLYHIERARCHYError
    
    Dim AssemblyElem()  As String
    Dim AssemblyType    As String
    Dim tempElem()      As String
    Dim AssyPath        As String
    Dim bRowIsBlank     As Boolean
    Dim iNdx            As Long
    
    If m_CheckDuplicatedSystemsCol Is Nothing Then
        Set m_CheckDuplicatedSystemsCol = New Collection
    End If
    
    If Not InStr(m_oWorksheet.Name, "ASSEMBLY_HIERARCHY") > 0 Then
        Exit Sub
    End If
    
    GetRanges   'Get range for iRowStart, iRowEnd, iColumnAttrType, etc.
    
    Dim systemIndexInLogFile As Long
    systemIndexInLogFile = 1
       
    'Initialize System hierarchy collection
    Set m_AssyHierarchyCol = New Collection
    Set m_AttributeValColl = New Collection
    Set m_AttributeNameColl = New Collection
    
    ReDim AssemblyElem(1)
    
    'Print #giFilePointer, m_strASSEMBLYHIERARCHYData
    
    For iRow = iRowStart + 1 To iRowEnd - 1
    
        'Skip blank rows
        bRowIsBlank = True
        For iColumn = SYSTEM1_COLUMN To iColumnAttrType
            If Not m_oWorksheet.Cells.Item(iRow, iColumn) = vbNullString Then
                bRowIsBlank = False
                Exit For
            End If
        Next iColumn
        If Not bRowIsBlank Then
        
            'Save system type
            AssemblyElem(0) = Trim(m_oWorksheet.Cells.Item(iRow, iColumnAttrType))
            AssemblyType = AssemblyElem(0)
            
            'Save system name to array
            For iColumn = SYSTEM1_COLUMN To (iColumnAttrType - 1)
                If Not m_oWorksheet.Cells.Item(iRow, iColumn) = vbNullString Then
                    ReDim Preserve AssemblyElem(iColumn)
                    AssemblyElem(iColumn - 1) = Trim(m_oWorksheet.Cells.Item(iRow, iColumn))
                    AssemblyElem(iColumn) = ""
                    
                    'Only add distinct object
                    AssyPath = ""
                    For iNdx = SYSTEM1_COLUMN - 1 To (iColumn - 1)
                        AssyPath = AssyPath & "\" & AssemblyElem(iNdx)
                    Next iNdx
                    
                    Dim systemPath As Variant
                    Dim bDuplicatedSystem As Boolean
                    bDuplicatedSystem = False
                    
                    For Each systemPath In m_CheckDuplicatedSystemsCol
                        If systemPath = AssyPath Then
                            bDuplicatedSystem = True
                        End If
                    Next systemPath

                    Dim oAttributeNames  As Collection
                    Set oAttributeNames = New Collection
                    
                    Dim oAttributeValues  As Collection
                    Set oAttributeValues = New Collection
                    
                    Dim iAttrColumn As Integer
                    For iAttrColumn = iColumnAttrType To 100
                        If Trim(m_oWorksheet.Cells.Item(iRowStart - 1, iAttrColumn)) = "" Then
                            Exit For
                        Else
                            If iRow = iRowStart + 1 Then
                                m_AttributeNameColl.Add Trim(m_oWorksheet.Cells.Item(iRowStart - 1, iAttrColumn))
                            End If
                            oAttributeValues.Add Trim(m_oWorksheet.Cells.Item(iRow, iAttrColumn))
                        End If
                    Next
'
'                    If bDuplicatedSystem Then
'
'                        Dim str_AssyPath As Variant
'                        For Each str_AssyPath In m_DuplicatedObjInfoCol
'                            If str_AssyPath = AssyPath Then GoTo ContinueCheck
'                        Next str_AssyPath
'                        m_DuplicatedObjInfoCol.Add AssyPath
'
'                    End If
  
                    m_AssyHierarchyCol.Add AssemblyElem
                    m_AttributeValColl.Add oAttributeValues
'                    Print #giFilePointer, systemIndexInLogFile & ":" & AssyPath
                    systemIndexInLogFile = systemIndexInLogFile + 1
                    
                End If
            Next iColumn
            
        End If
        
    Next iRow
    
    Set m_oWorksheet = Nothing
    
Exit Sub

ReadASSEMBLYHIERARCHYError:

    Print #giFilePointer, "Error reading system Hierarchy"
    Set m_oWorksheet = Nothing
    
End Sub

Private Function CheckAssemblyExists(AssemblyName As String) As GSCADAssembly.IJAssembly
    
    Dim oChildren As IJDTargetObjectCol
    Dim oChild  As GSCADAssembly.IJAssembly
    Dim onamedAssy As IJNamedItem
    Dim iCount As Integer
    
    If Not m_oParent Is Nothing Then
        ' Get all the children of parent assembly
        Set oChildren = m_oParent.GetChildren
    End If
    
    On Error Resume Next
    
    ' check the child assembly names with input name
    If Not oChildren Is Nothing Then
        For iCount = 1 To oChildren.Count
            Set oChild = oChildren.Item(iCount)
            If Not oChild Is Nothing Then
                If TypeOf oChild Is IJPlanningAssembly Then
                    Set onamedAssy = oChild
                    If onamedAssy.Name = AssemblyName Then
                        Set CheckAssemblyExists = oChild
                        Set oChild = Nothing
                        Exit For
                    End If
                    Set oChild = Nothing
                End If
            End If
        Next iCount
        Set oChildren = Nothing
    Else
        CheckAssemblyExists = Nothing
    End If

End Function

' ***************************************************************************
'
' Function
'   CreateAssembly()
'
' Abstract
'   Creates a new sub-assembly of the currently selected assembly
'
' Description
'   The function instantiates an Assembly Factory, and asks it to create a
'   new Assembly Entity.
'
' Notes
'   This method does not handle errors by itself.  Errors are passed to the
'   calling function.
'
' ***************************************************************************
Private Function CreateAssembly(ByVal oAssemblyName As String) As GSCADAssembly.IJAssembly
    Const METHOD = "CreateAssembly"
    On Error GoTo ErrorHandler
    
   
    Dim oNewObject                  As GSCADAssembly.IJAssembly
    Dim oAssemblyFactory            As GSCADAssembly.IJAssemblyFactory
    Dim oAssmBlockFactory           As GSCADAssemblyBlock.IJAssemblyBlockFactory
    Dim objAssembly                 As GSCADAssembly.IJAssemblyBase
   
    
    ' Create the assembly factory and call the API to create new assembly
    Set oAssemblyFactory = New GSCADAssembly.AssemblyFactory
    Set oNewObject = oAssemblyFactory.CreateAssembly(GetActiveConnection.GetResourceManager(GetActiveConnectionName), m_oParent)
    Set oAssemblyFactory = Nothing
    
    'set Assembly type
    Set objAssembly = oNewObject
    objAssembly.Type = GetFirstAssemType(oNewObject)
    
    'set Assembly name
    Dim oNamedAssembly  As IJNamedItem
    Set oNamedAssembly = objAssembly
    oNamedAssembly.Name = oAssemblyName
    
    'Set Default workcenter and equipments
    SetDefaultWorkcenterAndEquipment oNewObject
    
    Dim oRevMgr As IJRevision
    Dim oToDoList As IJToDoList
    Set oRevMgr = New REVISIONLib.JRevision

    oRevMgr.Compute oToDoList
    
    Set CreateAssembly = oNewObject          ' Return Created Assembly
    
    Set oNewObject = Nothing
    Set objAssembly = Nothing
    Set oAssemblyFactory = Nothing
    Set oAssmBlockFactory = Nothing
    Set oRevMgr = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE, METHOD
End Function

Private Function GetFirstAssemType(oNewObject As Object) As Long
Const METHOD = "GetFirstAssemType"
On Error GoTo ErrorHandler

    Dim objSRDQuery                 As IJSRDQuery
    Dim objWorkcenterQuery          As IJSRDWorkcenterQuery
    Dim strShortDescription1()      As String
    Dim sTableName                  As String
    Dim i                           As Long
    Dim strListOfAssemblyType()     As String
    Dim lListOfAssemblyTypeNumber() As Long
    Dim lAsssemblyType              As Long
    Dim lAsssembly                  As Long
    Dim strAssemblyCodelist         As String
    Dim oCodeListMetaData           As IJDCodeListMetaData
   
    Set oCodeListMetaData = oNewObject

    ' Create the non-persistent SRDServices.SRDQuery to get the Query Object for Workcenter
    Set objSRDQuery = New SRDQuery
    
    ' Fetch the SRDWorkcenterQuery (persisted in the catalog db) to perform SQL queries on the catalog database
    Set objWorkcenterQuery = objSRDQuery.GetWorkcenterQuery()
    
    If Not objWorkcenterQuery Is Nothing Then
        ' add Assemblytype to Assembly Type
        sTableName = "AssemblyType"
        objWorkcenterQuery.GetCodeListArray GetActiveConnection.GetResourceManager(GetActiveConnectionName), sTableName, strListOfAssemblyType, strShortDescription1, lListOfAssemblyTypeNumber
    End If
      
    strAssemblyCodelist = oCodeListMetaData.LongStringValue("AssemblyType", 1)
        
    If strAssemblyCodelist <> vbNullString Then
        GetFirstAssemType = 1
    Else
        GetFirstAssemType = lListOfAssemblyTypeNumber(0)
    End If
    
    Set objSRDQuery = Nothing
    Set objWorkcenterQuery = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE, METHOD
End Function

' ***************************************************************************
'
' Function
'   SetDefaultWorkcenterAndEquipment(oNewAssembly As GSCADAssembly.IJAssembly)
'
' Abstract
'   [in] the function has the newassembly as input
'   [out] Nothing
'   The function find  all workcenters and take the default workcenter which is the first
'   workcenter in the list and put it in the SetWorkCenter function.
'   On basis on the default workcenter the function find the default Equipment where
'   the first is default ExitEquipment (number 0)
'   the second is default JigFloor (number 1, but we don't actually use this)
'   the third is default WeldEquipment (number 2)
'   and the fourth is default BuildEquipment (number 3)
'
' ***************************************************************************
Private Sub SetDefaultWorkcenterAndEquipment(oNewAssembly As GSCADAssembly.IJAssembly)
    Const METHOD = "SetDefaultWorkcenterAndEquipment"
    On Error GoTo ErrorHandler
    
    Dim ListOfWorkcenters()         As String
    Dim ListOfEquipments()          As String
    Dim objWorkcenterQuery          As IJSRDWorkcenterQuery
    Dim objSRDQuery                 As IJSRDQuery
    Dim objPlnAssociateCatalog      As IJPlnAssociateCatalog
    Dim strWorkCenterName           As String
    Dim i                           As Long
    
    Set objPlnAssociateCatalog = oNewAssembly
        
    ' Create the non-persistent SRDServices.SRDQuery to get the Query Object for Workcenter
    Set objSRDQuery = New SRDQuery
    
    ' Fetch the SRDWorkcenterQuery (persisted in the catalog db) to perform SQL queries on the catalog database
    Set objWorkcenterQuery = objSRDQuery.GetWorkcenterQuery()
    
    'get list of all workcenter
    objWorkcenterQuery.GetAllWorkcenters ListOfWorkcenters
    
    'check about the workcenter is bulkload
    If UBound(ListOfWorkcenters) - LBound(ListOfWorkcenters) >= 0 Then
    
        If strWorkCenterName = vbNullString Then
            'add the default workcenter
            strWorkCenterName = ListOfWorkcenters(0)
        Else
            'Checking the possibility that the Workcenter stored in Preferences
            'is no longer present in the catalog. Use the first record in the list
            'in this case.
            For i = LBound(ListOfWorkcenters) To UBound(ListOfWorkcenters)
                If strWorkCenterName = ListOfWorkcenters(i) Then
                    Exit For
                Else
                    If i = UBound(ListOfWorkcenters) Then
                        strWorkCenterName = ListOfWorkcenters(0)
                    End If
                End If
            Next i
        End If
        
        'save the default value of the workcenter
        objPlnAssociateCatalog.SetWorkcenter strWorkCenterName
        
    End If
cleanup:
    Set objSRDQuery = Nothing
    Set objWorkcenterQuery = Nothing
    
Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE, METHOD
    GoTo cleanup
End Sub

'**************************************************************************************
'  Sub             : GetRanges
'  Description     : Read CableWays properties for a given data file and save m_CableWays
'
'**************************************************************************************
Private Sub GetRanges()
    Const MethodStr = "GetRanges"
    On Error Resume Next
    
    'initialize
    iColumnAttrType = 0
    iRowStart = 0
    iRowEnd = 0
    
    ' Find SYSTEMTYPE column
    m_oWorksheet.Cells.Find(What:=TYPE_KEY, After:=ActiveCell, LookIn:=xlValues, _
                      LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                      MatchCase:=False, MatchByte:=False).Activate
    iColumnAttrType = ActiveCell.Column
    ' Find START Row
    m_oWorksheet.Cells.Find(What:=START_KEY, After:=ActiveCell, LookIn:=xlValues, _
                      LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                      MatchCase:=False, MatchByte:=False).Activate
    iRowStart = ActiveCell.Row

    ' Find END  Row
    m_oWorksheet.Cells.Find(What:=END_KEY, After:=ActiveCell, LookIn:=xlValues, _
                      LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                      MatchCase:=False, MatchByte:=False).Activate
    iRowEnd = ActiveCell.Row

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description: Check Excel file has required sheet before reading
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckSheetExists(DataFile As String, strSheetName As String) As Boolean
    Const MethodStr = "CheckSheetExists"
    On Error GoTo CheckSheetExistsError
    
    Dim lSheetCount As Long
    Dim bSheetExists As Boolean
    
    bSheetExists = False
    
    ' Disable the display alerts
    m_oExcelApp.DisplayAlerts = False

    'Open Workbook named DataFile
    Set m_oWorkbook = m_oExcelApp.Workbooks.Open(DataFile)

    'Restore the display alerts flag
    m_oExcelApp.DisplayAlerts = True
    
    On Error Resume Next

    For lSheetCount = 1 To m_oWorkbook.Worksheets.Count
        Set m_oWorksheet = m_oWorkbook.Worksheets(lSheetCount)
            If strSheetName = m_oWorksheet.Name Then
                bSheetExists = True
                Exit For
            End If
    Next lSheetCount
    
    If Not bSheetExists Then
        Print #giFilePointer, " Sheet with name " & strSheetName & "doesn't exist in " & DataFile
    End If
    
    CheckSheetExists = bSheetExists
    
    Set m_oWorksheet = Nothing
    Call m_oWorkbook.Close(False)
    Set m_oWorkbook = Nothing
    
Exit Function

CheckSheetExistsError:
    'Release worksheet
    If Not m_oWorksheet Is Nothing Then Set m_oWorksheet = Nothing
    'Release workbook
    If Not m_oWorkbook Is Nothing Then
        Call m_oWorkbook.Close(False)
        Set m_oWorkbook = Nothing
    End If
    
End Function

'**************************************************************************************
' Description: DeleteAssemblyHierarchy(bstrRootAssy As String)
'**************************************************************************************
Public Sub DeleteAssemblyHierarchy(bstrRootAssy As String, bstrLogFile As String)
    Const MethodStr = "DeleteAssemblyHierarchy"
    On Error GoTo ReadExcelError
    
    Dim strErrMsg       As String
    Dim iIndex          As Integer
    Dim oChild          As Object
    Dim oParentAssembly As IJAssembly
    
    If giFilePointer <> 0 Then
        Close #giFilePointer
        giFilePointer = 0
    End If
    
    giFilePointer = FreeFile
    
    ' open log file for error reporting
    Open bstrLogFile For Append As #giFilePointer
    Print #giFilePointer, ""
    
    If Not bstrRootAssy = "" Then
        
        ' Get the root assembly name
        Dim oRootAssy As Object
        Set oRootAssy = GetObjectFromName(bstrRootAssy)
        
        If Not oRootAssy Is Nothing Then
            If TypeOf oRootAssy Is GSCADAssembly.IJAssembly Then
                
                Dim oPlnIntHelper As IJDPlnIntHelper
                Set oPlnIntHelper = New CPlnIntHelper
                
                ' Get all the child assemblies recursively
                Dim oChildren As IJElements
                Set oChildren = oPlnIntHelper.GetAssemblyChildren(oRootAssy, "IJAssemblyChild", True)
                
                Dim oAssemblyColl   As Collection
                Set oAssemblyColl = New Collection
                
                If Not oChildren Is Nothing Then
                    For iIndex = 1 To oChildren.Count
                        Set oChild = oChildren.Item(iIndex)
                        If Not oChild Is Nothing Then
                            If TypeOf oChild Is GSCADAssembly.IJAssembly Then
                                oAssemblyColl.Add oChild
                            Else
                                ' Disconnect the parts from assembly
                                Dim oAssyChild As IJAssemblyChild
                                Set oAssyChild = oChild
                                
                                Set oParentAssembly = oAssyChild.Parent
                                oParentAssembly.RemoveChild oChild
                                
                                Set oChild = Nothing
                                Set oParentAssembly = Nothing
                                Set oAssyChild = Nothing
                            End If
                        End If
                    Next
                    Set oParentAssembly = Nothing
                    Set oChild = Nothing
                    oChildren.Clear
                    Set oChildren = Nothing
                Else
                    Print #giFilePointer, "No objects under " & bstrRootAssy
                End If
                
                Dim iFlattenAssyCount As Long
                
                ' Flatten the child assemblies
                For iIndex = 1 To oAssemblyColl.Count
                    Dim oAssemblyToFlatten As Object
                    Set oAssemblyToFlatten = oAssemblyColl.Item(iIndex)
                    
                    Set oChildren = oPlnIntHelper.GetAssemblyChildren(oAssemblyToFlatten, "IJAssemblyChild", False)
                    DoFlattenAssembly oAssemblyToFlatten
                    
                    Set oAssemblyToFlatten = Nothing
                Next
                                
                Dim pStructAssocCompute As IJStructAssocCompute
                Set pStructAssocCompute = New StructAssocTools
                
                ' some assemblies move to the parent, so check the assoc falg and delete again
                For iIndex = 1 To oAssemblyColl.Count
                    Dim lAssocFlags As Long
                    pStructAssocCompute.GetAssocFlags oAssemblyColl.Item(iIndex), lAssocFlags
                    
                    ' check assoc flags and faltten the undeleted asseblies
                    ' Reason: when an assembly that has children is flattened, it's children gets moved to the parent assembly
                    If Not (lAssocFlags And RELATION_DELETED) = RELATION_DELETED Then
                        DoFlattenAssembly oAssemblyColl.Item(iIndex)
                    End If
                Next
                                
            End If
        Else
            Print #giFilePointer, "Root assembly name specified for Delete is invalid"
        End If
    Else
        Print #giFilePointer, "Root assembly name for Delete action is not specified"
    End If
    
cleanup:

    'Release all objects
    Close #giFilePointer
    Set oPlnIntHelper = Nothing
    Set oAssemblyColl = Nothing
    Set oChildren = Nothing
    Set oRootAssy = Nothing
    Set pStructAssocCompute = Nothing
    Exit Sub

ReadExcelError:
    Print #giFilePointer, strErrMsg
    Print #giFilePointer, "Delete Assembly Hierarchy  failed"
    GoTo cleanup
    Exit Sub
End Sub

Private Sub DoFlattenAssembly(oAssembly As GSCADAsmHlpers.IJAssembly)
    Const METHOD = "DoFlattenAssembly"
    On Error GoTo ErrorHandler
    
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = oAssembly
    
    If Not oNamedItem Is Nothing Then
        Print #giFilePointer, "Flattening assembly " & oNamedItem.Name
    Else
        Print #giFilePointer, "Assembly is nothing"
        Exit Sub
    End If
    
    If TypeOf oAssembly Is IJPlanningAssembly Then
        
        ' Get Planning specific interface of assembly
        Dim oPlanningAssembly As IJPlanningAssembly
        Set oPlanningAssembly = oAssembly
        
        ' Flatten the assembly
        oPlanningAssembly.Flatten

    ElseIf TypeOf oAssembly Is IJBlock Then
    
        ' Get block specific interface of assembly
        Dim oBlock As GSCADBlock.IJBlock
        Set oBlock = oAssembly
        
        ' Flatten the block
        oBlock.Flatten
    ElseIf TypeOf oAssembly Is IJAssemblyBlock Then
        Dim oAssemblyBlock As IJAssemblyBlock
        Set oAssemblyBlock = oAssembly
        
        'Flatten the AssemblyBlock
        oAssemblyBlock.Flatten
        
    End If
    
cleanup:
    Set oNamedItem = Nothing
    
    Exit Sub
ErrorHandler:
    ' Notify user
    Print #giFilePointer, "DoFlattenAssembly failed"
    GoTo cleanup
End Sub

Private Function GetAttribute(pObject As Object, oAttributeIID As Variant, strAttributeName As String) As IJDAttribute
Const METHOD = "GetAttribute"
On Error GoTo ErrorHandler
    
    Dim oAttributes     As IJDAttributes
    Set oAttributes = pObject
    
    Dim oAttributesCol  As IJDAttributesCol
    Set oAttributesCol = oAttributes.CollectionOfAttributes(oAttributeIID)
    Set oAttributes = Nothing

    Dim i               As Integer
    Dim oAttribute      As IJDAttribute
    
    For i = 1 To oAttributesCol.Count
        Set oAttribute = oAttributesCol.Item(i)
        If oAttribute.AttributeInfo.Name = strAttributeName Then
            Set GetAttribute = oAttribute
            Exit For
        End If
    Next i
     
cleanup:
    Set oAttributesCol = Nothing
    Set oAttribute = Nothing

    Exit Function
    
ErrorHandler:
    ' Notify user
    Print #giFilePointer, "GetAttribute failed"
    GoTo cleanup
End Function

'**************************************************************************************
' Description: AssignPartsToAssemblies(DataFilePath As String)
'**************************************************************************************
Public Sub AssignPartsToAssemblies(bstrLogFile As String)
    Const MethodStr = "AssignPartsToAssemblies"
    On Error GoTo ReadExcelError
        
    Dim strErrMsg       As String
    
    If giFilePointer <> 0 Then
        Close #giFilePointer
        giFilePointer = 0
    End If
    
    giFilePointer = FreeFile
    
    ' open log file for error reporting
    Open bstrLogFile For Append As #giFilePointer
    Print #giFilePointer, ""
    
    ' Assign parts to assemblies
    ExecuteQueryAndAssignPartsToAssembly
    
cleanup:
    'Release all objects
    Close #giFilePointer
    Exit Sub

ReadExcelError:
    Print #giFilePointer, strErrMsg
    Print #giFilePointer, "AssignPartsToAssemblies method failed"
    GoTo cleanup
    Exit Sub
End Sub

' Sample code for assigning parts
Private Sub ExecuteQueryAndAssignPartsToAssembly()
    Const MethodStr = "ExecuteQueryAndAssignPartsToAssembly"
    On Error GoTo ErrorHandler
    
    Dim oCommand            As JMiddleCommand
    Dim oAllObjects         As IEnumMoniker
    Dim oPOM                As IJDPOM

    Set oCommand = New JMiddleCommand
    Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    ' Get all the assemblies from model data base
    With oCommand
        .Prepared = False
        .QueryLanguage = LANGUAGE_SQL
        .ActiveConnection = oPOM.DatabaseID
        .CommandText = "SELECT oid FROM PLANNGASSEMBLY"
        Set oAllObjects = .SelectObjects()
    End With
    
    Dim oObjVariant     As Variant

    ' Get parts that match with assembly and assign
    If Not oAllObjects Is Nothing Then
        For Each oObjVariant In oAllObjects
            
            Dim strAsyyOID  As String
            strAsyyOID = oPOM.DbIdentifierFromMoniker(oObjVariant)
            
            Dim oAssembly   As IJAssembly
            Set oAssembly = oPOM.GetObject(oObjVariant)
            
            Dim oAllObjects2         As IEnumMoniker
            
            With oCommand
                .Prepared = False
                .QueryLanguage = LANGUAGE_SQL
                .ActiveConnection = oPOM.DatabaseID
                .CommandText = "select oid from COREBstrAttribute where iid = " & IID_IJSTXPartInfo & " and value = (select strname from CORENamedItem where oid = HEXTORAW('" & strAsyyOID & "'))"
                Set oAllObjects2 = .SelectObjects()
            End With
            
            Dim oObjVariant2     As Variant
            Dim oObj            As Object
            
            If Not oAllObjects2 Is Nothing Then
                For Each oObjVariant2 In oAllObjects2
                    
                    Set oObj = oPOM.GetObject(oObjVariant2)
                    If Not oObj Is Nothing Then
                        Dim oAssemblyChild   As IJAssemblyChild
                        Set oAssemblyChild = oObj
                        
                        Dim oAssyNamedItem As IJNamedItem
                        Set oAssyNamedItem = oAssembly
                        
                        Dim oPartNamedItem As IJNamedItem
                        Set oPartNamedItem = oAssemblyChild
                        
                        If Not oAssemblyChild.Parent Is oAssembly Then
                            If IsPartAssignableToAssembly(oAssemblyChild, oAssembly) = True Then
                                oAssembly.AddChild oAssemblyChild
                                Print #giFilePointer, "Assigning Part: " & oPartNamedItem.Name & " to Assemly: " & oAssyNamedItem.Name
                            Else
                                Print #giFilePointer, "Assigning Part: " & oPartNamedItem.Name & " to Assemly: " & oAssyNamedItem.Name & " failed as no permission rights"
                            End If
                        Else
                            Print #giFilePointer, "Relation between Part: " & oPartNamedItem.Name & " and Assemly: " & oAssyNamedItem.Name & " already exist"
                        End If
                    End If
                    
                    Set oObj = Nothing
                    Set oObjVariant2 = Nothing
                    Set oAssyNamedItem = Nothing
                    Set oPartNamedItem = Nothing
                    Set oAssemblyChild = Nothing
                Next
            End If
            
            Set oAssembly = Nothing
            Set oAllObjects2 = Nothing
            Set oObjVariant = Nothing
        Next
        
        Set oAllObjects = Nothing
    End If
    
    Set oCommand = Nothing
    Set oObj = Nothing
    Set oAllObjects = Nothing
    Set oObjVariant = Nothing
    Set oPOM = Nothing

Exit Sub

ErrorHandler:
    Print #giFilePointer, "ExecuteQueryAndAssignPartsToAssembly method failed"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function GetCollectionOfAttributesFromObject(pObject As Object, Optional strInputInterface As String) As Collection
    Const MethodStr = "GetCollectionOfAttributesFromObject"
    On Error GoTo ErrorHandler

    Dim pIJAttributeMD As IJDAttributeMetaData
    Dim pIJAttrib As IJDAttributes
    Dim pUserType As IJDUserType
    Dim clsid As Variant
    
    Dim oTempColl As Collection
    Set oTempColl = New Collection
    
    Set pIJAttributeMD = pObject
    Set pIJAttrib = pObject
    Set pUserType = pObject

    clsid = pUserType.RootClassType
    
    ' Get the Interfaces supported by this class
    Dim objInfoColl As IJDInfosCol
    Set objInfoColl = pIJAttributeMD.ClassInterfaces(clsid, PublicFlag_ALL)
    
    Dim i As Integer
    If Not objInfoColl Is Nothing Then
        For i = 1 To objInfoColl.Count
            Dim pInterfaceInfo As IJDInterfaceInfo
            Set pInterfaceInfo = objInfoColl.Item(i)
            
            ' Get the Interface IID
            Dim iid As Variant
            iid = pInterfaceInfo.Type

            ' This is to skip  the "System" Attributes. Attributes of IJRefDataItem and IJMfgItem object
            'if pInterfaceInfo.IsHardCoded = True then
            
            If Not strInputInterface = "" Then
                
                ' Get the Interface UserName
                Dim InterfaceUserName  As String
                InterfaceUserName = pInterfaceInfo.UserName
                
                ' if they're different don't need to fetch those attributes
                If Not strInputInterface = InterfaceUserName Then
                    GoTo NextItem
                End If
                
            End If
            
            Dim oTempAttrColl As IJDAttributesCol
            Set oTempAttrColl = pIJAttrib.CollectionOfAttributes(iid)
            
            Dim j As Integer
            For j = 1 To oTempAttrColl.Count
                If Not oTempAttrColl.Item(j) Is Nothing Then
                    oTempColl.Add oTempAttrColl.Item(j)
                End If
            Next
            
            If Not strInputInterface = "" Then
                Exit For
            End If
NextItem:
            Set oTempAttrColl = Nothing
            Set pInterfaceInfo = Nothing
        Next
    End If
    
    Set GetCollectionOfAttributesFromObject = oTempColl
    
    Set pIJAttributeMD = Nothing
    Set pIJAttrib = Nothing
    Set pUserType = Nothing
    
    Exit Function

ErrorHandler:
    Print #giFilePointer, "GetCollectionOfAttributesFromObject method failed"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetAllPermissionGroupsOfThePlant() As Collection
    Const strMethod = "GetPermissionGroupsOfThePlant"
    On Error GoTo ErrHandler
    
    Dim allPG As ADODB.Recordset
    Dim PGResult As New Collection
    
    Dim strQuery As String
    Dim oCmd As New JMiddleCommand
    Dim Field As ADODB.Field
    Dim vConnName As Variant
    Dim oCatOrModelPOM As IJDPOM
    
    'Get Model POM
    Set oCatOrModelPOM = GetPOM("Model")
    'Get connection name and
    vConnName = oCatOrModelPOM.DatabaseID

    oCmd.QueryLanguage = LANGUAGE_SQL
    oCmd.ActiveConnection = vConnName
    'Get only the PG's with write permission
    strQuery = "SELECT Name, PermissionGroupID FROM PRJMGTPermissionGroup where PermissionGroupID in (select PermissionGroupID from PRJMGTAccessControlRule where AccessRight = 'RUDCA')"
    oCmd.CommandText = strQuery
    Set allPG = oCmd.Execute
    
    While Not allPG.EOF
        PGResult.Add Array(allPG.Fields.Item(0).Value, allPG.Fields.Item(1).Value)
        allPG.MoveNext
    Wend
    
    Set GetAllPermissionGroupsOfThePlant = PGResult
    Set PGResult = Nothing
    Set oCmd = Nothing
    Set PGResult = Nothing
    Set allPG = Nothing
    Set Field = Nothing
    Set oCmd = Nothing
    Set oCatOrModelPOM = Nothing

    Exit Function

ErrHandler:
    Set GetAllPermissionGroupsOfThePlant = PGResult
    Set PGResult = Nothing
    Set oCmd = Nothing
    Set PGResult = Nothing
    Set allPG = Nothing
    Set Field = Nothing
    
    Print #giFilePointer, "GetAllPermissionGroupsOfThePlant method failed"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function CheckInputPerGrpIsValid(ByVal strPermissionGroup As String, ByRef lConditionID) As Boolean
    Const strMethod = "CheckInputPerGrpIsValid"
    On Error GoTo ErrHandler
    
    Dim oPerGrp As Variant

    For Each oPerGrp In m_oPgWriteAccColl
        If oPerGrp(0) = strPermissionGroup Then
            CheckInputPerGrpIsValid = True
            lConditionID = Val(oPerGrp(1))
            Exit Function
        End If
    Next
    
    Exit Function
    
ErrHandler:
    Print #giFilePointer, "CheckInputPerGrpIsValid method failed"
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetActiveConnection() As IJDAccessMiddle
    Const METHOD = "GetActiveConnection"
    On Error GoTo ErrorHandler
    
    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    oCmnAppGenericUtil.GetActiveConnection GetActiveConnection

    Set oCmnAppGenericUtil = Nothing
    Exit Function
    
ErrorHandler:
    Print #giFilePointer, "GetActiveConnection method failed"
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function GetActiveConnectionName() As String

    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    GetActiveConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType("Model")
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing
    
    Exit Function
    
ErrorHandler:
    Print #giFilePointer, "GetActiveConnectionName method failed"
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function GetConfigProjectRootFromDB() As Object
Const METHOD = "GetConfigProjectRootFromDB"
    
    Dim strQuery As String
    Dim oMnkrEnum As IJDEnumMoniker
    Dim vItem As Variant
    Dim oCatOrModelPOM As IJDPOM
    Dim oMiddleCmd As IJADOMiddleCommand
    Dim vConnName As Variant
    
    'query
    strQuery = "select ObjectOid from CORENamedObjects where ObjectName ='PRJMGT_ConfigProjectRoot'"
    
    'Get Model POM
    Set oCatOrModelPOM = GetPOM("Model")
    
    'Prepare command
    Set oMiddleCmd = New JMiddleCommand
    oMiddleCmd.Prepared = False
    oMiddleCmd.QueryLanguage = LANGUAGE_SQL
    
    'Get connection name and
    vConnName = oCatOrModelPOM.DatabaseID
    oMiddleCmd.ActiveConnection = vConnName
    
    'Run the query on the Middle cmd
    oMiddleCmd.CommandText = strQuery
    Set oMnkrEnum = oMiddleCmd.SelectObjects 'There should be only one ConfigProjectRoot record in oMnkrEnum

    If Not oMnkrEnum Is Nothing Then
        'Get ConfigProjectRoot
        For Each vItem In oMnkrEnum
            Set GetConfigProjectRootFromDB = oCatOrModelPOM.GetObject(vItem)
        Next vItem
    End If
    
    Set oMnkrEnum = Nothing
    Set oMiddleCmd = Nothing
    Set oCatOrModelPOM = Nothing
    
Exit Function
ErrorHandler:
    Print #giFilePointer, "GetConfigProjectRootFromDB method failed"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'*************************************************************************************
'Gets the resource manager given the database type.
'*************************************************************************************
Private Function GetPOM(strDBType As String) As IJDPOM
Const METHOD = "GetPOM"
On Error GoTo ErrHandler
    Dim oContext As IJContext
    Dim oAccessMiddle As IJDAccessMiddle
    Dim strConnectMiddle As String
    
    strConnectMiddle = "ConnectMiddle"
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService(strConnectMiddle)
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDBType)
    
    Exit Function
ErrHandler:
    Print #giFilePointer, "GetPOM method failed"
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
