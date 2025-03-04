Attribute VB_Name = "CoreHelpers"
Option Explicit
' depends on IMSRelation (IJDRelationshipCol)
' depends on IMSSymbolEntities (IJToDoListHelper)
Private Const E_FAIL = &H80004005

Private m_pRevision As IJRevision

Public Function GetRevision() As IJRevision
    If m_pRevision Is Nothing Then
        Set m_pRevision = New JRevision
    End If
    Set GetRevision = m_pRevision
End Function

Public Function Entity_GetRelationshipColOfRelatedEntities(oSource As Object, sNameOfRelation As String, sNameOfRole As String) As IJDRelationshipCol
    Dim pAssocRelations As IJDAssocRelation
    Set pAssocRelations = oSource
    
    Dim pRelationshipCol As IJDRelationshipCol
    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfRelation, sNameOfRole)
    
    Set Entity_GetRelationshipColOfRelatedEntities = pRelationshipCol
End Function
Public Function Entity_GetElementsOfRelatedEntities(oSource As Object, sNameOfRelation As String, sNameOfRole As String) As IJElements
    Dim pAssocRelations As IJDAssocRelation
    Set pAssocRelations = oSource
    
    Dim pRelationshipCol As IJDRelationshipCol
    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfRelation, sNameOfRole)
    
    Dim pElementsOfRelatedEntities As IJElements
    Set pElementsOfRelatedEntities = New JObjectCollection
    Dim lIndex As Long
    For lIndex = 1 To pRelationshipCol.Count
        Dim pRelationship As IJDRelationship
        Set pRelationship = pRelationshipCol.Item(lIndex)
        
        Dim oTarget As Object
        Set oTarget = pRelationship.Target
    
        Call pElementsOfRelatedEntities.Add(pRelationship.Target)
    Next
    Set Entity_GetElementsOfRelatedEntities = pElementsOfRelatedEntities
End Function
Public Sub Entity_DisconnectRelatedEntity(oSource As Object, sNameOfInterface As String, sNameOfRole As String)
    Dim pAssocRelations As IJDAssocRelation
    Set pAssocRelations = oSource
    
    Dim pRelationshipCol As IJDRelationshipCol
    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfInterface, sNameOfRole)
    
    Dim pRevision As IJRevision: Set pRevision = New JRevision
    Dim lIndex As Long
    For lIndex = 1 To pRelationshipCol.Count
        Dim pRelationship As IJDRelationship
        On Error Resume Next
        Set pRelationship = pRelationshipCol.Item(lIndex)
        On Error GoTo 0
        If Not pRelationship Is Nothing Then
            Call pRevision.RemoveRelationship(pRelationship)
        End If
    Next
End Sub
Public Sub ProcessError(pGeometricConstruction As IJGeometricConstruction, lNumber As Long)
    Dim pToDoListHelper As IJToDoListHelper
    Set pToDoListHelper = pGeometricConstruction
    
    Call pToDoListHelper.SetErrorInfo(pGeometricConstruction.TypeName + "_Errors", lNumber)
    Err.Raise E_FAIL
End Sub
Public Function Object_GetPOM(pObject As IJDObject) As IJDPOM
    On Error Resume Next
    Set Object_GetPOM = pObject.ResourceManager
    On Error GoTo 0
End Function
Public Function Object_Create(oObject As Object, pPOM As IJDPOM) As Object
    If Not pPOM Is Nothing Then Call GetRevision().Add(pPOM, Nothing, oObject)
    Set Object_Create = oObject
End Function
Public Function ResourceManager(oObject As Object) As IUnknown
    Dim pObject As IJDObject
    Set pObject = oObject
    Set ResourceManager = Nothing
    On Error Resume Next
    Set ResourceManager = pObject.ResourceManager
    On Error GoTo 0
End Function
Public Function POM_MakeObjectPersistent(pPOM As IJDPOM, oObject As Object) As Object
    Call GetRevision().Add(pPOM, Nothing, oObject)
    
    Set POM_MakeObjectPersistent = oObject
End Function
Public Sub Object_Delete(pObject As IJDObject)
    Call GetRevision().Delete(pObject)
End Sub
Public Function Collection_ContainsName(pCollection As Collection, sName As String) As Boolean
    'initialize result
    Let Collection_ContainsName = False
    
    ' loop on the collection
    Dim i As Integer
    For i = 1 To pCollection.Count
        ' retrieve name of the current item
        Dim sNameOfItem As String: Let sNameOfItem = pCollection.Item(i)
        
        ' compare item name to searched name
        If sName = sNameOfItem Then
            Let Collection_ContainsName = True
            Exit For
        End If
    Next
End Function
Public Sub Object_ShowName(oObject As Object, sMessage As String)
    If oObject Is Nothing Then
        Debug.Print sMessage + "= Nothing"
    Else
        Dim pNamedItem As IJNamedItem: Set pNamedItem = oObject
        Debug.Print sMessage + "= " + pNamedItem.Name
    End If
End Sub
Public Function Object_GetName(oObject As Object) As String
    If oObject Is Nothing Then
        Let Object_GetName = "Nothing"
    Else
        Dim pNamedItem As IJNamedItem: Set pNamedItem = oObject
        Let Object_GetName = pNamedItem.Name
    End If
End Function
Public Function Elements_GetKeys(pElements As IJElements) As String()
    ' prepare result
    Dim sKeys() As String: ReDim sKeys(1 To pElements.Count)
    
    ' loop on elements
    Dim i As Integer
    For i = 1 To pElements.Count
        Let sKeys(i) = pElements.GetKey(pElements(i))
    Next
    
    ' return result
    Let Elements_GetKeys = sKeys
End Function
Public Function Elements_GetDummyKeys(pElements As IJElements) As String()
    ' prepare result
    Dim sKeys() As String: ReDim sKeys(1 To pElements.Count)
    
    ' loop on elements
    Dim i As Integer
    For i = 1 To pElements.Count
        Let sKeys(i) = "#"
    Next
    
    ' return result
    Let Elements_GetDummyKeys = sKeys
End Function
Public Function Elements_GetValidElements(pElements As IJElements) As IJElements
    ' prepare result
    Dim pElementsOfValidElements As IJElements: Set pElementsOfValidElements = pElements.Clone
    Dim i As Integer
    For i = pElementsOfValidElements.Count To 1 Step -1
        If IsDeleted(pElementsOfValidElements(i)) Then pElementsOfValidElements.Remove (i)
    Next
    
    ' return result
    Set Elements_GetValidElements = pElementsOfValidElements
End Function
Public Function IsDeleted(oObject As Object) As Boolean
    Let IsDeleted = False
    Const RELATION_DELETED = &H1000000
    Dim pStructAssocCompute As IJStructAssocCompute: Set pStructAssocCompute = New StructAssocTools
    Dim lAssocFlags As Long
    Call pStructAssocCompute.GetAssocFlags(oObject, lAssocFlags)
    If (lAssocFlags And RELATION_DELETED) = RELATION_DELETED Then Let IsDeleted = True
End Function
Public Sub Object_ShowDatabaseID(oObject As Object, sMessage As String)
    Dim pObject As IJDObject: Set pObject = oObject
    
    ' get POM
    On Error Resume Next
    Dim pPOM As IJDPOM: Set pPOM = pObject.ResourceManager
    
    ' get moniker
    Dim pMoniker As IMoniker: Set pMoniker = pPOM.GetObjectMoniker(oObject)
    
    ' get database id
    Dim sDatabaseID As String: Let sDatabaseID = pPOM.DbIdentifierFromMoniker(pMoniker)
    Debug.Print "DatabaseID of " + sMessage + "= " + sDatabaseID
End Sub
Function ArrayOfStrings_ContainsString(sStrings() As String, sString As String) As Boolean
    ' prepare result
    Let ArrayOfStrings_ContainsString = False
    
    ' if the array is empty, return false
    Dim iSize As Integer: Let iSize = 0: On Error Resume Next: Let iSize = UBound(sStrings, 1): On Error GoTo 0
    If iSize = 0 Then Exit Function
    
    ' loop on items
    Dim i As Integer
    For i = 1 To iSize
        If sStrings(i) = sString Then
            Let ArrayOfStrings_ContainsString = True
            Exit For
        End If
    Next i
End Function
Sub Keys_RemoveKey(sKeys() As String, sKey As String)
    Dim iCount As Integer: Let iCount = UBound(sKeys, 1)
    Dim i As Integer
    For i = 1 To iCount
        If sKey = sKeys(i) Then
            sKeys(i) = ""
            Exit For
        End If
    Next
End Sub
Sub Elements_ReplaceElements(pElementsToBeReplaced As IJElements, pElementsReplacing As IJElements, Optional sKey As String = "")
    If pElementsToBeReplaced.Count > 1 Then
        MsgBox "Elements_ReplaceElements: should apply only to a single collection"
    Else
        Dim bIsElementSame As Boolean: Let bIsElementSame = False
        If pElementsToBeReplaced.Count = 1 Then
            If pElementsToBeReplaced(1) Is pElementsReplacing(1) Then Let bIsElementSame = True
        End If
        
        If Not bIsElementSame Then
            Call pElementsToBeReplaced.Clear
            If sKey = "" Then
                Call pElementsToBeReplaced.Add(pElementsReplacing(1))
            Else
                Call pElementsToBeReplaced.Add(pElementsReplacing(1), sKey)
            End If
        End If
    End If
End Sub
Sub Elements_ReplaceElementsWithKey(pElementsToBeReplaced As IJElements, pElementsReplacing As IJElements, sKey As String)
    If sKey = "#" Then Exit Sub
    
    Dim oElement As Object: On Error Resume Next: Set oElement = pElementsToBeReplaced(sKey): On Error GoTo 0
    If Not oElement Is pElementsReplacing(1) Then
        If Not oElement Is Nothing Then Call pElementsToBeReplaced.Remove(sKey)
        Call pElementsToBeReplaced.Add(pElementsReplacing(1), sKey)
    End If
End Sub
Sub Elements_ReplaceElementWithSameKey(pElements As IJElements, oElementOriginal As Object, oElementReplacing As Object)
    Call DebugIn("Elements_ReplaceElementWithSameKey")
    Call DebugInput("ElementOriginal", oElementOriginal)
    Call DebugInput("ElementReplacing", oElementReplacing)
    
    ' collect the keys and the elements
     Dim lCount As Long: Let lCount = pElements.Count
    Dim sKeys() As String: ReDim sKeys(1 To pElements.Count)
    Dim oElements() As Object: ReDim oElements(1 To pElements.Count)
    Dim j As Integer
    For j = 1 To lCount
       Set oElements(j) = pElements(j)
       sKeys(j) = pElements.GetKey(oElements(j))
    Next

   
    ' clear the collection
    Call pElements.Clear
    
    ' reload the collection
    For j = 1 To lCount
        If oElements(j) Is oElementOriginal Then
            If Not pElements.Contains(oElementReplacing) Then
                Call DebugMsg("Replace with key= " + sKeys(j))
                Call pElements.Add(oElementReplacing, sKeys(j))
            End If
        Else
            If Not pElements.Contains(oElements(j)) Then
                Call pElements.Add(oElements(j), sKeys(j))
            End If
        End If
    Next
    
    Call DebugOut
End Sub
Sub Elements_RenameKeyOfElement(pElements As IJElements, vKeyOld As Variant, vKeyNew As Variant)
    Call DebugIn("Elements_RenameKeyOfElement")
    Call DebugInput("KeyOld", vKeyOld)
    Call DebugInput("KeyNew", vKeyNew)
    
    Dim oElement As Object: Set oElement = pElements(vKeyOld)
    Call pElements.Remove(vKeyOld)
    If vKeyNew <> "" Then
        Call pElements.Add(oElement, vKeyNew)
    Else
        Call pElements.Add(oElement)
    End If
    
    Call DebugOut
End Sub
Sub Elements_RemoveUnneeded(pElements As IJElements, sKeysOfUnneededElements() As String)
    Dim pElementsOfUnneededElements As IJElements: Set pElementsOfUnneededElements = Elements_GetUnneeded(pElements, sKeysOfUnneededElements)
    Dim i As Integer
    For i = 1 To pElementsOfUnneededElements.Count
        Dim oUnneededElement As Object: Set oUnneededElement = pElementsOfUnneededElements(i)
        Dim sKey As String: Let sKey = pElementsOfUnneededElements.GetKey(oUnneededElement)
        If Mid(sKey, 1, 2) = "TR" Then
            ' we are removing a trimimng plane
            'Call pElementsOfUnneededElements.Remove(sKey)
        End If
        Call pElements.Remove(sKey)
    Next i
    Call Elements_DeleteUnneeded(pElements, pElementsOfUnneededElements)
End Sub
Private Sub Elements_DeleteUnneeded(pElements As IJElements, pElementsOfUnneededElements As IJElements)
    ' remove the un-needed elements
    'Call pElements.RemoveElements(pElementsOfUnneededElements)
    
    ' delete them
    Dim pRevision As IJRevision: Set pRevision = New JRevision
    Dim i As Integer
    For i = 1 To pElementsOfUnneededElements.Count
        Call pRevision.Delete(pElementsOfUnneededElements(i))
    Next
End Sub
Private Function Elements_GetUnneeded(pElements As IJElements, sKeysOfUnneededElements() As String) As IJElements
    ' prepare the result
    Dim pElementsOfUnneededElements As IJElements: Set pElementsOfUnneededElements = New JObjectCollection
    
    If True Then
        Dim iCount As Integer: Let iCount = UBound(sKeysOfUnneededElements, 1)
        Dim i As Integer
        For i = 1 To iCount
            If sKeysOfUnneededElements(i) <> "" Then
                Dim oElement As Object: Set oElement = pElements(sKeysOfUnneededElements(i))
                Call pElementsOfUnneededElements.Add(oElement, sKeysOfUnneededElements(i))
            End If
        Next
    End If
    
    ' return the result
    Set Elements_GetUnneeded = pElementsOfUnneededElements
End Function

