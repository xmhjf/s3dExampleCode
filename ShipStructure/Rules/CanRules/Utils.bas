Attribute VB_Name = "Utils"
Option Explicit

Const MODULE = "CanRuleUtils"


Function Entity_GetElementsOfRelatedEntities(oSource As Object, sNameOfInterface As String, sNameOfRole As String) As IJElements
    Dim pAssocRelations As IJDAssocRelation
    Set pAssocRelations = oSource
    
    Dim pRelationshipCol As IJDRelationshipCol
    Set pRelationshipCol = pAssocRelations.CollectionRelations(sNameOfInterface, sNameOfRole)
    
    Dim pElementsOfRelatedEntities As IJElements
    Set pElementsOfRelatedEntities = New JObjectCollection
    Dim lIndex As Long
    For lIndex = 1 To pRelationshipCol.count
        Dim pRelationship As IJDRelationship
        Set pRelationship = pRelationshipCol.Item(lIndex)
        
        Dim oTarget As Object
        Set oTarget = pRelationship.Target
    
        Call pElementsOfRelatedEntities.Add(pRelationship.Target)
    Next
    Set Entity_GetElementsOfRelatedEntities = pElementsOfRelatedEntities
End Function

Public Function GetCatalogDBConnection() As IJDPOM
    Const METHOD = "GetCatalogDBConnection"
    On Error GoTo ErrHandler
        
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim pConnMiddle As IJDConnectMiddle
    Dim pAccessMiddle As IJDAccessMiddle
    
    Dim jContext As IJContext
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
 
    Set pConnMiddle = jContext.GetService("ConnectMiddle")
 
    Set pAccessMiddle = pConnMiddle
 
    Dim strCatlogDB As String
    strCatlogDB = oDBTypeConfig.get_DataBaseFromDBType("Catalog")
    Set GetCatalogDBConnection = pAccessMiddle.GetResourceManager(strCatlogDB)
  
      
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set pConnMiddle = Nothing
    Set pAccessMiddle = Nothing
Exit Function
ErrHandler:
    MsgBox "Error in GCCans.Can1:GetCatalogDBConnection"
End Function

Public Function GetPOM(oObj As IJDObject) As Object

    Set GetPOM = oObj.ResourceManager

    Exit Function

End Function

Public Sub ToDoListNotify(strCodelistTable As String, nToDoListErrorNum As Long, oObjectInError As Object, oObjectToUpdate As Object)
    Const METHOD = "ToDoListNotify"
    On Error GoTo ErrHandler

    Dim oToDoListHelper As IJToDoListHelper ' Set ToDoListHelper = pointer to the GC Object

    Set oToDoListHelper = oObjectInError
    If Not oToDoListHelper Is Nothing Then
        If oObjectToUpdate Is Nothing Then
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum
        Else
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum, oObjectToUpdate
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    ' Report the error but do not return the error since this call maybe in their
    '   error handler
    HandleError "CommonError", METHOD
    Err.Clear
End Sub
Public Sub HandleError(sModule As String, sMethod As String)
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    Set oEditErrors = Nothing
End Sub

Public Sub SetPG(oInputObj As IJDObject, oOutputObj As IJDObject)

    oOutputObj.PermissionGroup = oInputObj.PermissionGroup

    Exit Sub

End Sub

Public Sub GetUnitVectorTangent(oCurve As IJCurve, ByRef uvecX As Double, ByRef uvecY As Double, ByRef uvecZ As Double)
    Const METHOD = "GetUnitVectorTangent"
    On Error GoTo ErrorHandler

    Dim parStart As Double, parEnd As Double
    Dim PosX As Double, PosY As Double, PosZ As Double
    Dim vecX As Double, vecY As Double, vecZ As Double
    Dim vec2X As Double, vec2Y As Double, vec2Z As Double
    Dim vecLen As Double

    oCurve.ParamRange parStart, parEnd
    oCurve.Evaluate parStart, PosX, PosY, PosZ, vecX, vecY, vecZ, vec2X, vec2Y, vec2Z
    vecLen = vecX * vecX + vecY * vecY + vecZ * vecZ
    vecLen = Sqr(vecLen)
    If vecLen > 0.000001 Then
        uvecX = vecX / vecLen
        uvecY = vecY / vecLen
        uvecZ = vecZ / vecLen
    Else
        uvecX = 1#
        uvecY = 0
        uvecZ = 0
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Function GetAttributeCollection(oBO As Object, attrInterface As String) As Object
    Const METHOD = "GetAttributeCollection"
    On Error GoTo ErrorHandler
    Dim pIJAttrbs As IJDAttributes
    
    If Not oBO Is Nothing Then
        Set pIJAttrbs = oBO
        On Error Resume Next
        Set GetAttributeCollection = pIJAttrbs.CollectionOfAttributes(attrInterface)
        Err.Clear
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function GetAttributeValue(oAttrCollection As CollectionProxy, attrName As String, ByRef outAttrValue As Variant) As Boolean
    Const METHOD = "GetAttributeValue"
    On Error GoTo ErrorHandler
    
    Dim attrValueEmpty As Variant

    Dim oAttr As IJDAttribute

    outAttrValue = attrValueEmpty        'set output to empty.
    GetAttributeValue = False

    If Not oAttrCollection Is Nothing Then
        On Error Resume Next
        Set oAttr = oAttrCollection.Item(attrName)
        If Err.Number = 0 Then
            outAttrValue = oAttr.Value
            GetAttributeValue = True
        Else
            Err.Clear
        End If
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Function SetAttributeValue(oAttrCollection As CollectionProxy, attrName As String, inAttrValue As Variant) As Boolean
    Const METHOD = "SetAttributeValue"
    On Error GoTo ErrorHandler
    
    Dim tol As Double, current As Double
    Dim varCurrent As Variant
    Dim oAttr As IJDAttribute

    SetAttributeValue = False

    If Not oAttrCollection Is Nothing Then

        On Error Resume Next
        Set oAttr = oAttrCollection.Item(attrName)
        
        If Err.Number = 0 Then

            varCurrent = oAttr.Value
                        
            If VarType(varCurrent) = vbDouble Then

                tol = 0.000001
                current = oAttr.Value
                
                If Abs(current - CDbl(inAttrValue)) > tol Then
                    oAttr.Value = inAttrValue
                End If
             
            ElseIf inAttrValue <> varCurrent Then
                oAttr.Value = inAttrValue
            ElseIf IsEmpty(varCurrent) = True Then
                oAttr.Value = inAttrValue
            End If
                 
            SetAttributeValue = True
       
        Else
            Err.Clear
        End If

    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub GetMemberPoint(oCurve As IJCurve, PortId As SPSMemberAxisPortIndex, ByRef PosX As Double, ByRef PosY As Double, ByRef PosZ As Double)
    Const METHOD = "GetMemberPoint"
    On Error GoTo ErrorHandler

    Dim vecX As Double, vecY As Double, vecZ As Double

    If PortId = SPSMemberAxisStart Then
        oCurve.EndPoints PosX, PosY, PosZ, vecX, vecY, vecZ
    Else
        oCurve.EndPoints vecX, vecY, vecZ, PosX, PosY, PosZ
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

