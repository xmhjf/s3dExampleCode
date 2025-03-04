Attribute VB_Name = "Common"
Option Explicit
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : common.bas
'
'Author : RP
'
'Description :
'    SmartPlant Structural common part macro bas file
'
'History:
'
' May 21, 2003   JS     Replaced the Msgbox in HandleError
'                           with JServerErrors AddFromErr method
' Apr 13, 2006   AS     Added functions for Pad -SetInputs4Class_Pad
'
' Mar 14, 2008   SS     DI#134831  Changed the code from CreateObject() to SP3DCreateObject()
'                                  as the symbol is no longer registered.
'
'********************************************************************
Private Const MODULE = "Common"

Private Const MODELDATABASE = "Model"
Private Const CATALOGDATABASE = "Catalog"
Public Const SPSvbError = vbObjectError + 512

Public Enum InputsIndex4Class_Pad
  padMemberPartIndex = 1
  padSurfaceIndex = 2
  padQuestionIndex = 10
End Enum

' Name of inputs of the Pad familly
Private Const PAD_MEMBERPART = "MemberPart"
Private Const PAD_SURFACE = "Surface"
Public Const E_FAIL = -2147467259   'Hard code the efail id

Public Function GetResourceMgr() As IUnknown
Const MT = "GetResourceMgr"
On Error GoTo ErrorHandler

     Dim jContext As IJContext
     Dim oDBTypeConfig As IJDBTypeConfiguration
     Dim oConnectMiddle As IJDAccessMiddle
     Dim strModelDBID As String
     
     'Get the middle context
     Set jContext = GetJContext()
     
     Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
     Set oConnectMiddle = jContext.GetService("ConnectMiddle")
     
     strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
     Set GetResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
 
Exit Function
ErrorHandler: HandleError MODULE, MT
End Function
'
' Routine to get the related parts of the customPlatePart when it is
' a part of macros like Assembly Connection, Equipment Foundation, etc.
'
Public Function GetRelatedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements) As String
Const MT = "GetRelatedParts"
On Error GoTo ErrorHandler
GetRelatedParts = "ERROR"

    Dim pDesignParent As IJDesignParent
    Dim pIJElements As IJElements
    Dim pIJDesignChild As IJDesignChild
    
    Dim i As Integer
    Dim result As Boolean
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim object As Object
    Dim pPort As IJPort
    Dim pObjectCollection As IJDObjectCollection
    Set pObjectCollection = New JObjectCollection

    Set pIJDesignChild = pPartObject
    If Not pIJDesignChild Is Nothing Then
    Else
        Exit Function
    End If
    
    Set pDesignParent = pIJDesignChild.GetParent
    If Not pDesignParent Is Nothing Then
    Else
        Exit Function
    End If

    pDesignParent.GetChildren pObjectCollection, vbNullString
    
    Dim pSubParent As IJDesignParent
    Dim pSubCollection As IJDObjectCollection
    Dim subObject As Object
    
    ' check to see if the design child is a design parent by itself;
    ' if so, add its children to the IFC list
    On Error Resume Next
    For Each object In pObjectCollection
        
        If pPartObject Is object Then
        Else
            On Error Resume Next
            Set pSubParent = object
            If Not pSubParent Is Nothing Then
                Set pSubCollection = New JObjectCollection
                pSubParent.GetChildren pSubCollection, vbNullString
                ' tr 66565 - should check for list count before adding
                If pSubCollection.Count > 0 Then
                    For Each subObject In pSubCollection
                        pIJMonUnks.Add subObject
                    Next
                Else
                    pIJMonUnks.Add object
                End If
                pSubCollection.Clear
                Set pSubCollection = Nothing
                Set pSubParent = Nothing
            Else
                pIJMonUnks.Add object
            End If
        End If
    Next

    Dim oSmartOcc As IJSmartOccurrence
    Dim pUnk As Object
    Dim oSupportingSurface As Object
    Dim oRefProxy As IJDReferenceProxy
    Dim oStructPort As IJPort
 
    ' check the context the plate part is in to return the connectables
    If TypeOf pDesignParent Is ISPSEquipFoundation Then
       
        Dim FndDefServices As ISPSEquipFndnDefServices
        Dim oItem As IJSmartItem
       
        Set oSmartOcc = pDesignParent
        Set oItem = oSmartOcc.ItemObject
        Set FndDefServices = SP3DCreateObject(oItem.definition)
    
        Dim Supported As IJElements
        Set Supported = New JObjectCollection
        Dim Supporting As IJElements    ' tr 66565 - type error should be IJElements
        Set Supporting = New JObjectCollection
    
        FndDefServices.GetInputs pDesignParent, Supported, Supporting
        
        Dim oFndPort As IJPort
        Dim oEquipment As IJEquipment
        
        For i = 1 To Supported.Count
            Set oFndPort = Supported.Item(i)
            If Not oFndPort Is Nothing Then
                Set oEquipment = oFndPort.Connectable
                If Not oEquipment Is Nothing Then
                    pIJMonUnks.Add oEquipment
                End If
            End If
            Set oFndPort = Nothing
            Set oEquipment = Nothing
        Next i
        
        ' ' tr 66565 - need to add supporting surface to ifc connectable list
        If Supporting.Count > 0 Then
            On Error Resume Next
            Set oSupportingSurface = Supporting.Item(1)
            If Not oSupportingSurface Is Nothing Then
                On Error Resume Next
                Set oRefProxy = oSupportingSurface
    
                If Not oRefProxy Is Nothing Then
                    Set pUnk = oRefProxy.Reference
                    pIJMonUnks.Add pUnk
                    Set pUnk = Nothing
                Else ' for slab surfaces
                    On Error Resume Next
                    Set oStructPort = oSupportingSurface
                    If Not oStructPort Is Nothing Then
                        Set pUnk = oStructPort.Connectable
                    End If
                    If Not pUnk Is Nothing Then
                        pIJMonUnks.Add pUnk
                    End If
                    Set pUnk = Nothing
                End If
                Set oSupportingSurface = Nothing
                Set oRefProxy = Nothing
            End If
        End If
    
    ElseIf TypeOf pDesignParent Is IJStructAssemblyConnection Then
    
        Set oStructConn = pDesignParent
        If Not oStructConn Is Nothing Then
        Else
            Exit Function
        End If
    
        oStructConn.enumPorts colPorts
        If colPorts.Count = 0 Then
            Exit Function
        End If
        For i = 1 To colPorts.Count
            Set pPort = colPorts.Item(i)
            pIJMonUnks.Add pPort.Connectable
        Next
        
        ' tr 66565 - assembly conn should add ref coll list to ifc connected parts in addition
        ' to appconnection connectable
        
        On Error Resume Next
        Set oSmartOcc = pDesignParent
        
        If Not oSmartOcc Is Nothing Then
            Dim oRefColl As IJDReferencesCollection
            Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
            
            If oRefColl.IJDEditJDArgument.GetCount >= 1 Then ' support plane
                Set oSupportingSurface = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
                If Not oSupportingSurface Is Nothing Then
                    On Error Resume Next
                    Set oRefProxy = oSupportingSurface
        
                    If Not oRefProxy Is Nothing Then
                        Set pUnk = oRefProxy.Reference
                        pIJMonUnks.Add pUnk
                        Set pUnk = Nothing
                    Else ' for slab surfaces
                        On Error Resume Next
                        Set oStructPort = oSupportingSurface
                        If Not oStructPort Is Nothing Then
                            Set pUnk = oStructPort.Connectable
                        End If
                        If Not pUnk Is Nothing Then
                            pIJMonUnks.Add pUnk
                        End If
                        Set pUnk = Nothing
                    End If
                    Set oSupportingSurface = Nothing
                    Set oRefProxy = Nothing
                End If
            End If
            
        End If
        
    End If
      
    GetRelatedParts = ""
    
Exit Function
ErrorHandler: HandleError MODULE, MT
End Function

Public Sub ConnectSmartOccurrence(pSO As IJSmartOccurrence, pRefColl As IJDReferencesCollection)
Const MT = "ConnectSmartOccurrence"
 On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pSO
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    pCollectionHelper.Add pRefColl, "RC", pRelationshipHelper
    Set pRevision = New JRevision
    pRevision.AddRelationship pRelationshipHelper
  
 Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub


Public Sub SetCustomPlatePartGenerationAE(pCPP As IJStructCustomPlatePart)
Const METHOD = "SetCustomPlatePartGenerationAE"
On Error GoTo ErrorHandler
      Dim oGenerationAE As StructCustomPlatePartGenerationAE
      Dim oIJStructGeometryPattern As IJStructGeometryPattern
      Dim CollOfParents As IJElements
      Set CollOfParents = New JObjectCollection

      CollOfParents.Add pCPP

      Set oIJStructGeometryPattern = pCPP

      oIJStructGeometryPattern.SetGeometryPattern "StructCustomPlatePart.StructCustomPlatePartGenerationAE.1", CollOfParents, oGenerationAE
Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
End Sub

Sub SetInputs4Class_Pad(pDef As IJDSymbolDefinition)
Const METHOD = "SetInputs4Class_Pad"
  'MsgBox METHOD
On Error GoTo ErrorHandler
  Dim pInput As IJDInput
  Set pInput = New DInput

  pInput.Name = PAD_MEMBERPART
  pInput.Description = "Supported Member Part"
  pInput.index = padMemberPartIndex
  pDef.IJDInputs.Add pInput
  pInput.Reset
    
  pInput.Name = PAD_SURFACE
  pInput.Description = "Supporting Surface"
  pInput.index = padSurfaceIndex
  pDef.IJDInputs.Add pInput
  pInput.Reset
Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear

End Sub

Sub GetArguments4Class_Pad(pRep As IJDRepresentationDuringGame, pMemberPart As Object, pSurface As Object)
  
  Dim pSelector As IJDSymbolDefinition
  Set pSelector = GetSelector(pRep)
  
  Set pMemberPart = pSelector.IJDInputs.GetInputAtIndex(padMemberPartIndex).IJDInputDuringGame.result
  Set pSurface = pSelector.IJDInputs.GetInputAtIndex(padSurfaceIndex).IJDInputDuringGame.result

End Sub

Public Function GetSelector(pRepSCM As IJDRepresentationStdCustomMethod) As IJDSymbolDefinition
  Dim pRepDG As IJDRepresentationDuringGame
  Set pRepDG = pRepSCM
  Set GetSelector = pRepDG.definition
End Function
