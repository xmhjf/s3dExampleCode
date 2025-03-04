VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssemblyParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: AssemblyParts.cls
'
'
' Abstract:
'
'  This rule assigns parts in "Unassigned" folder to assemblies automatically.
'  Parts include:
'    Plate parts(system derived and stand alone)
'    Profiles part(system derived and stand alone)
'    ER parts(ER profile part, ER plate part)
'    Brackets
'    Collars
'
'  Parts under "Unassigned" are sorted by their areas.
'  Start from the biggest one, create an assembly for it,and assign following parts to it:
'    . Butt connected parts
'        From same root system,or manual logical connection
'    . Brackets which do not have any children and meet either A or B
'        A. The bracket uses current plate as support
'        B. The bracket uses a profile as support,the profile is on current plate,
'           and current plate is a TBh
'    .Profile parts
'    .ER parts
'
'    If no connected brackets,profiles,or ERs,and current plate is not Hull or deck
'    . Bounded part which do not have any children
'
Option Explicit

Implements IAutomatAssembly

Private Const m_sClassName = "AssemblyParts"
Private Const m_sRuleName = m_sProjectName + m_sClassName + ".cls"
Private Const m_sModule As String = m_sProjectPath + m_sClassName
Private Const m_bDebug As Boolean = False

' Parts should be in the block
'
Private Function IAutomatAssembly_DoAssembliesByRule(ByVal oElements As GSCADPlnIntHelper.IJElements, ByVal eOptions As GSCADPlnIntHelper.EComboOptions) As Boolean
          
   Const sMETHOD As String = "IAutomtassembly_DoAssembliesByRule"
   On Error GoTo ErrorHandler
   
   IAutomatAssembly_DoAssembliesByRule = True
   
   Dim stUnAssignedPartInfo As UNASSIGNED_PART_INFO
   Dim nBlockIndex As Long
   Dim oBlock As IJBlock
   
   stUnAssignedPartInfo.nPlatePartCount = 0
   stUnAssignedPartInfo.nCollarCount = 0
   stUnAssignedPartInfo.nProfilePartCount = 0
   stUnAssignedPartInfo.nERPartCount = 0
      
   For nBlockIndex = 1 To oElements.Count
      Set oBlock = oElements.Item(nBlockIndex)
      GetUnassignedPartInfo oBlock, _
                            stUnAssignedPartInfo, _
                            eOptions
      SortPlatePartByArea stUnAssignedPartInfo
      
      If m_bDebug = True Then
         LogPartInfo stUnAssignedPartInfo
      End If
      
      Dim nPlatePartIndex As Long
      Dim oConnectedPart As Object
      Dim oButtConnectedParts As Collection
      Dim oTeeConnectedParts As Collection
      Dim oStiffeningParts As Collection
      Dim oERParts As Collection
      Dim oBracketParts As Collection
      Dim oAssemblyParts As Collection
      Dim oPlate As IJPlate
      
      For nPlatePartIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
         If stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).bAssigned = False Then
            Set oAssemblyParts = New Collection
            GetConnectedParts _
                         CONNECTED_PART_MASK_ALL, _
                         nPlatePartIndex, _
                         Nothing, _
                         True, _
                         stUnAssignedPartInfo, _
                         oButtConnectedParts, _
                         oTeeConnectedParts, _
                         oStiffeningParts, _
                         oERParts, _
                         oBracketParts
                         
            ' Following parts will be assigned to this assembly:
            ' .Butt connected parts
            ' .If the plate has children then
            '    Bracket parts
            '    Stiffening parts
            '      If plate is TBH,brackets which are connected to the stiffners and do not have children
            '    ER Parts
            '  else
            '    Tee connected plate parts which do not have children
            '
            
            ' Butt parts
            If Not oButtConnectedParts Is Nothing Then
               ProcessButtConnectedParts _
                                   oButtConnectedParts, _
                                   stUnAssignedPartInfo, _
                                   oAssemblyParts
               Set oButtConnectedParts = Nothing
            End If
            
            
            If Not oBracketParts Is Nothing Then
               ProcessDirectlyConnectedBracketParts _
                                   oBracketParts, _
                                   stUnAssignedPartInfo, _
                                   oAssemblyParts
            End If
                        
            Set oPlate = stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).oPart
            If Not oStiffeningParts Is Nothing Then
               ProcessStiffeningParts _
                                oPlate.plateType, _
                                oStiffeningParts, _
                                stUnAssignedPartInfo, _
                                oAssemblyParts
            End If
            
            If Not oERParts Is Nothing Then
               ProcessConnectedERParts _
                                oERParts, _
                                stUnAssignedPartInfo, _
                                oAssemblyParts
            End If
               
            If oBracketParts Is Nothing And _
               oStiffeningParts Is Nothing And _
               oERParts Is Nothing Then
               
               If Not (oPlate.plateType = Hull Or oPlate.plateType = DeckPlate) Then
                  ' This plate has no children,and it's not hull,not deck
                  ' get Tee connected plates which do not have children neither
                  '
                  If Not oTeeConnectedParts Is Nothing Then
                     ProcessTeeConnectedPartsWithoutChildren _
                                    oTeeConnectedParts, _
                                    stUnAssignedPartInfo, _
                                    oAssemblyParts
                     Set oTeeConnectedParts = Nothing
                  End If
               End If
            Else
               Set oBracketParts = Nothing
               Set oStiffeningParts = Nothing
               Set oERParts = Nothing
            End If
            Set oPlate = Nothing
            
            If m_bDebug = True Then
               LogAssemblyMembers oAssemblyParts
            End If
            
            Dim oNewAssembly As IJAssembly
            Dim nAssemblyPartIndex As Long
            
            Set oNewAssembly = CreateAssembly(oBlock)
            If Not oNewAssembly Is Nothing Then
                SetAssemblyNamingRule oNewAssembly, "CAssembly"
                SetAssemblyPropertiesByRule oNewAssembly
               For nAssemblyPartIndex = 1 To oAssemblyParts.Count
                  oNewAssembly.AddChild oAssemblyParts.Item(nAssemblyPartIndex)
               Next
            End If
            Set oAssemblyParts = Nothing
         End If ' Not assigned
      Next ' Each plate
      
      ' After all plate parts are processed, assign collars
      If stUnAssignedPartInfo.nCollarCount > 0 Then
         AssignCollarParts stUnAssignedPartInfo
      End If
   Next
   
   Exit Function
   
ErrorHandler:
   MsgBox Err.Description
End Function
'
' If penetrated plate is not in assembly collar info
'    Create penetrated plate info in assembly collar info
' Return penetrated plate info index in assembly collar info
'
Private Function GetPenetratedPlateIndexFromAssemblyCollarInfo( _
      oCollarPenetratedPart As Object, _
      ByRef stAssemblyCollarInfo As ASSEMBLY_COLLAR_INFO) As Long
      
   Const sMETHOD As String = "GetPenetratedPlateIndexFromAssemblyCollarInfo"
   On Error GoTo ErrorHandler
   
   GetPenetratedPlateIndexFromAssemblyCollarInfo = -1
   
   Dim bEntryExisted As Boolean
   Dim nPenetratedIndex As Long
   Dim stPenetratedPlateCollarInfo As PENETRATED_PLATE_COLLAR_INFO
   
   bEntryExisted = False
   For nPenetratedIndex = 0 To stAssemblyCollarInfo.nPenetratePlateCount - 1
      stPenetratedPlateCollarInfo = stAssemblyCollarInfo.arrayPenetratedPlateCollarInfo(nPenetratedIndex)

      If oCollarPenetratedPart Is stPenetratedPlateCollarInfo.oPenetratedPart Then
         bEntryExisted = True
         Exit For
      End If
   Next
      
   If bEntryExisted = False Then
      Dim nPenetratedCount As Long
      Dim oSDHelper As New StructDetailHelper
      Dim oPenetratedRootSystem As IJSystem
      
      nPenetratedIndex = stAssemblyCollarInfo.nPenetratePlateCount
      nPenetratedCount = nPenetratedIndex + 1
      stAssemblyCollarInfo.nPenetratePlateCount = nPenetratedCount
      
      ReDim Preserve stAssemblyCollarInfo.arrayPenetratedPlateCollarInfo(nPenetratedCount)
      
      Set stAssemblyCollarInfo.arrayPenetratedPlateCollarInfo(nPenetratedIndex).oPenetratedPart = oCollarPenetratedPart
      oSDHelper.IsPartDerivedFromSystem oCollarPenetratedPart, oPenetratedRootSystem, True
      Set stAssemblyCollarInfo.arrayPenetratedPlateCollarInfo(nPenetratedIndex).oPenetratedRoot = oPenetratedRootSystem
   End If
   
   GetPenetratedPlateIndexFromAssemblyCollarInfo = nPenetratedIndex
   
   Exit Function
   
ErrorHandler:
   MsgBox Err.Description
End Function
'
'  If base plate info is not in penetrated plate collar info
'     Create base plate info
'  Return base plate info index in penetrated plate collar info
Private Function GetBasePlateIndexFromPenetratedPlateCollarInfo( _
              ByVal oBasePlate As Object, _
              ByRef stPenetratedPlateCollarInfo As PENETRATED_PLATE_COLLAR_INFO) As Long
   Const sMETHOD As String = "GetBasePlateIndexFromPenetratedPlateCollarInfo"
   On Error GoTo ErrorHandler
              
   GetBasePlateIndexFromPenetratedPlateCollarInfo = -1
   
   Dim bEntryExisted As Boolean
   Dim nBasePlateIndex As Long
   
   bEntryExisted = False
   For nBasePlateIndex = 0 To stPenetratedPlateCollarInfo.nBasePlateCount - 1
      If oBasePlate Is stPenetratedPlateCollarInfo.arrayBasePlateCollarInfo(nBasePlateIndex).oBasePlate Then
         bEntryExisted = True
         Exit For
      End If
   Next
   
   If bEntryExisted = False Then
      Dim nBasePlateCount As Long
      
      nBasePlateIndex = stPenetratedPlateCollarInfo.nBasePlateCount
      ' Increase base plate count
      nBasePlateCount = nBasePlateIndex + 1
      stPenetratedPlateCollarInfo.nBasePlateCount = nBasePlateCount

      ReDim Preserve stPenetratedPlateCollarInfo.arrayBasePlateCollarInfo(nBasePlateCount)
      Set stPenetratedPlateCollarInfo.arrayBasePlateCollarInfo(nBasePlateIndex).oBasePlate = oBasePlate
   End If
   
   GetBasePlateIndexFromPenetratedPlateCollarInfo = nBasePlateIndex
   
   Exit Function
   
ErrorHandler:
   MsgBox Err.Description
End Function
'
' Store collar to base plate collar info (collars are grouped by base plates)
'
Private Sub StoreCollarInBasePlateInfo( _
        ByVal oCollar, _
        ByRef stBasePlateCollarInfo As BASE_PLATE_COLLAR_INFO)
   Const sMETHOD As String = "StoreCollarInBasePlateInfo"
   On Error GoTo ErrorHandler
        
   Dim nBasePlateCollarIndex As Long
   Dim nBasePlateCollarCount As Long
           
   nBasePlateCollarIndex = stBasePlateCollarInfo.nCollarCount
   nBasePlateCollarCount = nBasePlateCollarIndex + 1
   stBasePlateCollarInfo.nCollarCount = nBasePlateCollarCount
   
   ReDim Preserve stBasePlateCollarInfo.arrayCollar(nBasePlateCollarCount)
   Set stBasePlateCollarInfo.arrayCollar(nBasePlateCollarIndex) = oCollar
   
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub

' Collars are grouped by:
'    Assembly
'       Penetrated Plate
'          BasePlate
'
' Layout in object explorer:
'
'    AssemblyA
'        Collars
'           Collar1
'              Collar11
'              Collar12
'              ...
'           Collar2
'              Collar21
'              Collar22
'              ...
'    AssemblyB
'        Collars
'           Collar1
'              Collar11
'              Collar12
'              ...
'           Collar2
'              Collar21
'              Collar22
'              ...

' Collect collars on the same penetrated plate
' For each plate, group collars based on base plate
'
Private Sub AssignCollarParts( _
                    stUnAssignedPartInfo As UNASSIGNED_PART_INFO)
   Const sMETHOD As String = "AssignCollarParts"
   On Error GoTo ErrorHandler
                    
   Dim arrayAssemblyCollarInfo() As ASSEMBLY_COLLAR_INFO
   
   Dim nCollarIndex As Long
   Dim oCollar As Object
   
   Dim oPenetrated As Object
   Dim oPenetrating As Object
   Dim oBasePlate As Object ' Plate that penetrating profile is on
   
   Dim oAssemblyChild As IJAssemblyChild
   Dim oAssembly As IJAssembly
   Dim oNI As IJNamedItem
   
   Dim bEntryExisted As Boolean
   Dim nAssemblyCount As Long
   Dim nAssemblyIndex As Long
   Dim nPenetratedIndex As Long
   Dim nBasePlateIndex As Long
   
   nAssemblyCount = 0
   For nCollarIndex = 0 To stUnAssignedPartInfo.nCollarCount - 1
      '
      ' Each collar is associated with an assembly, penetrated plate, and a baseplate
      '
      Set oCollar = stUnAssignedPartInfo.oCollar(nCollarIndex)
      GetCollarSupportStructures _
                          oCollar, _
                          oPenetrating, _
                          oPenetrated, _
                          oBasePlate
      Set oCollar = Nothing
      Set oPenetrating = Nothing
      
      ' Check if we already have the assembly info for penetrated plate
      Set oAssemblyChild = oPenetrated
      Set oAssembly = oAssemblyChild.Parent
      Set oAssemblyChild = Nothing
      
      bEntryExisted = False
      For nAssemblyIndex = 0 To nAssemblyCount - 1
         If oAssembly Is arrayAssemblyCollarInfo(nAssemblyIndex).oAssembly Then
            bEntryExisted = True
            Exit For
         End If
      Next
      
      If bEntryExisted = False Then
         nAssemblyIndex = nAssemblyCount
         nAssemblyCount = nAssemblyCount + 1
         
         ReDim Preserve arrayAssemblyCollarInfo(nAssemblyCount)
         Set arrayAssemblyCollarInfo(nAssemblyIndex).oAssembly = oAssembly
      End If
      Set oAssembly = Nothing

      ' Get the penetrated plate index
      nPenetratedIndex = GetPenetratedPlateIndexFromAssemblyCollarInfo( _
                                  oPenetrated, _
                                  arrayAssemblyCollarInfo(nAssemblyIndex))
                                  
      If Not nPenetratedIndex < 0 Then
         nBasePlateIndex = GetBasePlateIndexFromPenetratedPlateCollarInfo( _
                  oBasePlate, _
                  arrayAssemblyCollarInfo(nAssemblyIndex).arrayPenetratedPlateCollarInfo(nPenetratedIndex))
                  
         If Not nBasePlateIndex < 0 Then
            ' Group collar by base plate
            StoreCollarInBasePlateInfo _
               stUnAssignedPartInfo.oCollar(nCollarIndex), _
               arrayAssemblyCollarInfo(nAssemblyIndex).arrayPenetratedPlateCollarInfo(nPenetratedIndex).arrayBasePlateCollarInfo(nBasePlateIndex)
         End If
      End If
      Set oBasePlate = Nothing
      Set oPenetrated = Nothing
   Next

   ' Create assemblies
   Dim oCollarsSubAssembly As IJAssembly
   Dim oCollarSubAssembly As IJAssembly
   Dim nCollarSubAssemblyCount As Long
   Dim nBasePlateCollarIndex As Long
   
   Dim stPenetratedPlateCollarInfo As PENETRATED_PLATE_COLLAR_INFO
   Dim stBasePlateCollarInfo As BASE_PLATE_COLLAR_INFO
   Dim stAssemblyCollarInfo As ASSEMBLY_COLLAR_INFO
   
   For nAssemblyIndex = 0 To nAssemblyCount - 1
      stAssemblyCollarInfo = arrayAssemblyCollarInfo(nAssemblyIndex)
      If stAssemblyCollarInfo.nPenetratePlateCount > 0 Then
         ' Create sub assembly named as Collars
         Set oCollarsSubAssembly = CreateAssembly(stAssemblyCollarInfo.oAssembly)
         If Not oCollarsSubAssembly Is Nothing Then
            Set oNI = oCollarsSubAssembly
            oNI.Name = "Collars"
            
            'apply the assembly properties rule
            SetAssemblyPropertiesByRule oCollarsSubAssembly
            
            nCollarSubAssemblyCount = 0
            For nPenetratedIndex = 0 To stAssemblyCollarInfo.nPenetratePlateCount - 1
               stPenetratedPlateCollarInfo = stAssemblyCollarInfo.arrayPenetratedPlateCollarInfo(nPenetratedIndex)
               For nBasePlateIndex = 0 To stPenetratedPlateCollarInfo.nBasePlateCount - 1
                  ' Create a sub assembly named as CollarX under Collars sub assembly
                  Set oCollarSubAssembly = CreateAssembly(oCollarsSubAssembly)
                  If Not oCollarSubAssembly Is Nothing Then
                     nCollarSubAssemblyCount = nCollarSubAssemblyCount + 1
                     Set oNI = oCollarSubAssembly
                     oNI.Name = "Collar" & nCollarSubAssemblyCount
                     
                     'apply the assembly properties rule
                     SetAssemblyPropertiesByRule oCollarSubAssembly
                     
                     stBasePlateCollarInfo = stPenetratedPlateCollarInfo.arrayBasePlateCollarInfo(nBasePlateIndex)
                     For nBasePlateCollarIndex = 0 To stBasePlateCollarInfo.nCollarCount - 1
                        oCollarSubAssembly.AddChild stBasePlateCollarInfo.arrayCollar(nBasePlateCollarIndex)
                     Next
                  End If
               Next
            Next
         End If
      End If
   Next
      
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub
'
'  Put butt connected part in the collection of parts for an assembly
'  Flag the part as assigned so it will not be processed again
'
Private Sub ProcessButtConnectedParts( _
                    ByVal oButtConnectedParts As Collection, _
                    ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO, _
                    ByRef oAssemblyParts As Collection)
   Const sMETHOD As String = "ProcessButtConnectedParts"
   On Error GoTo ErrorHandler

   Dim nConnectedIndex As Long
   Dim nPlateIndex As Long
   Dim oConnectedPart As Object
   
   For nConnectedIndex = 1 To oButtConnectedParts.Count
      Set oConnectedPart = oButtConnectedParts.Item(nConnectedIndex)
      For nPlateIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
         If oConnectedPart Is stUnAssignedPartInfo.oPlatePartInfo(nPlateIndex).oPart Then
            ' Add this part to the assembly
            oAssemblyParts.Add oConnectedPart
            
            ' Set assigned flag
            stUnAssignedPartInfo.oPlatePartInfo(nPlateIndex).bAssigned = True
            Exit For
         End If
      Next
   Next
   
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub
'
'  Put directly connected bracket parts in the collection of parts for an assembly
'  Flag the part as assigned so it will not be processed again
'
Private Sub ProcessDirectlyConnectedBracketParts( _
                                 ByVal oDirectlyConnectedBracketParts As Collection, _
                                 ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO, _
                                 ByRef oAssemblyParts As Collection)
   Const sMETHOD As String = "ProcessDirectlyConnectedBracketParts"
   On Error GoTo ErrorHandler
                                 
   Dim nConnectedPartIndex As Long
   Dim nPlatePartIndex As Long
   
   Dim oConnectedPart As Object
   Dim oStiffeningParts As Collection
   Dim oERParts As Collection
   
   ' Directly connected brackets which do not have any children
   For nConnectedPartIndex = 1 To oDirectlyConnectedBracketParts.Count
      Set oConnectedPart = oDirectlyConnectedBracketParts.Item(nConnectedPartIndex)
      GetConnectedParts _
                    CONNECTED_PART_MASK_STIFFENER Or _
                    CONNECTED_PART_MASK_ER, _
                    -1, _
                    oConnectedPart, _
                    False, _
                    stUnAssignedPartInfo, _
                    Nothing, _
                    Nothing, _
                    oStiffeningParts, _
                    oERParts, _
                    Nothing
                    
      If oStiffeningParts Is Nothing And _
         oERParts Is Nothing Then
         For nPlatePartIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
            If oConnectedPart Is stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).oPart Then
               ' Add this part to the assembly
               oAssemblyParts.Add oConnectedPart
               
               ' Set assigned flag
               stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).bAssigned = True
               Exit For
            End If
         Next
      End If
   Next
   
   Exit Sub
   
ErrorHandler:
   MsgBox Err.Description
End Sub

'
'  Put stiffening parts in the collection of parts for an assembly
'  Flag the part as assigned so it will not be processed again
'
Private Sub ProcessStiffeningParts( _
                                 ByVal eStiffenedPlateType As StructPlateType, _
                                 ByVal oStiffeningParts As Collection, _
                                 ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO, _
                                 ByRef oAssemblyParts As Collection)
   Const sMETHOD As String = "ProcessStiffeningParts"
   On Error GoTo ErrorHandler

   If oStiffeningParts Is Nothing Then
      Exit Sub
   End If
   
   Dim nConnectedPartIndex As Long
   Dim nPlatePartIndex As Long
   Dim oProfilePartIndex As Long
   
   Dim oConnectedPart As Object
   Dim stProfilePartInfo As PROFILE_PART_INFO
   
   Dim stPlatePartInfo As PLATE_PART_INFO
   Dim nSupportIndex As Long
   Dim oBracketSupport As Object
   
   Dim oStiffeningPartsTemp As Collection
   Dim oERPartsTemp As Collection
   
   ' Check each connected profile parts
   For nConnectedPartIndex = 1 To oStiffeningParts.Count
   
      Set oConnectedPart = oStiffeningParts.Item(nConnectedPartIndex)
      For oProfilePartIndex = 0 To stUnAssignedPartInfo.nProfilePartCount - 1
      
         stProfilePartInfo = stUnAssignedPartInfo.oProfilePartInfo(oProfilePartIndex)
         If oConnectedPart Is stProfilePartInfo.oPart Then
         
            ' Add further check here if needed
            
            ' Add this part to the assembly
            oAssemblyParts.Add oConnectedPart
         
            ' Set assigned flag
            stUnAssignedPartInfo.oProfilePartInfo(oProfilePartIndex).bAssigned = True
            
            ' If plate is TBH, get brackets that are connected to this profile and do NOT have children
            If eStiffenedPlateType = TBulkheadPlate Then
               For nPlatePartIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
                  stPlatePartInfo = stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex)
                  
                  If stPlatePartInfo.ePlateType = BracketPlate Then
                     For nSupportIndex = 1 To stPlatePartInfo.oBracketSupports.Count
                        Set oBracketSupport = stPlatePartInfo.oBracketSupports.Item(nSupportIndex)
                        If oBracketSupport Is stProfilePartInfo.oPart Or _
                           oBracketSupport Is stProfilePartInfo.oRootSystem Then
                           
                           ' Profile is bracket support
                           ' Check if the bracket has children
                           GetConnectedParts _
                                          CONNECTED_PART_MASK_STIFFENER Or _
                                          CONNECTED_PART_MASK_ER, _
                                          -1, _
                                          stPlatePartInfo.oPart, _
                                          False, _
                                          stUnAssignedPartInfo, _
                                          Nothing, _
                                          Nothing, _
                                          oStiffeningPartsTemp, _
                                          oERPartsTemp, _
                                          Nothing
                           If oStiffeningPartsTemp Is Nothing And _
                              oERPartsTemp Is Nothing Then
                              
                              ' Add this part to the assembly
                              oAssemblyParts.Add stPlatePartInfo.oPart
                              
                              ' Set assigned flag
                              stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).bAssigned = True
                           Else
                              Set oStiffeningPartsTemp = Nothing
                              Set oERPartsTemp = Nothing
                           End If
                           
                           Exit For
                        End If
                     Next
                  End If
               Next
            End If
            
            Exit For
         End If
      Next
   Next
   
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub
'
'  Put connected ER parts in the collection of parts for an assembly
'  Flag the part as assigned so it will not be processed again
'
Private Sub ProcessConnectedERParts( _
                                 ByVal oConnectedERParts As Collection, _
                                 ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO, _
                                 ByRef oAssemblyParts As Collection)
   On Error GoTo ErrorHandler
   Const sMETHOD As String = "ProcessConnectedERParts"
                                 
   If oConnectedERParts Is Nothing Then
      Exit Sub
   End If
   
   Dim nConnectedPartIndex As Long
   Dim nERPartIndex As Long
   Dim oConnectedPart As Object
   
   For nConnectedPartIndex = 1 To oConnectedERParts.Count
      Set oConnectedPart = oConnectedERParts.Item(nConnectedPartIndex)
      
      For nERPartIndex = 0 To stUnAssignedPartInfo.nERPartCount - 1
         If oConnectedPart Is stUnAssignedPartInfo.oERPartInfo(nERPartIndex).oPart Then
            ' Add this part to the assembly
            oAssemblyParts.Add oConnectedPart
         
            ' Set assigned flag
            stUnAssignedPartInfo.oERPartInfo(nERPartIndex).bAssigned = True
            Exit For
         End If
      Next
   Next
   
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub

'
'  Put TEE connected parts in the collection of parts for an assembly
'  Flag the part as assigned so it will not be processed again
'
Private Sub ProcessTeeConnectedPartsWithoutChildren( _
                                 ByVal oTeeConnectedParts As Collection, _
                                 ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO, _
                                 ByRef oAssemblyParts As Collection)
   Const sMETHOD As String = "ProcessTeeConnectedPartsWithoutChildren"
   On Error GoTo ErrorHandler
                                 
   If oTeeConnectedParts Is Nothing Then
      Exit Sub
   End If
   
   Dim nConnectedPartIndex As Long
   Dim nPlatePartIndex As Long
   
   Dim oConnectedPart As Object
   Dim oStiffeningParts As Collection
   Dim oERParts As Collection
   Dim oBracketParts As Collection
   
   For nConnectedPartIndex = 1 To oTeeConnectedParts.Count
      Set oConnectedPart = oTeeConnectedParts.Item(nConnectedPartIndex)
      GetConnectedParts _
                     CONNECTED_PART_MASK_STIFFENER Or _
                     CONNECTED_PART_MASK_BRACKET Or _
                     CONNECTED_PART_MASK_ER, _
                     -1, _
                     oConnectedPart, _
                     False, _
                     stUnAssignedPartInfo, _
                     Nothing, _
                     Nothing, _
                     oStiffeningParts, _
                     oERParts, _
                     oBracketParts
      If oStiffeningParts Is Nothing And _
         oERParts Is Nothing And _
         oBracketParts Is Nothing Then
         
         For nPlatePartIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
            If oConnectedPart Is stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).oPart Then
               ' Add this part to the assembly
               oAssemblyParts.Add oConnectedPart
               
               ' Set assigned flag
               stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).bAssigned = True
               Exit For
            End If
         Next
      Else
         Set oStiffeningParts = Nothing
         Set oERParts = Nothing
         Set oBracketParts = Nothing
      End If
   Next
   
   Exit Sub
   
ErrorHandler:
   MsgBox Err.Description

End Sub
'
'  Get unassigned part in a block
'  For profile parts, get their stiffened plates and root systems
'  For plate parts, calculate its area
'  For brackets, collect their supports
'
Private Sub GetUnassignedPartInfo( _
                 ByVal oBlock As IJBlock, _
                 ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO, _
                 ByRef eOption As EComboOptions)
   Const sMETHOD As String = "GetUnassignedPartInfo"
   On Error GoTo ErrorHandler

   Dim nIndex As Long
   Dim oAssembly As IJAssembly
   Dim oAssemblyBase As IJAssemblyBase
   
   Dim oPlnIntHelper          As IJDPlnIntHelper
   Dim oAssemblyChildren      As IJElements
   Dim nChildIndex            As Long
   Dim oChild As Object
   
   Dim oSDHelper As New StructDetailHelper
   Dim oSDPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim oSDProfilePartWrapper As New StructDetailObjects.ProfilePart
   Dim oSDERPartWrapper As New StructDetailObjects.EdgeReinforcement
   Dim oPlateAttributes As IJPlateAttributes
   
   Dim oStructGraph As IJStructGraph
   Dim oRootSystem As IJSystem
   Dim oPort As IJPort
   Dim oDModelBody As IJDModelBody
   Dim dAccyAchieved As Double
   Dim dLength As Double
   Dim dArea As Double
   Dim dVolume As Double

   Dim oStiffener As IJStiffener
   Dim oStiffenedPlate As Object
   Dim oStructEntityUtil As IJStructEntityUtils
   Dim oSDSmartAttributes As IJSDSmartPlateAttributes
   
   Set oStructEntityUtil = New GSCADCreateModifyUtilities.StructEntityUtils
   Set oPlnIntHelper = New CPlnIntHelper
   Set oPlateAttributes = New PlateUtils
   Set oSDSmartAttributes = New SDSmartPlateUtils
      
   If eOption = e_AllParts Then
   '
   ElseIf eOption = e_UnProcParts Then
      Set oAssemblyBase = oBlock
      Set oAssembly = oAssemblyBase.GetUnProcessedParts(oBlock)
      
      Set oAssemblyChildren = oPlnIntHelper.GetStoredProcAssemblyChildren(oAssembly, vbNullString, False)
      For nChildIndex = 1 To oAssemblyChildren.Count
         Set oChild = oAssemblyChildren.Item(nChildIndex)
         
         If TypeOf oChild Is IJStiffener Then
            Set oStiffener = oChild
            Set oStiffenedPlate = oStiffener.PlateSystem
            Set oRootSystem = Nothing
            If TypeOf oStiffenedPlate Is IJSystem Then
               Set oRootSystem = oStructEntityUtil.RootSystem(oStiffenedPlate)
               If TypeOf oRootSystem Is IJPlate Then
                  Set oStiffenedPlate = oRootSystem
               End If
            End If
         
            Set oStructGraph = oChild
            oSDHelper.IsPartDerivedFromSystem oStructGraph, oRootSystem, True
            Set oStructGraph = Nothing
            
            If TypeOf oChild Is IJProfileER Then
               stUnAssignedPartInfo.nERPartCount = stUnAssignedPartInfo.nERPartCount + 1
               ReDim Preserve stUnAssignedPartInfo.oERPartInfo(stUnAssignedPartInfo.nERPartCount)
               
               Set stUnAssignedPartInfo.oERPartInfo(stUnAssignedPartInfo.nERPartCount - 1).oPart = oChild
               Set stUnAssignedPartInfo.oERPartInfo(stUnAssignedPartInfo.nERPartCount - 1).oStiffenedPlate = oStiffenedPlate
               Set stUnAssignedPartInfo.oERPartInfo(stUnAssignedPartInfo.nERPartCount - 1).oRootSystem = oRootSystem
            Else
               stUnAssignedPartInfo.nProfilePartCount = stUnAssignedPartInfo.nProfilePartCount + 1
               ReDim Preserve stUnAssignedPartInfo.oProfilePartInfo(stUnAssignedPartInfo.nProfilePartCount)
               
               Set stUnAssignedPartInfo.oProfilePartInfo(stUnAssignedPartInfo.nProfilePartCount - 1).oPart = oChild
               Set stUnAssignedPartInfo.oProfilePartInfo(stUnAssignedPartInfo.nProfilePartCount - 1).oStiffenedPlate = oStiffenedPlate
               Set stUnAssignedPartInfo.oProfilePartInfo(stUnAssignedPartInfo.nProfilePartCount - 1).oRootSystem = oRootSystem
            End If
            
         ElseIf TypeOf oChild Is IJPlatePart Then
            Dim oPlate As IJPlate
            
            If TypeOf oChild Is IJCollarPart Then
               stUnAssignedPartInfo.nCollarCount = stUnAssignedPartInfo.nCollarCount + 1
               ReDim Preserve stUnAssignedPartInfo.oCollar(stUnAssignedPartInfo.nCollarCount - 1)
               Set stUnAssignedPartInfo.oCollar(stUnAssignedPartInfo.nCollarCount - 1) = oChild
            Else
               Set oSDPlatePartWrapper.Object = oChild
               Set oPort = oSDPlatePartWrapper.BasePort(BPT_Base)
               Set oDModelBody = oPort.Geometry
               If Not oDModelBody Is Nothing Then
                  oDModelBody.GetDimMetrics 0.01, dAccyAchieved, dLength, dArea, dVolume
               End If

               stUnAssignedPartInfo.nPlatePartCount = stUnAssignedPartInfo.nPlatePartCount + 1
               ReDim Preserve stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1)
               Set stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oPart = oChild
               
               Set oStructGraph = oChild
               oSDHelper.IsPartDerivedFromSystem oStructGraph, oRootSystem, True
               Set oStructGraph = Nothing
               Set stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oRootSystem = oRootSystem
               stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).dArea = dArea
                              
                  
               If Not oRootSystem Is Nothing Then
                  Dim bIsBracketByPlane As Boolean
                  Dim bIsTrippingBracket As Boolean
                  Dim oBracketPlane As Object

                  bIsBracketByPlane = False
                  bIsTrippingBracket = False
               
                  bIsBracketByPlane = oPlateAttributes.IsBracketByPlane(oRootSystem)
                  
                  If bIsBracketByPlane = True Then
                     Dim oUPoint As IJPoint
                     Dim oVPoint As IJPoint
                     Dim sRootSelector As String
                     
                     oPlateAttributes.GetInput_BracketByPlane _
                                             oRootSystem, _
                                             oBracketPlane, _
                                             oUPoint, _
                                             oVPoint, _
                                             sRootSelector, _
                                             stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oBracketSupports
                  Else
                     bIsTrippingBracket = oPlateAttributes.IsTrippingBracket(oRootSystem)
                     If bIsTrippingBracket = True Then
                        Dim oUVector As IJDVector
                        Dim oVVector As IJDVector
                        Dim oSupport1 As Object
                        Dim oSupport2 As Object
                        Dim oSupport3 As Object
                        Dim oLandingCurveAE As IJPlaneByElements_AE
                        
                        oPlateAttributes.GetInput_TrippingBracket _
                                             oRootSystem, _
                                             oUVector, _
                                             oVVector, _
                                             oSupport1, _
                                             oSupport2, _
                                             oSupport3, _
                                             oLandingCurveAE
                        Set stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oBracketSupports = New JObjectCollection
                        If Not oSupport1 Is Nothing Then
                           stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oBracketSupports.Add oSupport1
                        End If
                        If Not oSupport2 Is Nothing Then
                           stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oBracketSupports.Add oSupport2
                        End If
                        If Not oSupport3 Is Nothing Then
                           stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oBracketSupports.Add oSupport3
                        End If
                        
                     End If
                     
                  End If
                  
                  If bIsBracketByPlane = True Or bIsTrippingBracket = True Then
                     stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).ePlateType = BracketPlate
                  Else
                     Set oPlate = oChild
                     stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).ePlateType = oPlate.plateType
                     Set oPlate = Nothing
                  End If
                  
               Else
                  If TypeOf oChild Is IJSmartPlate Then
                     Dim oSmartPlate As IJSmartPlate
                     Dim oRefPlane As Object
                     Dim oSupports As IJElements

                     Dim oSupport As Object
                     
                     Set oSmartPlate = oChild
                     If oSmartPlate.SmartPlateType = spType_BRACKET Then
                        stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).ePlateType = BracketPlate
                        Set stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oBracketSupports = New JObjectCollection
                        
                        oSDSmartAttributes.GetInputs_Bracket _
                               oSmartPlate, _
                               oBracketPlane, _
                               oSupports
                        For nIndex = 1 To oSupports.Count
                           If TypeOf oSupports.Item(nIndex) Is IJPort Then
                              Set oPort = oSupports.Item(nIndex)
                              Set oSupport = oPort.Connectable
                           ElseIf TypeOf oSupports.Item(nIndex) Is IJConnectable Then
                              Set oSupport = oSupports.Item(nIndex)
                           End If
                           
                           If Not oSupport Is Nothing Then
                              stUnAssignedPartInfo.oPlatePartInfo(stUnAssignedPartInfo.nPlatePartCount - 1).oBracketSupports.Add oSupport
                           End If
                        Next
                     End If
                  End If
               End If
            End If
         Else
            ' Objects not handled
            
         End If
      Next
   End If

   Set oPlnIntHelper = Nothing
   
   Exit Sub
ErrorHandler:
   MsgBox Err.Description

End Sub

'
'  Get connected parts based on connected part mask
'     Butt connected parts
'     Tee connected parts
'     Stiffening parts
'     ER parts
'     Bracket parts
'
'  If nStartPartIndex is not negative,oThisPart is not passed in,
'     it will be retrieved from stUnAssignedPartInfo based on this index
'  If nStartPartIndex is negative, oThisPart must be given
'
'  If bCheckAssignedFlag is true,only un-assigned parts will be returned
'   If bCheckAssignedFlag is false, all the parts will be returned,no matter they are assigned or not
Private Function GetConnectedParts( _
                      ByVal nConnectedPartMask As Long, _
                      ByVal nStartPartIndex As Long, _
                      ByVal oThisPart As Object, _
                      ByVal bCheckAssignedFlag As Boolean, _
                      ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO, _
                      ByRef oButtConnectedParts As Collection, _
                      ByRef oTeeConnectedParts As Collection, _
                      ByRef oStiffeningParts As Collection, _
                      ByRef oERParts As Collection, _
                      ByRef oBracketParts As Collection) As Boolean
   Const sMETHOD As String = "GetConnectedParts"
   On Error GoTo ErrorHandler
   GetConnectedParts = True
   
   Dim stThisPlatePartInfo As PLATE_PART_INFO
   Dim nPlatePartIndex As Long
   Dim oButtConnectedPartsTemp As Collection
   Dim oTeeConnectedPartsTemp As Collection
   Dim oStiffeningPartsTemp As Collection
   Dim oERPartsTemp As Collection
   Dim oBracketPartsTemp As Collection
   Dim bAddIt As Boolean
   Dim oPartsFromSameRoot As Collection

   Set oButtConnectedPartsTemp = New Collection
   Set oTeeConnectedPartsTemp = New Collection
   Set oStiffeningPartsTemp = New Collection
   Set oERPartsTemp = New Collection
   Set oBracketPartsTemp = New Collection
   Set oPartsFromSameRoot = New Collection
   
   If nStartPartIndex < 0 And Not oThisPart Is Nothing Then
      For nPlatePartIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
         If oThisPart Is stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).oPart Then
            stThisPlatePartInfo = stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex)
            nStartPartIndex = 0
            Exit For
         End If
      Next
   Else
      stThisPlatePartInfo = stUnAssignedPartInfo.oPlatePartInfo(nStartPartIndex)
      If oThisPart Is Nothing Then
         Set oThisPart = stThisPlatePartInfo.oPart
      End If
   End If
   
   If (nConnectedPartMask And CONNECTED_PART_MASK_BUTT) > 0 Then
      For nPlatePartIndex = nStartPartIndex To stUnAssignedPartInfo.nPlatePartCount - 1
         Dim stOtherPlatePartInfo As PLATE_PART_INFO
         
         stOtherPlatePartInfo = stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex)
         If (Not stThisPlatePartInfo.oRootSystem Is Nothing And _
             stThisPlatePartInfo.oRootSystem Is stOtherPlatePartInfo.oRootSystem) Or _
            (stThisPlatePartInfo.oRootSystem Is Nothing And _
             stThisPlatePartInfo.oPart Is stOtherPlatePartInfo.oPart) Then
             
            bAddIt = False
            If bCheckAssignedFlag = True Then
               If stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).bAssigned = False Then
                  bAddIt = True
               End If
            Else
               bAddIt = True
            End If
            
            If bAddIt = True Then
               oButtConnectedPartsTemp.Add stOtherPlatePartInfo.oPart
            End If
            
            oPartsFromSameRoot.Add stOtherPlatePartInfo.oPart
            If stThisPlatePartInfo.oRootSystem Is Nothing Then
               Exit For
            End If
         End If
      Next
   End If
      
   If Not stThisPlatePartInfo.oRootSystem Is Nothing Then
      ' System derived part, get all the parts from the same root system
      ' Get bounded system derived plate parts, brackets
      '   Butt connections from manual logical connection
      '   Bounding connection
      '
      Dim oThisStructConnectable As IJStructConnectable
      Dim oConnectedObjectEnum As IEnumUnknown
      Dim oConnectionEnum As IEnumUnknown
      
      Set oThisStructConnectable = stThisPlatePartInfo.oRootSystem
      oThisStructConnectable.enumConnectedObjects2 _
                                      ConnectionLogical, _
                                      oConnectedObjectEnum, _
                                      oConnectionEnum
                                      
      If Not oConnectedObjectEnum Is Nothing Then
         Dim oConvertUtils As New CCollectionConversions
         Dim oConnectedObjects As Collection
         Dim oConnections As Collection
         Dim nConnIndex As Long
         Dim oStructConnection As IJStructConnection
         
         Set oConnectedObjects = New Collection
         Set oConnections = New Collection
         
         oConvertUtils.CreateVBCollectionFromIEnumUnknown oConnectedObjectEnum, oConnectedObjects
         oConvertUtils.CreateVBCollectionFromIEnumUnknown oConnectionEnum, oConnections
         
         Dim oPorts As IJElements
         Dim oPort As IJPort
         Dim oThisPartPort As IJPort
         Dim oConnectedPort As IJPort
         Dim oConnectedObject As Object
            
         For nConnIndex = 1 To oConnections.Count
            Set oStructConnection = oConnections.Item(nConnIndex)
            oStructConnection.enumPortsLive oPorts

            If oPorts.Count = 2 Then
               Set oPort = oPorts.Item(1)
               If oPort.Connectable Is stThisPlatePartInfo.oRootSystem Then
                  Set oThisPartPort = oPort
                  Set oConnectedPort = oPorts.Item(2)
               Else
                  Set oThisPartPort = oPorts.Item(2)
                  Set oConnectedPort = oPort
               End If
               
               Set oConnectedObject = oConnectedObjects.Item(nConnIndex)
               
               If TypeOf oConnectedPort.Connectable Is IJPlate Then
                  If oThisPartPort.Type = PortEdge Then
                  
                     If oConnectedPort.Type = PortEdge Then
                        ' Manual connection
                        If (nConnectedPartMask And CONNECTED_PART_MASK_BUTT) > 0 Then
                           For nPlatePartIndex = nStartPartIndex + 1 To stUnAssignedPartInfo.nPlatePartCount - 1
                              stOtherPlatePartInfo = stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex)

                              Dim sOtherRootName As String

                              If oConnectedObject Is stOtherPlatePartInfo.oRootSystem Then
                                 bAddIt = False
                                 If bCheckAssignedFlag = True Then
                                    If stOtherPlatePartInfo.bAssigned = False Then
                                       bAddIt = True
                                    End If
                                 Else
                                    bAddIt = True
                                 End If
                                 
                                 If bAddIt = True Then
                                    oButtConnectedPartsTemp.Add stOtherPlatePartInfo.oPart
                                 End If
                              End If
                           Next
                        End If
                     End If
                  ElseIf oThisPartPort.Type = PortFace Then
                     If oConnectedPort.Type = PortEdge Then
                        ' Bounded connection
                        If (nConnectedPartMask And CONNECTED_PART_MASK_TEE) > 0 Then
                           For nPlatePartIndex = nStartPartIndex + 1 To stUnAssignedPartInfo.nPlatePartCount - 1
                              stOtherPlatePartInfo = stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex)
                              If oConnectedObject Is stOtherPlatePartInfo.oRootSystem Then
                                 bAddIt = False
                                 
                                 If stOtherPlatePartInfo.ePlateType <> BracketPlate And _
                                    stOtherPlatePartInfo.ePlateType <> DeckPlate And _
                                    stOtherPlatePartInfo.ePlateType <> Hull Then
                                    '
                                    ' Note:
                                    '   1.Bracket will be processd later.
                                    '     For a lapped tripping stiffener, there is no connection between it and support plate
                                    '   2.Deck and Hull plate parts are excluded from Tee Connected part,
                                    '     no matter deck or hull is boundary or is bounded
                                    If bCheckAssignedFlag = True Then
                                       If stOtherPlatePartInfo.bAssigned = False Then
                                          bAddIt = True
                                       End If
                                    Else
                                       bAddIt = True
                                    End If
                                 End If
                                 
                                 If bAddIt = True Then
                                    oTeeConnectedPartsTemp.Add stOtherPlatePartInfo.oPart
                                 End If
                              End If
                           Next
                        End If
                     End If
                  End If
               Else
                  ' Connected is profile
                  '   Stiffening profile is handled later
                  '   Boundary profile is ignored for this plate (It's handled with its OWN stiffened plate)
               End If
            End If
         Next
      End If
   End If
   
   Dim oSDPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim oPlatePartBoundaries As Collection
   Dim nBoundaryIndex As Long
   Dim nIndex As Long
   Dim oBoundaryRootSystem As Object
   Dim oSDHelper As New StructDetailHelper
   Dim stBoundaryData As BoundaryData
   Dim stProfilePartInfo As PROFILE_PART_INFO
   Dim bCheck As Boolean
   
   '
   ' Tee connected stand alone plate parts
   ' Stand alone brackets
   '
   Dim oSDPlatePartAttributes As IJSDPlatePartAttributes
   Dim eSAPlateType As SAPartTypes
   Dim oSmartPlate As IJSmartPlate
   Dim bIsSupport As Boolean
   
   Set oSDPlatePartAttributes = New SDPlatePartUtils

   If nConnectedPartMask And CONNECTED_PART_MASK_BRACKET > 0 Or _
      nConnectedPartMask And CONNECTED_PART_MASK_TEE > 0 Then
      
      For nPlatePartIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
         stOtherPlatePartInfo = stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex)
         
         If Not stThisPlatePartInfo.oPart Is stOtherPlatePartInfo.oPart Then
            If stOtherPlatePartInfo.ePlateType = BracketPlate Then
               If nConnectedPartMask And CONNECTED_PART_MASK_BRACKET > 0 And Not stOtherPlatePartInfo.oBracketSupports Is Nothing Then
                  For nIndex = 1 To stOtherPlatePartInfo.oBracketSupports.Count
                     If (Not stThisPlatePartInfo.oRootSystem Is Nothing And _
                         stThisPlatePartInfo.oRootSystem Is stOtherPlatePartInfo.oBracketSupports.Item(nIndex)) Or _
                        (stThisPlatePartInfo.oRootSystem Is Nothing And _
                         stThisPlatePartInfo.oPart Is stOtherPlatePartInfo.oBracketSupports.Item(nIndex)) Then
                         
                        bAddIt = False
                        If bCheckAssignedFlag = True Then
                           If stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).bAssigned = False Then
                              bAddIt = True
                           End If
                        Else
                           bAddIt = True
                        End If
                        
                        If bAddIt = True Then
                           oBracketPartsTemp.Add stOtherPlatePartInfo.oPart
                        End If
                     End If
                  Next
               End If
            Else
               If nConnectedPartMask And CONNECTED_PART_MASK_TEE > 0 And _
                  TypeOf stOtherPlatePartInfo.oPart Is IJDStandAlonePlatePart Then
                  
                  eSAPlateType = oSDPlatePartAttributes.GetSAPlatePartType(stOtherPlatePartInfo.oPart)
                  If Not eSAPlateType = SD_STAND_ALONE_PlanarPlateERPart Then
                     ' Note:PlatePart.PlateBoundaries does not work properly for ER plate part,
                     '      exclude ER plate parts, they can not be bounded anyway
                     Set oSDPlatePartWrapper.Object = stOtherPlatePartInfo.oPart
                     Set oPlatePartBoundaries = oSDPlatePartWrapper.PlateBoundaries()
                     For nBoundaryIndex = 1 To oPlatePartBoundaries.Count
                        stBoundaryData = oPlatePartBoundaries(nBoundaryIndex)
                        For nIndex = 1 To oPartsFromSameRoot.Count
                           If TypeOf stBoundaryData.Boundary Is IJPlate And _
                              stBoundaryData.Boundary Is oPartsFromSameRoot.Item(nIndex) Then
                              bAddIt = False
                              If bCheckAssignedFlag = True Then
                                 If stUnAssignedPartInfo.oPlatePartInfo(nPlatePartIndex).bAssigned = False Then
                                    bAddIt = True
                                 End If
                              Else
                                 bAddIt = True
                              End If
                              
                              If bAddIt = True Then
                                 oTeeConnectedPartsTemp.Add stOtherPlatePartInfo.oPart
                              End If
                           End If
                        Next
                     Next
                  End If
               End If
            End If
         End If
      Next
   End If
   
   ' Get stiffening profile parts,including system derived and stand alone profile parts
   If (nConnectedPartMask And CONNECTED_PART_MASK_STIFFENER) > 0 Then
      For nIndex = 0 To stUnAssignedPartInfo.nProfilePartCount - 1
         stProfilePartInfo = stUnAssignedPartInfo.oProfilePartInfo(nIndex)
         
         If stThisPlatePartInfo.oRootSystem Is stProfilePartInfo.oStiffenedPlate Or _
            stThisPlatePartInfo.oPart Is stProfilePartInfo.oStiffenedPlate Then
            bAddIt = False
            If bCheckAssignedFlag = True Then
               If stProfilePartInfo.bAssigned = False Then
                  bAddIt = True
               End If
            Else
               bAddIt = True
            End If
            
            If bAddIt = True Then
               oStiffeningPartsTemp.Add stProfilePartInfo.oPart
            End If
         End If
      Next
   End If
   
   ' ER parts
   If (nConnectedPartMask And CONNECTED_PART_MASK_ER) > 0 Then
      For nIndex = 0 To stUnAssignedPartInfo.nERPartCount - 1
         stProfilePartInfo = stUnAssignedPartInfo.oERPartInfo(nIndex)
         If stThisPlatePartInfo.oRootSystem Is stProfilePartInfo.oStiffenedPlate Or _
            stThisPlatePartInfo.oPart Is stProfilePartInfo.oStiffenedPlate Then
            bAddIt = False
            If bCheckAssignedFlag = True Then
               If stProfilePartInfo.bAssigned = False Then
                  bAddIt = True
               End If
            Else
               bAddIt = True
            End If
            
            If bAddIt = True Then
               oERPartsTemp.Add stProfilePartInfo.oPart
            End If
         End If
      Next
   End If

   ' Set retuned collection
   If oButtConnectedPartsTemp.Count > 0 Then
      Set oButtConnectedParts = oButtConnectedPartsTemp
   Else
      Set oButtConnectedParts = Nothing
   End If
   
   If oTeeConnectedPartsTemp.Count > 0 Then
      Set oTeeConnectedParts = oTeeConnectedPartsTemp
   Else
      Set oTeeConnectedParts = Nothing
   End If
                      
   If oStiffeningPartsTemp.Count > 0 Then
      Set oStiffeningParts = oStiffeningPartsTemp
   Else
      Set oStiffeningParts = Nothing
   End If
                      
   If oERPartsTemp.Count > 0 Then
      Set oERParts = oERPartsTemp
   Else
      Set oERParts = Nothing
   End If
                      
   If oBracketPartsTemp.Count > 0 Then
      Set oBracketParts = oBracketPartsTemp
   Else
      Set oBracketParts = Nothing
   End If
   
   Exit Function
   
ErrorHandler:
   GetConnectedParts = False
End Function


'Sets the properties on the given assembly based on the AssemblyProperties rule
Private Sub SetAssemblyPropertiesByRule(oAssembly As IJAssembly)
    On Error GoTo ErrHandler
    Dim oAssemblyHelper As PlnAssembly
    oAssemblyHelper = New PlnAssembly

    oAssemblyHelper = oAssembly
    oAssemblyHelper.SetAssemblyPropertiesByRule
    Exit Sub
ErrHandler:
    Err.Clear
End Sub
