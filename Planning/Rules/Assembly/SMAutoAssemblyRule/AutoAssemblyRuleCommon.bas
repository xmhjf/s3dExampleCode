Attribute VB_Name = "AutoAssemblyRuleCommon"
Option Explicit

Public Const CUSTOMERID = "SM"
Public Const m_sProjectName As String = CUSTOMERID + "AutoAssemblyRule"
Public Const m_sProjectPath As String = "M:\SharedContent\Src\Planning\Rules\Assembly\" + m_sProjectName + "\"

Public Const IID_IJAssemblyChild As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"

Public Const CONNECTED_PART_MASK_ALL As Long = &HFFFFFFFF
Public Const CONNECTED_PART_MASK_BUTT As Long = &H1
Public Const CONNECTED_PART_MASK_TEE As Long = &H2
Public Const CONNECTED_PART_MASK_STIFFENER As Long = &H4
Public Const CONNECTED_PART_MASK_BRACKET As Long = &H8
Public Const CONNECTED_PART_MASK_ER As Long = &H10

Public Type BASE_PLATE_COLLAR_INFO
   oBasePlate As Object
   nCollarCount As Long
   arrayCollar() As Object
End Type

Public Type PENETRATED_PLATE_COLLAR_INFO
   oPenetratedPart As Object
   oPenetratedRoot As Object
   nBasePlateCount As Long
   arrayBasePlateCollarInfo() As BASE_PLATE_COLLAR_INFO
End Type

Public Type ASSEMBLY_COLLAR_INFO
   oAssembly As IJAssembly
   nPenetratePlateCount As Long
   arrayPenetratedPlateCollarInfo() As PENETRATED_PLATE_COLLAR_INFO
End Type

Public Type PLATE_PART_INFO ' Brackets also use this
   oPart As Object
   bAssigned As Boolean
   dArea As Double
   oRootSystem As Object
   ePlateType As StructPlateType
   oBracketSupports As IJElements ' For bracket use only
End Type

Public Type PROFILE_PART_INFO
   oPart As Object
   bAssigned As Double
   oRootSystem As Object
   oStiffenedPlate As Object
End Type

Public Type UNASSIGNED_PART_INFO
   oPlatePartInfo() As PLATE_PART_INFO ' Including bracket
   oCollar() As Object
   
   oProfilePartInfo() As PROFILE_PART_INFO
   oERPartInfo() As PROFILE_PART_INFO
   
   nPlatePartCount As Long
   nCollarCount As Long
   
   nProfilePartCount As Long
   nERPartCount As Long
End Type

Private Const m_sModule As String = m_sProjectPath + "AutoAssemblyRuleCommon.bas"


Public Sub LogPartInfo(stUnAssignedPartInfo As UNASSIGNED_PART_INFO)
   Dim oNI As IJNamedItem
   Dim strFileName As String
   Dim nFileNumber As Integer
   Dim sName As String
   Dim nIndex As Long
   
   strFileName = "c:\temp\UnsignedPart.txt"
   nFileNumber = FreeFile
   Open strFileName For Append As nFileNumber
   
   Print #nFileNumber, "Plate part,including bracket: " & stUnAssignedPartInfo.nPlatePartCount
   For nIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
      Set oNI = stUnAssignedPartInfo.oPlatePartInfo(nIndex).oPart
      sName = oNI.Name
      Print #nFileNumber, sName & ", area = " & stUnAssignedPartInfo.oPlatePartInfo(nIndex).dArea
   Next
   Print #nFileNumber, vbNewLine
      
   Print #nFileNumber, "Collar: " & stUnAssignedPartInfo.nCollarCount
   For nIndex = 0 To stUnAssignedPartInfo.nCollarCount - 1
      Set oNI = stUnAssignedPartInfo.oCollar(nIndex)
      sName = oNI.Name
   Next
   Print #nFileNumber, vbNewLine
   
   Print #nFileNumber, "Stiffener: " & stUnAssignedPartInfo.nProfilePartCount & vbNewLine
   For nIndex = 0 To stUnAssignedPartInfo.nProfilePartCount - 1
      Set oNI = stUnAssignedPartInfo.oProfilePartInfo(nIndex).oPart
      sName = oNI.Name
      Print #nFileNumber, sName
   Next
   Print #nFileNumber, vbNewLine
      
   Print #nFileNumber, "ER: " & stUnAssignedPartInfo.nERPartCount & vbNewLine
   For nIndex = 0 To stUnAssignedPartInfo.nERPartCount - 1
      Set oNI = stUnAssignedPartInfo.oERPartInfo(nIndex).oPart
      sName = oNI.Name
      Print #nFileNumber, sName
   Next

   Close #nFileNumber
End Sub

Public Sub LogAssemblyMembers(oAssemblyMember As Collection)

   If oAssemblyMember Is Nothing Then
      Exit Sub
   End If
   
   Dim oNI As IJNamedItem
   Dim strFileName As String
   Dim nFileNumber As Integer
   Dim sName As String
   Dim nIndex As Long

   strFileName = "c:\temp\AssemblyMember.txt"
   nFileNumber = FreeFile
   Open strFileName For Append As nFileNumber
      
   Print #nFileNumber, "Assembly Members: "
   For nIndex = 1 To oAssemblyMember.Count
      Set oNI = oAssemblyMember.Item(nIndex)
      sName = oNI.Name
      Print #nFileNumber, "        " & sName
   Next
   Print #nFileNumber, vbNewLine
   
   Close #nFileNumber
End Sub

Public Sub SortPlatePartByArea(ByRef stUnAssignedPartInfo As UNASSIGNED_PART_INFO)
   Dim nOuterIndex As Long
   Dim nInnerIndex As Long
   Dim nToSwapIndex As Long
   
   Dim stTempPartInfo As PLATE_PART_INFO
   Dim dMaxArea As Double
   
   Dim oGraphicRange As IJRangeAlias
   Dim dLowX As Double
   Dim dLowY As Double
   Dim dLowZ As Double
   
   Dim dHiX As Double
   Dim dHiY As Double
   Dim dHiZ As Double
   Dim dOuterY As Double
   Dim dInnerY As Double
   Dim oBox  As GBox
   
   For nOuterIndex = 0 To stUnAssignedPartInfo.nPlatePartCount - 1
      nToSwapIndex = nOuterIndex
      dMaxArea = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).dArea
      For nInnerIndex = nOuterIndex + 1 To stUnAssignedPartInfo.nPlatePartCount - 1
         If dMaxArea < stUnAssignedPartInfo.oPlatePartInfo(nInnerIndex).dArea Then
            dMaxArea = stUnAssignedPartInfo.oPlatePartInfo(nInnerIndex).dArea
            nToSwapIndex = nInnerIndex
         ElseIf Abs(dMaxArea - stUnAssignedPartInfo.oPlatePartInfo(nInnerIndex).dArea) < 0.00001 Then
            Set oGraphicRange = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).oPart
            oBox = oGraphicRange.GetRange
            
            dLowX = oBox.m_low.X
            dLowY = oBox.m_low.Y
            dLowZ = oBox.m_low.Z
            dHiX = oBox.m_high.X
            dHiY = oBox.m_high.Y
            dHiZ = oBox.m_high.Z
            
            Set oGraphicRange = Nothing
            
            Set oGraphicRange = stUnAssignedPartInfo.oPlatePartInfo(nInnerIndex).oPart
            oBox = oGraphicRange.GetRange
            
            dLowX = oBox.m_low.X
            dLowY = oBox.m_low.Y
            dLowZ = oBox.m_low.Z
            dHiX = oBox.m_high.X
            dHiY = oBox.m_high.Y
            dHiZ = oBox.m_high.Z
            
            dOuterY = (dLowY + dHiY) / 2
            dInnerY = (dLowY + dHiY) / 2
            
            If dOuterY < 0 And dInnerY > 0 Then
               nToSwapIndex = nInnerIndex
            End If
         End If
      Next
      
      If nToSwapIndex <> nOuterIndex Then
         ' Need to swap
         '
         ' Move nOuterIndex content to temporary storage
         Set stTempPartInfo.oPart = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).oPart
         stTempPartInfo.dArea = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).dArea
         stTempPartInfo.bAssigned = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).bAssigned
         stTempPartInfo.ePlateType = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).ePlateType
         Set stTempPartInfo.oRootSystem = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).oRootSystem
         Set stTempPartInfo.oBracketSupports = stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).oBracketSupports
         
         ' Move nToSwapIndex content to nOuterIndex
         stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).bAssigned = stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).bAssigned
         stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).dArea = stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).dArea
         stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).ePlateType = stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).ePlateType
         Set stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).oPart = stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).oPart
         Set stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).oRootSystem = stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).oRootSystem
         Set stUnAssignedPartInfo.oPlatePartInfo(nOuterIndex).oBracketSupports = stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).oBracketSupports
                  
         ' Move stored content to nToSwapIndex
         stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).bAssigned = stTempPartInfo.bAssigned
         stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).dArea = stTempPartInfo.dArea
         stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).ePlateType = stTempPartInfo.ePlateType
         Set stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).oPart = stTempPartInfo.oPart
         Set stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).oRootSystem = stTempPartInfo.oRootSystem
         Set stUnAssignedPartInfo.oPlatePartInfo(nToSwapIndex).oBracketSupports = stTempPartInfo.oBracketSupports
      End If
   Next
End Sub

Public Function CreateAssembly( _
         ByVal oParent As GSCADAssembly.IJAssembly) As GSCADAssembly.IJAssembly
         
    Const sMETHOD As String = "CreateAssembly"
    On Error GoTo ErrorHandler
    
    Dim oFactory As GSCADAssembly.IJAssemblyFactory
    Dim oNewAssembly As GSCADAssembly.IJAssembly
    
    Set oFactory = New GSCADAssembly.AssemblyFactory
    Set oNewAssembly = oFactory.CreateAssembly( _
                                   GetPOM("Model"), _
                                   oParent)
    Set oFactory = Nothing
    
    Set CreateAssembly = oNewAssembly
    
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
End Function

Public Function SetAssemblyNamingRule( _
             oAssembly As IJAssembly, _
             sRuleName As String) As Boolean
             
   SetAssemblyNamingRule = False
   
   Dim oNamingRuleHelper As IJDNamingRulesHelper
   Dim oRules As IJElements
   Dim oNamingRule As GSCADGenericNamingRulesFacelets.NameRuleHolderFacelets
    
   Set oNamingRuleHelper = New GSCADNameRuleHlpr.NamingRulesHelper
    
   On Error Resume Next
   oNamingRuleHelper.GetEntityNamingRulesGivenName sRuleName, oRules

   If oRules Is Nothing Then
      ' No such a rule, the assembly will not be named!
      Exit Function
   End If
   
   ' Iterate the rules and select the correct one
   ' For now, use first one.
   Set oNamingRule = oRules.Item(1)
   Set oRules = Nothing
   If Not oNamingRule Is Nothing Then
       Dim oNameRuleAE As GSCADGenNameRuleAE.IJNameRuleAE
       
       oNamingRuleHelper.AddNamingRelations oAssembly, oNamingRule, oNameRuleAE
       Set oNamingRule = Nothing
       Set oNameRuleAE = Nothing
       SetAssemblyNamingRule = True
   End If
   
ErrorHandler:
   Set oNamingRuleHelper = Nothing
   
   Exit Function
   
End Function

'
' This method is not really needed. However, when using collar wrapper to do the same,
' always get ref count problem. Have no time to investigate.
'
Public Sub GetCollarSupportStructures( _
         ByVal oCollar, _
         ByRef oPenetrating As Object, _
         ByRef oPenetrated As Object, _
         ByRef oBasePlate As Object)
         
   If oCollar Is Nothing Then
      Exit Sub
   End If
   
   ' Get penetrating and base plate
   Dim oSlot As Object
   Dim sRootName As String
   Dim oSDCollarAttributes As IJSDCollarAttributes
   
   Set oSDCollarAttributes = New SDCollarUtils

    
   oSDCollarAttributes.GetInput_Collar _
                               oCollar, _
                               oBasePlate, _
                               oPenetrating, _
                               sRootName, _
                               oSlot
   Set oCollar = Nothing
   Set oSDCollarAttributes = Nothing

   ' Get penetrated
   Dim oSDFeatureUtils As IJSDFeatureAttributes
   Dim oSlotPenetrating As Object
   Dim oSlotPenetratedGeom As Object
   Dim oSlotBasePlate As Object
   
   Set oSDFeatureUtils = New SDFeatureUtils
   oSDFeatureUtils.get_SlotInputs _
                             oSlot, _
                             oSlotPenetrating, _
                             oSlotPenetratedGeom, _
                             oSlotBasePlate
   Set oSDFeatureUtils = Nothing
   Set oSlot = Nothing
   Set oSlotPenetrating = Nothing
   Set oSlotBasePlate = Nothing
   
   ' The penetrated object input is the geometry, change it to the part
   Dim oStructGeometryHelper As StructGeometryHelper
   
   Set oStructGeometryHelper = New StructGeometryHelper
   oStructGeometryHelper.RecursiveGetStructEntityIUnknown _
                           oSlotPenetratedGeom, _
                           oPenetrated
   Set oStructGeometryHelper = Nothing
   Set oSlotPenetratedGeom = Nothing
End Sub

Public Function GetPOM(strDbType As String) As IUnknown
Const sMETHOD = "GetPOM"
On Error GoTo ErrHandler
    Dim oContext As IJContext
    Dim oAccessMiddle As IJDAccessMiddle

    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
 
    Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

