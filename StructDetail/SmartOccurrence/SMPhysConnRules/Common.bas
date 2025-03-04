Attribute VB_Name = "Common"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   File:           Common.bas
'   Author:         Alligators Team(India)
'   Creation Date:  Tuesday, Nov 16 2010
'   Description:
'      This module is used for Tee weld selection smart item.
'
'   Change History:
'   dd.mmm.yyyy     who          change description
'   -----------     ---          ------------------
'   16.Nov.2010     svsmylav     TR-187029: GetMoldedSide function is modified to
'                                work with PC which is neither a web cut nor a flange cut.
'
'   2.July.2012     vbbheema    TR-208253: Removed the message boxes fron
'                                GetMoldedSide method as part fix for TR-208253

'   6.July.2015     pkakula     TR-274819: Added a null check in GetParentPhysicalConnectionSelectorDefinition() for oSmartitem.
'                               This check is to avoid record exceptions dump during the run of Marine SOM ATPs.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Const m_sProjectID = CUSTOMERID + "PhysConnRules"

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SMPhysConnRules\Common.bas"
Private Const IID_IJStructGeometry As String = "{6034AD40-FA0B-11d1-B2FD-080036024603}"
Private Const IID_IJStructSplit As String = "{0A31DCF2-45EB-11d5-8126-00105AE5AAE5}"

Public oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate

' enum of values to indicate which side of a plate connection is chamfered while processing
' butt-weld parameters.
Public Enum pcrChamferedSidesInfo
    pcr_NoChamfer = 1
    pcr_BaseChamfer = 2
    pcr_OffsetChamfer = 3
    pcr_BothChamfer = 4
End Enum

' enum of values to define the weld groove types used in IJWeldSymbol.  This is a duplicate
' definition of the "official" CodeList definition, which can be found in
' M:\CommonSchema\Middle\Schema\Codelists\CommonSchemaCodelists.xls
Public Enum pcrWeldGroove
    pcr_WG_None = 0
    pcr_WG_Square = 1
    pcr_WG_V = 2
    pcr_WG_Bevel = 3
    pcr_WG_U = 4
    pcr_WG_J = 5
    pcr_WG_FlareV = 6
    pcr_WG_FlareBevel = 7
End Enum

Private Const WELD_SYMBOL_NONE = 0
Private Const WELD_SYMBOL_FILLET = 1
Private Const WELD_SYMBOL_STAGGERED_FILLET = 2

Public Function GetRefSide(oPCObj As Object, Optional oConnectObj As Object) As String

    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim sMoldedSide As String
    
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = oPCObj
    
    If oConnectObj Is Nothing Then
        Set oConnectObj = oPhysConn.ConnectedObject1
    End If

    Dim pHelper As New StructDetailObjects.Helper

    Select Case pHelper.ObjectType(oConnectObj)
        Case SDOBJECT_PLATE

            Dim oWeldPlate As New StructDetailObjects.PlatePart
            Set oWeldPlate.object = oConnectObj
            sMoldedSide = oWeldPlate.MoldedSide

        Case SDOBJECT_STIFFENER
            Dim oSysChild As IJDesignChild
            Set oSysChild = oPhysConn.object
            
            If TypeOf oSysChild.GetParent Is IJStructFeature Then
                Dim objWebOrFlange As IJStructFeature
                Set objWebOrFlange = oSysChild.GetParent
                If objWebOrFlange.get_StructFeatureType = SF_WebCut Then
                    Dim oSDProfile As New StructDetailObjects.ProfilePart
                    Set oSDProfile.object = oConnectObj
                    
                    If oSDProfile.sectionType = "HalfR" Then
                          sMoldedSide = "Top"
                    ElseIf oSDProfile.sectionType = "R" Then
                          sMoldedSide = "Outer"
                    Else
                          sMoldedSide = "WebLeft"
                    End If
                    Set oSDProfile = Nothing
                    
                ElseIf objWebOrFlange.get_StructFeatureType = SF_FlangeCut Then
                    Dim oFlangeObj As New StructDetailObjects.FlangeCut
                    Set oFlangeObj.object = objWebOrFlange
                    Dim bIsTopFlange As Boolean
                    bIsTopFlange = oFlangeObj.IsTopFlange
                    If bIsTopFlange Then
                        sMoldedSide = "TopFlangeTopFace"
                    Else
                        sMoldedSide = "BottomFlangeBottomFace"
                    End If
                End If
            Else
                sMoldedSide = "WebLeft"
            End If
            
        Case SDOBJECT_BEAM
            sMoldedSide = "WebLeft"
    End Select
    
'    End If
    
    GetRefSide = sMoldedSide

  Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetRefSide", strError).Number
End Function

'------------------------------------------------------------------------------------------------------------
' Procedure (Function):
'     DegToRad (Double)
'
' Description:
'     Converts an angle measure in degrees to its equivalent in radians.
'
' Arguments:
'     dAngle    Double    The angle measure in degrees.
'------------------------------------------------------------------------------------------------------------
Public Function DegToRad(dAngle As Double) As Double

    Const PI As Double = 3.141592654

    DegToRad = dAngle * PI / 180  'Radians=Degrees*Pi/180
End Function

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     Get_ParameterRuleData
'
' Description:
'   Retrieves Parameter Rule data required by
'       all Tee Welds, Chain Weld, and ZigZag Weld Parameter Rules
'
'   The Data is retrieved from
'       the "Standard" Tee Weld selectors
'   Or
'       the "Chamfer" Tee Weld selectors
'
' Arguments:
'------------------------------------------------------------------------------------------------------------
Public Sub Get_ParameterRuleData(pPRL As IJDParameterLogic, _
                                 sStandardItemName As String, _
                                 sClassSociety As String, _
                                 sCategory As String, _
                                 sBevelMethod As String, _
                                 dThickness1 As Double, _
                                 dThickness2 As Double)
    On Error GoTo ErrorHandler
    Dim dChamferThickness As Double
  
    ' Get Class Arguments
    Dim oPhysConn As StructDetailObjects.PhysicalConn
    Set oPhysConn = New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = pPRL.SmartOccurrence
  
    ' Check where the Parameter Rule data is to be Retrieve From
    ' Is this a "Standard" Physical Connection or a Chamfer Physical Connection
    If LCase(Trim(pPRL.SmartItem.Name)) = LCase(Trim(sStandardItemName)) Then
        GetSelectorAnswer pPRL, "Category", sCategory
        GetSelectorAnswer pPRL, "BevelAngleMethod", sBevelMethod
        GetSelectorAnswer pPRL, "ClassSociety", sClassSociety
        
        GetPhysicalConnPartsThickness pPRL.SmartOccurrence, _
                                      dThickness1, _
                                      dThickness2
        If dThickness1 < 0.000001 Then
           dThickness1 = oPhysConn.Object1Thickness
        End If
        If dThickness2 < 0.000001 Then
           dThickness2 = oPhysConn.Object2Thickness
        End If
        
    Else
        ' Get data from "Chamfer" selector
        GetSelectorAnswer pPRL, "Category", sCategory
        GetSelectorAnswer pPRL, "BevelAngleMethod", sBevelMethod
        GetSelectorAnswer pPRL, "ClassSociety", sClassSociety
         
        dChamferThickness = pPRL.SelectorAnswer(CUSTOMERID + "PhysConnRules.ChamferTeeWeldSel", "ChamferThickness")
        dThickness1 = dChamferThickness
        dThickness2 = oPhysConn.Object2Thickness
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Get_ParameterRuleData").Number
End Sub

Public Function GetParentPhysicalConnectionSelectorDefinition(ByRef oPCParent As Object) As IJDSymbolDefinition
    
    On Error GoTo ErrorHandler
    
    If Not oPCParent Is Nothing Then
    
        If TypeOf oPCParent Is IJStructPhysicalConnection Then
    
            If TypeOf oPCParent Is IJSmartOccurrence Then
                
                Dim oSmartItem As IJSmartItem
                Dim oSmartParent As Object
                Dim oSmartOccurrence As IJSmartOccurrence
                
                Set oSmartOccurrence = oPCParent
                Set oSmartItem = oSmartOccurrence.ItemObject
                    
                If Not oSmartItem Is Nothing Then
                
                    Set oSmartParent = oSmartItem.Parent
                    If TypeOf oSmartParent Is IJSmartClass Then
                        
                        Dim osmartclass As IJSmartClass
                        Dim oSelectorDefinition As IJDSymbolDefinition
                        
                        Set osmartclass = oSmartParent
                        Set osmartclass = osmartclass.Parent
                        If Not osmartclass Is Nothing Then
                            
                            Set oSelectorDefinition = osmartclass.SelectionRuleDef
                            
                            If Not oSelectorDefinition Is Nothing Then
                                
                                Set GetParentPhysicalConnectionSelectorDefinition = oSelectorDefinition
                            End If
                        End If
                    End If
                                
                End If
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetParentPhysicalConnectionSelectorDefinition").Number
End Function
Public Sub SetAnswerFromParentPhysicalConnection(ByRef oSymbolDefinition As IJDSymbolDefinition, ByVal AnswerName As String)
    
    On Error GoTo ErrorHandler
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
        
    ' Get the parent PC
    Dim oPCParent As Object
    Dim oSystemChild As IJSystemChild
    
    Set oSystemChild = pSL.SmartOccurrence
    
    Set oPCParent = oSystemChild.GetParent
    
    If Not oPCParent Is Nothing Then
        Dim oSelectorDefinition As IJDSymbolDefinition
        
        Set oSelectorDefinition = GetParentPhysicalConnectionSelectorDefinition(oPCParent)
        
        If Not oSelectorDefinition Is Nothing Then
            Dim oCommonHelper As DefinitionHlprs.CommonHelper
            Set oCommonHelper = New DefinitionHlprs.CommonHelper
            
            pSL.Answer(AnswerName) = oCommonHelper.GetAnswer(oPCParent, oSelectorDefinition, AnswerName)
                
            
        End If
        
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetAnswerFromParentPhysicalConnection").Number
End Sub

'''''''''
Public Sub GetPhysicalConnPartsThickness( _
        ByVal oPhysConn As Object, _
        ByRef dObject1Thickness As Double, _
        ByRef dObject2Thickness As Double)
   On Error GoTo ErrorHandler
        
   Dim bIsParentEndCut As Boolean
   
   dObject1Thickness = 0
   dObject2Thickness = 0
   bIsParentEndCut = False
   If Not TypeOf oPhysConn Is IJDesignChild Then
      Exit Sub
   End If
   
   Dim oDesignChild As IJDesignChild
   Dim oPCParent As Object
   
   Set oDesignChild = oPhysConn
   Set oPCParent = oDesignChild.GetParent
   Set oDesignChild = Nothing
   If TypeOf oPCParent Is IJStructFeature Then
      Dim eFeatureType As StructFeatureTypes
      Dim oStructFeature As IJStructFeature
      
      Set oStructFeature = oPCParent
      eFeatureType = oStructFeature.get_StructFeatureType
      If eFeatureType = SF_WebCut Or _
         eFeatureType = SF_FlangeCut Then
         bIsParentEndCut = True
      End If
   End If
   Set oPCParent = Nothing
   
   Dim oPCWrapper As New StructDetailObjects.PhysicalConn
   
   Set oPCWrapper.object = oPhysConn
   If bIsParentEndCut = False Then
      dObject1Thickness = oPCWrapper.Object1Thickness
      dObject2Thickness = oPCWrapper.Object2Thickness
      
   Else
      ' PC under end cut,check if it is known Multiple PCs case
      Dim oSO As IJSmartOccurrence
      Dim oSI As IJSmartItem
      Dim sItemName As String
      
      Set oSO = oStructFeature
      Set oStructFeature = Nothing
      Set oSI = oSO.ItemObject
      Set oSO = Nothing
      sItemName = oSI.Name
      Set oSI = Nothing
      
      Select Case sItemName
         ' Add any other possible cases as needed
         Case "WebCut_PC_SeamAngleCase1", _
              "WebCut_PC_SeamAngleCase2", _
              "WebCut_PC_SeamAngleCase3", _
              "WebCut_PC_SeamAngleCase4", _
              "WebCut_PC_SeamAngleCase5"
         
            ' Note: Xid value can be found in symbol and end cut rule.
            Dim oStructPort As IJStructPort
            Dim nXid As Long
            
            Set oStructPort = oPCWrapper.Port1
            nXid = oStructPort.SectionID
            Set oStructPort = Nothing
            
            If nXid = 8193 Then
               ' At web
               eFeatureType = SF_WebCut
            ElseIf nXid = 8194 Then
               ' At flange
               eFeatureType = SF_FlangeCut
            End If
         Case Else
            ' Suppose this end cut has one physical connection
            ' Use its type to get thickness
      End Select
      
      
      GetEndCutPCPartThickness _
                         oPCWrapper.ConnectedObject1, _
                         eFeatureType, _
                         dObject1Thickness
                    
      GetEndCutPCPartThickness _
                         oPCWrapper.ConnectedObject2, _
                         eFeatureType, _
                         dObject2Thickness
   End If
   Set oPCWrapper = Nothing
   
   Exit Sub

ErrorHandler:
   Err.Raise LogError(Err, MODULE, "GetPhysicalConnPartsThickness").Number

End Sub

Private Sub GetEndCutPCPartThickness( _
         ByVal oPart As Object, _
         ByVal eEndCutType As StructFeatureTypes, _
         ByRef dPartThickness)
   On Error GoTo ErrorHandler
   
   dPartThickness = 0
   If TypeOf oPart Is IJPlate Then
      Dim oPlate As IJPlate
      
      Set oPlate = oPart
      dPartThickness = oPlate.thickness
      Set oPlate = Nothing
      
   ElseIf TypeOf oPart Is IJStiffenerPart Then
      Dim oSDProfileWrapper As New StructDetailObjects.ProfilePart
      Dim oProfilePart As GSCADCreateModifyUtilities.IJProfileAttributes
      
      Set oSDProfileWrapper.object = oPart
      Set oProfilePart = New GSCADCreateModifyUtilities.ProfileUtils
      If eEndCutType = SF_FlangeCut Then
         dPartThickness = oSDProfileWrapper.flangeThickness
      ElseIf eEndCutType = SF_WebCut Then
         If oSDProfileWrapper.sectionType = "HalfR" Or _
            oSDProfileWrapper.sectionType = "R" Then
             dPartThickness = oProfilePart.GetProfileHeight(oPart)
         Else
             dPartThickness = oSDProfileWrapper.webthickness
         End If
      Else
         LogToFile "Unknown end cut type found in GetEndCutPCPartThickness"
      End If
      Set oSDProfileWrapper = Nothing
      Set oProfilePart = Nothing
      
   ElseIf TypeOf oPart Is IJBeamPart Then
      Dim oSDBeamWrapper As New StructDetailObjects.BeamPart
      
      Set oSDBeamWrapper.object = oPart
      If eEndCutType = SF_FlangeCut Then
         dPartThickness = oSDBeamWrapper.flangeThickness
      ElseIf eEndCutType = SF_WebCut Then
         dPartThickness = oSDBeamWrapper.webthickness
      Else
         LogToFile "Unknown end cut type found in GetEndCutPCPartThickness"
      End If
      Set oSDBeamWrapper = Nothing
      
   ElseIf TypeOf oPart Is ISPSMemberPartPrismatic Then
      Dim oSDMemberWrapper As New StructDetailObjects.MemberPart
      
      Set oSDMemberWrapper.object = oPart
      If eEndCutType = SF_FlangeCut Then
         dPartThickness = oSDMemberWrapper.flangeThickness
      ElseIf eEndCutType = SF_WebCut Then
         dPartThickness = oSDMemberWrapper.webthickness
      Else
         LogToFile "Unknown end cut type found in GetEndCutPCPartThickness"
      End If
      Set oSDMemberWrapper = Nothing
   Else
      LogToFile "Unknown part type found in GetEndCutPCPartThickness"
   End If
   
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "GetEndCutPCPartThickness").Number
   
End Sub

Private Sub LogToFile(strMsg As String)
   Dim nFileNumber As Integer
   Dim strFileName As String
   Dim strRow As String
   
   Dim bLog As Boolean
   
   bLog = False
   If bLog = True Then
      nFileNumber = FreeFile
      strFileName = "c:\temp\PhysicalConnRule.log"
      Open strFileName For Append As nFileNumber
      
      Print #nFileNumber, Now() & ": " & strMsg
      Close nFileNumber
   End If
End Sub

' GetChamferedSidesInfo translates the input chamfer string into an enum value indicating which sides of the
' overall connection are chamfered.  The main purpose for this translation is to properly flip the sides of
' chamfers on the non-reference part if the molded surface normals are not in the same direction
' This is required because each chamfer is placed/named based on the molded normal of the part on which it is
' placed. but the parameter rules will be processed using offsets defined relative to the normal of the
' reference part only
Public Function GetChamferedSidesInfo(ByRef oRefPlate As Object, _
                                      ByRef oNRPlate As Object, _
                                      ByRef sChamferType As String, _
                                      iRefPartNum As Integer) As pcrChamferedSidesInfo
                                      
    On Error GoTo ErrorHandler
    
    If sChamferType = "None" Then
        GetChamferedSidesInfo = pcr_NoChamfer
        
    ElseIf sChamferType = "Obj1Offset" Or _
            sChamferType = "Obj2Offset" Or _
            sChamferType = "Obj1Base" Or _
            sChamferType = "Obj2Base" Then
            
        ' We have chamfer on one side, check to see which side it is.  If it is on the reference part, it is the
        ' same side indicated by the chamfer type.  If it is on the non-reference part, it is only on the side
        ' indicated by the chamfer type if the two surface normals agree.  Otherwise, it is on the other side.
        Dim bFlipSides As Boolean
        If (iRefPartNum = 1 And (sChamferType = "Obj1Offset" Or sChamferType = "Obj1Base")) Or _
           (iRefPartNum = 2 And (sChamferType = "Obj2Offset" Or sChamferType = "Obj2Base")) Then
            ' the chamfer is on the reference part - the side does not get flipped
            bFlipSides = False
        Else
            ' the chamfer is on the NR part, must check for whether the two molded surfaces have consistent normals
            Dim oRefPlatePart As New StructDetailObjects.PlatePart
            Dim bSameDir As Boolean
            bSameDir = True   ' default in case of failure
            If Not oRefPlate Is Nothing Then
                If TypeOf oRefPlate Is IJPlate Then
                    Set oRefPlatePart.object = oRefPlate
                    bSameDir = oRefPlatePart.CompareNormalOfMoldedSurfaces(oNRPlate)
                End If
            End If
            
            If (bSameDir) Then
                bFlipSides = False
            Else
                bFlipSides = True
            End If
        End If
        
        If (bFlipSides = False And (sChamferType = "Obj1Offset" Or sChamferType = "Obj2Offset")) Or _
           (bFlipSides = True And (sChamferType = "Obj1Base" Or sChamferType = "Obj2Base")) Then
            GetChamferedSidesInfo = pcr_OffsetChamfer
        Else
            GetChamferedSidesInfo = pcr_BaseChamfer
        End If
    
    Else
        ' all of the other types indicate a chamfer on both sides
        GetChamferedSidesInfo = pcr_BothChamfer
    End If

Exit Function
                                  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetChamferedSidesInfo").Number
End Function

' GetButtWeldOverlappingThicknessAndAdditions computes the overlapping thickness of two parts at a
' physical connection.  If the parts do not fully align and are not chamfered, it returns the step
' changes in thickness between the parts as additions that have to be applied to one or the other
' part to get from the common thickness to the given side of the part.  If no addition is required or
' this is the smaller part on this side, the value will be zero.  The values will never be negative.
' The values will be zero on any side of the plate where a chamfer has been added to remove the step
' change.
'
' in addition to the above values, this subroutine returns a flag that is true if the two parts are plates
' have opposite normals on the molded surface.  Otherwise, this flag will be false.
Public Sub GetButtWeldOverlappingThicknessAndAdditions(pPRL As IJDParameterLogic, _
                                                       ByRef oPartRef As Object, _
                                                       ByRef oPartNR As Object, _
                                                       ByRef sChamferType As String, _
                                                       ByVal iRefPartNum As Integer, _
                                                       ByRef dThickness As Double, _
                                                       ByRef dAdditionBaseRef As Double, _
                                                       ByRef dAdditionOffsetRef As Double, _
                                                       ByRef dAdditionBaseNR As Double, _
                                                       ByRef dAdditionOffsetNR As Double, _
                                                       ByRef bNormalsReversed As Boolean)
    
    On Error GoTo ErrorHandler
    
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Dim oPlateRef As New StructDetailObjects.PlatePart
    Dim oPlateNR As New StructDetailObjects.PlatePart
    Dim bSameDir As Boolean
    
    Set oPhysConn.object = pPRL.SmartOccurrence
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    ' initialize the normal directions to the same direction
    bNormalsReversed = False
    
    ' get the offsets for reference and non-reference plate
    ' if not plates, the offsets will be 0
    If TypeOf oPhysConn.ConnectedObject1 Is IJPlate And _
       TypeOf oPhysConn.ConnectedObject2 Is IJPlate Then
       
        ' these are plates
        Set oPlateRef.object = oPartRef
        Set oPlateNR.object = oPartNR
       
        ' determine if the plate normals are in the same direction
        bSameDir = oPlateRef.CompareNormalOfMoldedSurfaces(oPlateNR.object)
        If bSameDir = False Then
            bNormalsReversed = True
        End If
        
        ' get the offset values -- the offsets for both the Reference part and the
        ' NR part will be measured using the position and normal of the reference
        ' part molded surface
        Dim offsetToBaseRef As Double
        Dim offsetToBaseNR As Double
        Dim offsetToOffsetRef As Double
        Dim offsetToOffsetNR As Double
        
        ' get the delta values -- these are used to set the addition values
        Dim dBaseDelta As Double
        Dim dOffsetDelta As Double
        
        ' get the offset values, relative to the normal of the reference plate
        offsetToBaseRef = oPlateRef.OffsetToBaseFace
        offsetToOffsetRef = oPlateRef.OffsetToOffsetFace
        offsetToBaseNR = oPlateNR.OffsetToBaseFace(oPlateRef)
        offsetToOffsetNR = oPlateNR.OffsetToOffsetFace(oPlateRef)
        
        ' *******************************************************************************
        ' Compute the deltas to find the difference in thickness between the two parts at
        ' a particular face.  The delta is always from the reference to the non-reference
        ' part and the choice of base/offset is based on the normal of the reference part
        ' molded surface.
        ' The deltas are set to zero for any side that has a chamfer because the chamfer
        ' trims the thicker part to the leve of the thinner one and there is no step
        ' *******************************************************************************
        Dim chamferSides As pcrChamferedSidesInfo
        chamferSides = GetChamferedSidesInfo(oPartRef, oPartNR, sChamferType, iRefPartNum)
        
        ' first the base delta for anything without a base chamfer
        ' this is the delta from the reference to the non-reference part
        If chamferSides = pcr_NoChamfer Or chamferSides = pcr_OffsetChamfer Then
            If bNormalsReversed = False Then
                dBaseDelta = offsetToBaseNR - offsetToBaseRef
            Else
                dBaseDelta = 0
            End If
            
        Else
            ' the part has a chamfer on the base face or both faces, the effective base face offset is 0
            dBaseDelta = 0
        End If
        
        ' now the offset delta for anything without an offset chamfer
        ' this is the delta from the reference to the non-reference part
        If chamferSides = pcr_NoChamfer Or chamferSides = pcr_BaseChamfer Then
            If bNormalsReversed = False Then
                dOffsetDelta = offsetToOffsetNR - offsetToOffsetRef
            Else
                dOffsetDelta = 0
            End If
        Else
            ' the part has a chamfer on the offset face or both faces, the effective offset face offset is 0
            dOffsetDelta = 0
        End If
        
        ' ***********************************************************************************
        ' only set the values if they represent an addition.  Note that, because the sign of
        ' the offset is based on the single sheet normal, the sign of the offset that defines
        ' an addition is switched between BASE and OFFSET
        ' ***********************************************************************************
        If dBaseDelta > 0 Then
            ' the reference part extends past the NR part on the base side
            dAdditionBaseRef = dBaseDelta
        Else
            ' the value is 0 or the NR part extends past the reference part on the base side
            dAdditionBaseNR = Abs(dBaseDelta)
        End If
        If dOffsetDelta > 0 Then
            ' the NR part extends past the reference part on the offset side
            dAdditionOffsetNR = dOffsetDelta
        Else
            ' the value is 0 or the reference part extends past the NR part on the offset side
            dAdditionOffsetRef = Abs(dOffsetDelta)
        End If
        
        'TR 171555 -- This special case is handled to get the offset between the molded surfaces of
        'the two connected parts as the wrappers OffsetToBaseFace and OffsetToOffsetFace fail to get the
        'MoldedOffset when a CAN is connected to tube.
        
        If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then ' ref part is part1
            Set oPort1 = oPhysConn.Port1
            Set oPort2 = oPhysConn.Port2
        Else ' ref part is part2
            Set oPort1 = oPhysConn.Port2
            Set oPort2 = oPhysConn.Port1
        End If
        'ensured that ref part and non ref part are passed.
        Dim bIsConnBtnTubeAndCan As Boolean
        Dim dMoldedOffset As Double
        
        bIsConnBtnTubeAndCan = IsConnectionBetweenTubeAndCAN(oPort1, oPort2, dMoldedOffset)

        If bIsConnBtnTubeAndCan Then
           offsetToBaseNR = dMoldedOffset + offsetToBaseNR
           offsetToOffsetNR = dMoldedOffset + offsetToOffsetNR
        End If
        ' ****************************************************************************
        ' find the overlapping thickness
        '
        ' because the offsets are all measured relative to the molded surface normal
        ' we know that the BASE values are less than the OFFSET values.  Therefore,
        ' to get the overlap, find the difference between the largest BASE offset and
        ' the smallest OFFSET value
        ' ****************************************************************************
        Dim dBigBaseValue As Double
        Dim dSmallOffsetValue As Double
        If offsetToBaseRef > offsetToBaseNR Then
            dBigBaseValue = offsetToBaseRef
        Else
            dBigBaseValue = offsetToBaseNR
        End If
            
        If offsetToOffsetRef < offsetToOffsetNR Then
            dSmallOffsetValue = offsetToOffsetRef
        Else
            dSmallOffsetValue = offsetToOffsetNR
        End If
            
        dThickness = dSmallOffsetValue - dBigBaseValue
     
    Else
        ' this is not two plate parts, set the thickness to the reference part thickness
        ' put the addition on the non-reference offset side (not correct, but we have no better information)
        Dim oProfilePart As GSCADCreateModifyUtilities.IJProfileAttributes
        
        If TypeOf oPhysConn.ConnectedObject1 Is IJStiffener And _
           TypeOf oPhysConn.ConnectedObject2 Is IJStiffener Then
           
           Set oProfilePart = New GSCADCreateModifyUtilities.ProfileUtils
           Dim oStiffener1 As New StructDetailObjects.ProfilePart
           Dim oStiffener2 As New StructDetailObjects.ProfilePart
           
           Set oStiffener1.object = oPhysConn.ConnectedObject1
           Set oStiffener2.object = oPhysConn.ConnectedObject2
           
           If iRefPartNum = 2 Then
              If (oStiffener1.sectionType = "HalfR" And oStiffener2.sectionType = "HalfR") Then
                   dThickness = oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject2)
                   dAdditionOffsetNR = (oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject1)) - dThickness
              ElseIf (oStiffener1.sectionType = "R" And oStiffener2.sectionType = "R") Then
                   'For Round Cross Sections Bevel thickness is measured from Inside(Centre) to Outside Surface i.e nothing but Radius
                   dThickness = (oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject2)) / 2
                   dAdditionOffsetNR = (oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject1) / 2) - dThickness
              Else
                 dThickness = oPhysConn.Object2Thickness
                 dAdditionOffsetNR = oPhysConn.Object1Thickness - dThickness
              End If
           Else
              If (oStiffener1.sectionType = "HalfR" And oStiffener2.sectionType = "HalfR") Then
                  dThickness = oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject1)
                  dAdditionOffsetNR = (oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject2)) - dThickness
              ElseIf (oStiffener1.sectionType = "R" And oStiffener2.sectionType = "R") Then
                  'For Round Cross Sections Bevel thickness is measured from Inside(Centre) to Outside Surface i.e nothing but Radius
                  dThickness = (oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject1)) / 2
                  dAdditionOffsetNR = (oProfilePart.GetProfileHeight(oPhysConn.ConnectedObject2) / 2) - dThickness
              Else
                 dThickness = oPhysConn.Object1Thickness
                 dAdditionOffsetNR = oPhysConn.Object2Thickness - dThickness
              End If
           End If
           
           Set oStiffener1 = Nothing
           Set oStiffener2 = Nothing
           
        Else
           If iRefPartNum = 2 Then
              dThickness = oPhysConn.Object2Thickness
              dAdditionOffsetNR = oPhysConn.Object1Thickness - dThickness
           Else
              dThickness = oPhysConn.Object1Thickness
              dAdditionOffsetNR = oPhysConn.Object2Thickness - dThickness
           End If
        End If
        Set oProfilePart = Nothing
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetButtWeldOverlappingThicknessAndAdditions").Number
End Sub

Public Sub GetWeldParmInfo( _
               ByVal sWeldName As String, _
               ByRef arrayParmInfo() As PARAMETER_INFO, _
               ByRef nParmCount As Integer)
   On Error GoTo ErrorHandler
   Dim nParmIndex As Integer
   
   Select Case sWeldName
      Case LAP_WELD1, LAP_WELD2
         nParmCount = 18
         ReDim arrayParmInfo(1 To nParmCount)
                  
         arrayParmInfo(1).sName = PRIMARY_SIDE_GROOVE_SIZE
         arrayParmInfo(1).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(1).vValue = 0
         
         arrayParmInfo(2).sName = PRIMARY_SIDE_ACTUAL_THROAT_THICKNESS
         arrayParmInfo(2).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(2).vValue = 0
         
         arrayParmInfo(3).sName = SECONDARY_SIDE_ACTUAL_THROAT_THICKNESS
         arrayParmInfo(3).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(3).vValue = 0
         
         arrayParmInfo(4).sName = PRIMARY_SIDE_NOMINAL_THROAT_THICKNESS
         arrayParmInfo(4).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(4).vValue = 0
         
         arrayParmInfo(5).sName = SECONDARY_SIDE_NOMINAL_THROAT_THICKNESS
         arrayParmInfo(5).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(5).vValue = 0
         
         arrayParmInfo(6).sName = PRIMARY_SIDE_LENGTH
         arrayParmInfo(6).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(6).vValue = 0
         
         arrayParmInfo(7).sName = SECONDARY_SIDE_LENGTH
         arrayParmInfo(7).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(7).vValue = 0
         
         arrayParmInfo(8).sName = PRIMARY_SIDE_PITCH
         arrayParmInfo(8).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(8).vValue = 0
         
         arrayParmInfo(9).sName = SECONDARY_SIDE_PITCH
         arrayParmInfo(9).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(9).vValue = 0
         
         arrayParmInfo(10).sName = PRIMARY_SIDE_GROOVE_ANGLE
         arrayParmInfo(10).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(10).vValue = 0
         
         arrayParmInfo(11).sName = SECONDARY_SIDE_GROOVE_ANGLE
         arrayParmInfo(11).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(11).vValue = 0
         
         arrayParmInfo(12).sName = ALL_AROUND
         arrayParmInfo(12).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(12).vValue = 0
         
         arrayParmInfo(13).sName = TAIL_NOTES
         arrayParmInfo(13).eArgumentType = imsARGUMENT_IS_BSTR
         arrayParmInfo(13).vValue = ""
         
         arrayParmInfo(14).sName = TAIL_NOTE_IS_REFERENCE
         arrayParmInfo(14).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(14).vValue = 0
         
         arrayParmInfo(15).sName = PRIMARY_SIDE_ACTUAL_LEG_LENGTH
         arrayParmInfo(15).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(15).vValue = 0
         
         arrayParmInfo(16).sName = SECONDARY_SIDE_ACTUAL_LEG_LENGTH
         arrayParmInfo(16).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(16).vValue = 0
         
         arrayParmInfo(17).sName = PRIMARY_SIDE_SYMBOL
         arrayParmInfo(17).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(17).vValue = 0
         
         arrayParmInfo(18).sName = SECONDARY_SIDE_SYMBOL
         arrayParmInfo(18).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(18).vValue = 0
         
      Case CHAIN_WELD, FILLET_WELD1, FILLET_WELD2, _
           TEE_WELD_CHILL, TEE_WELD_K, TEE_WELD_V, TEE_WELD_X, TEE_WELD_Y, ZIG_ZAG_WELD, _
           TEE_WELD_SQUARE, STAGGERED_WELD
         If sWeldName = TEE_WELD_CHILL Then
            nParmCount = 22
         Else
            nParmCount = 21
         End If
         
         ReDim arrayParmInfo(1 To nParmCount)
                  
         arrayParmInfo(1).sName = PRIMARY_SIDE_GROOVE
         arrayParmInfo(1).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(1).vValue = 0
         
         arrayParmInfo(2).sName = SECONDARY_SIDE_GROOVE
         arrayParmInfo(2).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(2).vValue = 0
         
         arrayParmInfo(3).sName = PRIMARY_SIDE_GROOVE_SIZE
         arrayParmInfo(3).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(3).vValue = 0
         
         arrayParmInfo(4).sName = SECONDARY_SIDE_GROOVE_SIZE
         arrayParmInfo(4).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(4).vValue = 0
         
         arrayParmInfo(5).sName = PRIMARY_SIDE_ACTUAL_THROAT_THICKNESS
         arrayParmInfo(5).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(5).vValue = 0
         
         arrayParmInfo(6).sName = SECONDARY_SIDE_ACTUAL_THROAT_THICKNESS
         arrayParmInfo(6).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(6).vValue = 0
         
         arrayParmInfo(7).sName = PRIMARY_SIDE_NOMINAL_THROAT_THICKNESS
         arrayParmInfo(7).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(7).vValue = 0
         
         arrayParmInfo(8).sName = SECONDARY_SIDE_NOMINAL_THROAT_THICKNESS
         arrayParmInfo(8).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(8).vValue = 0
         
         arrayParmInfo(9).sName = PRIMARY_SIDE_LENGTH
         arrayParmInfo(9).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(9).vValue = 0
         
         arrayParmInfo(10).sName = SECONDARY_SIDE_LENGTH
         arrayParmInfo(10).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(10).vValue = 0
         
         arrayParmInfo(11).sName = PRIMARY_SIDE_PITCH
         arrayParmInfo(11).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(11).vValue = 0
         
         arrayParmInfo(12).sName = SECONDARY_SIDE_PITCH
         arrayParmInfo(12).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(12).vValue = 0
         
         arrayParmInfo(13).sName = PRIMARY_SIDE_ROOT_OPENING
         arrayParmInfo(13).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(13).vValue = 0
         
         arrayParmInfo(14).sName = PRIMARY_SIDE_GROOVE_ANGLE
         arrayParmInfo(14).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(14).vValue = 0
         
         arrayParmInfo(15).sName = SECONDARY_SIDE_GROOVE_ANGLE
         arrayParmInfo(15).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(15).vValue = 0
         
         arrayParmInfo(16).sName = TAIL_NOTES
         arrayParmInfo(16).eArgumentType = imsARGUMENT_IS_BSTR
         arrayParmInfo(16).vValue = ""
         
         arrayParmInfo(17).sName = TAIL_NOTE_IS_REFERENCE
         arrayParmInfo(17).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(17).vValue = 0
      
         arrayParmInfo(18).sName = PRIMARY_SIDE_ACTUAL_LEG_LENGTH
         arrayParmInfo(18).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(18).vValue = 0
         
         arrayParmInfo(19).sName = SECONDARY_SIDE_ACTUAL_LEG_LENGTH
         arrayParmInfo(19).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(19).vValue = 0
         nParmIndex = 19
         
         If sWeldName = TEE_WELD_CHILL Then
            arrayParmInfo(20).sName = SECONDARY_SIDE_ROOT_OPENING
            arrayParmInfo(20).eArgumentType = imsARGUMENT_IS_DOUBLE
            arrayParmInfo(20).vValue = 0
            nParmIndex = 20
         End If
         
         nParmIndex = nParmIndex + 1
         arrayParmInfo(nParmIndex).sName = PRIMARY_SIDE_SYMBOL
         arrayParmInfo(nParmIndex).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(nParmIndex).vValue = 0

         nParmIndex = nParmIndex + 1
         arrayParmInfo(nParmIndex).sName = SECONDARY_SIDE_SYMBOL
         arrayParmInfo(nParmIndex).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(nParmIndex).vValue = 0
      
      Case BUTT_WELD_I, BUTT_WELD_IV, BUTT_WELD_IX, BUTT_WELD_K, _
           BUTT_WELD_V, BUTT_WELD_X, BUTT_WELD_Y

         nParmCount = 12
         ReDim arrayParmInfo(1 To nParmCount)
                  
         arrayParmInfo(1).sName = PRIMARY_SIDE_GROOVE
         arrayParmInfo(1).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(1).vValue = 0
         
         arrayParmInfo(2).sName = SECONDARY_SIDE_GROOVE
         arrayParmInfo(2).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(2).vValue = 0
         
         arrayParmInfo(3).sName = PRIMARY_SIDE_GROOVE_SIZE
         arrayParmInfo(3).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(3).vValue = 0
         
         arrayParmInfo(4).sName = SECONDARY_SIDE_GROOVE_SIZE
         arrayParmInfo(4).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(4).vValue = 0
         
         arrayParmInfo(5).sName = PRIMARY_SIDE_ACTUAL_THROAT_THICKNESS
         arrayParmInfo(5).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(5).vValue = 0
         
         arrayParmInfo(6).sName = SECONDARY_SIDE_ACTUAL_THROAT_THICKNESS
         arrayParmInfo(6).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(6).vValue = 0
         
         arrayParmInfo(7).sName = PRIMARY_SIDE_ROOT_OPENING
         arrayParmInfo(7).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(7).vValue = 0
         
         arrayParmInfo(8).sName = SECONDARY_SIDE_ROOT_OPENING
         arrayParmInfo(8).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(8).vValue = 0
         
         arrayParmInfo(9).sName = PRIMARY_SIDE_GROOVE_ANGLE
         arrayParmInfo(9).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(9).vValue = 0
         
         arrayParmInfo(10).sName = SECONDARY_SIDE_GROOVE_ANGLE
         arrayParmInfo(10).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(10).vValue = 0
         
         arrayParmInfo(11).sName = TAIL_NOTES
         arrayParmInfo(11).eArgumentType = imsARGUMENT_IS_BSTR
         arrayParmInfo(11).vValue = ""
         
         arrayParmInfo(12).sName = TAIL_NOTE_IS_REFERENCE
         arrayParmInfo(12).eArgumentType = imsARGUMENT_IS_DOUBLE
         arrayParmInfo(12).vValue = 0
      
      Case Else
      
   End Select
   
   Exit Sub
ErrorHandler:
   Err.Raise Err.Number, "SMPhysConnRules::GetWeldParmInfo"
   
End Sub
'
' Based on weld name, get weld parameter info, and add them as parameter rule output
'
Public Sub AddWeldParmRuleOutputs( _
                    ByVal sWeldName, _
                    ByVal oOH As IJDOutputsHelper)
   On Error GoTo ErrorHandler
   
   Dim arrayParmInfo() As PARAMETER_INFO
   Dim nParmCount As Integer
   Dim nParmIndex As Integer
    
   GetWeldParmInfo sWeldName, arrayParmInfo, nParmCount
   For nParmIndex = 1 To nParmCount
      oOH.SetOutput arrayParmInfo(nParmIndex).sName, _
                    arrayParmInfo(nParmIndex).eArgumentType
   Next
   
   Exit Sub
   
ErrorHandler:
   Err.Raise Err.Number, "SMPhysConnRules::AddWeldParmRuleOutputs"
End Sub
'
'  Set weld parameter values
'
Public Sub SetWeldParmValues( _
            ByVal oPRL As IJDParameterLogic, _
            ByVal nParmCount As Integer, _
            ByRef arrayParmInfo() As PARAMETER_INFO)
   On Error GoTo ErrorHandler
    
   Dim nParmIndex As Integer
    
   For nParmIndex = 1 To nParmCount
      oPRL.Add arrayParmInfo(nParmIndex).sName, _
               arrayParmInfo(nParmIndex).vValue
   Next
    
   Exit Sub
    
ErrorHandler:
   Err.Raise Err.Number, "SMPhysConnRules::SetWeldParmValues"
End Sub

'
'   Given a symbol and array of input names,
'     - Check if these inputs are overridden and get overridden values
'     -  If the caller wants to get a constant sized array for different symbols that
'        share some, but not, values, it is possible to pass in a name of "ParmUnsupported"
'        for array values that are kept as fillers.  These parameters will always return a
'        flag indicating that they have not been overridden.
'
Public Sub GetSymbolParameterInfo( _
                ByVal oSymbol As IJDSymbol, _
                ByVal nParmCount As Long, _
                ByRef arrayParmInfo() As PARAMETER_INFO)
   On Error GoTo ErrorHandler
                
   If oSymbol Is Nothing Then
      Exit Sub
   End If
   
   If Not (UBound(arrayParmInfo) - LBound(arrayParmInfo) + 1) = nParmCount Then
      Exit Sub
   End If
   
   Dim oParmRuleSymDef As IJDSymbolDefinition
   Dim oSO As IJSmartOccurrence
   Dim oSI As IJSmartItem
   Dim oOutputCtl As IJOutputControl
   
   Set oOutputCtl = oSymbol
   Set oSO = oSymbol
   Set oSI = oSO.ItemObject
   Set oParmRuleSymDef = oSI.ParameterRuleDef
   
   Dim nParmIndex As Long
   Dim nOverriddenCount As Long
   Dim bEditable As Boolean
   
   nOverriddenCount = 0
   For nParmIndex = 1 To nParmCount
   
      ' initialize the override flag to false.  It will be set true only for valid parameters
      ' that have been overridden
      arrayParmInfo(nParmIndex).bOverridden = False
      If arrayParmInfo(nParmIndex).sName <> "ParmUnsupported" Then
         bEditable = oOutputCtl.OutputEditable(oParmRuleSymDef, arrayParmInfo(nParmIndex).sName)
         If bEditable = True Then
            ' Overridden - mark this array value as overidden -- will retrieve the value later
            ' keep the count so we can end early if none are overridden
            arrayParmInfo(nParmIndex).bOverridden = True
            nOverriddenCount = nOverriddenCount + 1
         End If
      End If
   Next
   
   If nOverriddenCount = 0 Then
      Exit Sub
   End If
   
        Dim nSetCount As Long
    
   ' Get the part symbol definition helepr
   Dim oSymbolDefHelper As IJDSymbolDefHelper
   Set oSymbolDefHelper = oSI
     
   If Not oSymbolDefHelper Is Nothing Then
       Dim parametersAndProperties() As Variant
       oSymbolDefHelper.GetParametersAndProperties parametersAndProperties()
       
       Dim n0LBound As Long
       Dim n0UBound As Long
       Dim n1LBound As Long
           
       n0LBound = LBound(parametersAndProperties, 1)
       n0UBound = UBound(parametersAndProperties, 1)
       n1LBound = LBound(parametersAndProperties, 2)
       
       'iterate through the symbol input safearray values
       Dim iIndex As Long
       For iIndex = n0LBound To n0UBound
           Dim varSymbolParameter As Variant

           varSymbolParameter = parametersAndProperties(iIndex, n1LBound + 1)
           Dim sSymbolParameter As String
           sSymbolParameter = varSymbolParameter

           Dim oAttributes As IJDAttributes
           Dim vAttributeValue As Variant
           Dim varInterfaceType As Variant
           varInterfaceType = parametersAndProperties(iIndex, n1LBound + 7)
           
           Dim varPartOccIID As Variant
           If varInterfaceType = 2 Then ' Occurrence Attribute
               ' Get the interface for the parameter
               varPartOccIID = parametersAndProperties(iIndex, n1LBound + 11)

               ' Get the name of the part occurrence attribute
               Dim varPartOccAttribute As Variant
               varPartOccAttribute = parametersAndProperties(iIndex, n1LBound + 5)

               Dim oAttributesCol As IJDAttributesCol
               Set oAttributes = oSymbol
               Set oAttributesCol = oAttributes.CollectionOfAttributes(varPartOccIID)
               If Not oAttributesCol Is Nothing Then
                   ' Get the specific occurrence attribute base d on the attribute name
                   Dim oPartOccAttribute As IJDAttribute
                   Set oPartOccAttribute = oAttributesCol.Item(varPartOccAttribute)

                   vAttributeValue = oPartOccAttribute.Value
               End If
           ElseIf varInterfaceType = 1 Then ' Definition Attribute
               ' Get the interface for the parameter
               varPartOccIID = parametersAndProperties(iIndex, n1LBound + 9)

               ' Get the name of the part attribute
               Dim varPartAttribute As Variant
               varPartAttribute = parametersAndProperties(iIndex, n1LBound + 3)

               Set oAttributes = oSI
               Set oAttributesCol = oAttributes.CollectionOfAttributes(varPartOccIID)
               If Not oAttributesCol Is Nothing Then
                   ' Get the specific occurrence attribute base d on the attribute name
                   Dim oPartAttribute As IJDAttribute
                   Set oPartAttribute = oAttributesCol.Item(varPartAttribute)

                   vAttributeValue = oPartAttribute.Value
               End If
           End If
           
           ' check if the symbol parameter names are the same
           For nParmIndex = 1 To nParmCount
               If sSymbolParameter = arrayParmInfo(nParmIndex).sName And _
                  arrayParmInfo(nParmIndex).bOverridden = True Then
                   arrayParmInfo(nParmIndex).vValue = vAttributeValue
                   nSetCount = nSetCount + 1
                   Exit For
               End If
           Next
       Next iIndex
   End If
   
   Exit Sub
   
ErrorHandler:
   Err.Raise Err.Number, "SMPhysConnRules::GetSymbolParameterInfo"
End Sub


' Set the calculated weld parameter info for TEE weld types
'
' The decisions can be made based on the specific weld type passed in, but all TEE weld types
' can use the same logic, as defined here.  Customers may want to modify this behavior for
' specific classes of TEE weld.
'
'This procedure will use the values provided by the user if they have been overriden.
'
' The input values, except for oPRL, sWeldName, and GrooveType correspond to the actual values stored by
' the parm rule, not the intermediate name used within the parm rule.  For example, some rules may carry
' a value as "refSide", but then store it as "antiRefSide" based on some other flag.  We need the value
' actually stored.
'
' For each of these input values, we will determine if the value was overriden by the user and use the
' value from the parameter grid if it has been overridden
'
' bRefIsMolded refers to whether the ref side of the reference part is molded
' this can be overridden if the user has overridden the "ReferenceSide" parameter
'
Public Sub SetCalculatedTeeWeldParams( _
            ByVal pPRL As IJDParameterLogic, _
            ByVal sWeldName As String, _
            ByVal bRefIsMolded As Boolean, _
            ByVal dRefSideFirstBevelDepth As Double, _
            ByVal dAntiRefSideFirstBevelDepth As Double, _
            ByVal dRefSideFirstBevelAngle As Double, _
            ByVal dAntiRefSideFirstBevelAngle As Double, _
            ByVal GrooveType As Long, _
            ByVal dMoldedFillet As Double, _
            ByVal dAntiMoldedFillet As Double, _
            ByVal dRootGap As Double, _
            ByVal dLength As Double, _
            ByVal dPitch As Double)

  On Error GoTo ErrorHandler
  
  ' check all symbol parameter values used in setting weld parameter values to see which are overridden
  ' there are a total of 9 values that can be overridden, but not all are valid for all weld types
  ' Create the array with 9 values and initialize them to "ParmUnsupported".  We will set the valid ones
  ' based on the cross section.  The following is the full set of values and proper alignment
  '
  ' arrayBevelParmInfo(1).sName = "MoldedFillet"
  ' arrayBevelParmInfo(2).sName = "AntiMoldedFillet"
  ' arrayBevelParmInfo(3).sName = "Pitch"
  ' arrayBevelParmInfo(4).sName = "Length"
  ' arrayBevelParmInfo(5).sName = "RefSideFirstBevelDepth"
  ' arrayBevelParmInfo(6).sName = "RefSideFirstBevelAngle"
  ' arrayBevelParmInfo(7).sName = "AntiRefSideFirstBevelDepth"
  ' arrayBevelParmInfo(8).sName = "AntiRefSideFirstBevelAngle"
  ' arrayBevelParmInfo(9).sName = "RootGap"
  ' arrayBevelParmInfo(10).sName = "ReferenceSide"
  ' arrayBevelParmInfo(11).sName = "FilletMeasureMethod"
  
  ' local variable to handle possible override of reference side
  Dim bRefIsMoldedLocal As Boolean
  bRefIsMoldedLocal = bRefIsMolded
  
  ' initialize
  Dim arrayBevelParmInfo(1 To 11) As PARAMETER_INFO
  Dim index As Long
  For index = 1 To 11
     arrayBevelParmInfo(index).sName = "ParmUnsupported"
  Next index

  Select Case sWeldName
     
     Case CHAIN_WELD, STAGGERED_WELD, ZIG_ZAG_WELD
        arrayBevelParmInfo(1).sName = "MoldedFillet"
        arrayBevelParmInfo(2).sName = "AntiMoldedFillet"
        arrayBevelParmInfo(3).sName = "Pitch"
        arrayBevelParmInfo(4).sName = "Length"
        arrayBevelParmInfo(10).sName = "ReferenceSide"
        arrayBevelParmInfo(11).sName = "FilletMeasureMethod"
        
     Case FILLET_WELD1, FILLET_WELD2
        arrayBevelParmInfo(1).sName = "MoldedFillet"
        arrayBevelParmInfo(2).sName = "AntiMoldedFillet"
        arrayBevelParmInfo(10).sName = "ReferenceSide"
        arrayBevelParmInfo(11).sName = "FilletMeasureMethod"
        
     Case TEE_WELD_CHILL, TEE_WELD_K, TEE_WELD_V, TEE_WELD_X, TEE_WELD_Y
        arrayBevelParmInfo(1).sName = "MoldedFillet"
        arrayBevelParmInfo(2).sName = "AntiMoldedFillet"
        arrayBevelParmInfo(5).sName = "RefSideFirstBevelDepth"
        arrayBevelParmInfo(6).sName = "RefSideFirstBevelAngle"
        arrayBevelParmInfo(7).sName = "AntiRefSideFirstBevelDepth"
        arrayBevelParmInfo(8).sName = "AntiRefSideFirstBevelAngle"
        arrayBevelParmInfo(10).sName = "ReferenceSide"
        If sWeldName = TEE_WELD_CHILL Then
           arrayBevelParmInfo(9).sName = "RootGap"
        End If
        arrayBevelParmInfo(11).sName = "FilletMeasureMethod"

     Case TEE_WELD_SQUARE
        arrayBevelParmInfo(1).sName = "MoldedFillet"
        arrayBevelParmInfo(2).sName = "AntiMoldedFillet"
        arrayBevelParmInfo(5).sName = "RefSideFirstBevelDepth"
        arrayBevelParmInfo(6).sName = "RefSideFirstBevelAngle"
        arrayBevelParmInfo(7).sName = "AntiRefSideFirstBevelDepth"
        arrayBevelParmInfo(8).sName = "AntiRefSideFirstBevelAngle"
        arrayBevelParmInfo(10).sName = "ReferenceSide"
        arrayBevelParmInfo(11).sName = "FilletMeasureMethod"

  End Select
    
  ' get any overridden values
  GetSymbolParameterInfo pPRL.SmartOccurrence, 11, arrayBevelParmInfo
  
  ' update any overridden values
  If arrayBevelParmInfo(1).bOverridden = True Then
     dMoldedFillet = arrayBevelParmInfo(1).vValue
  End If
    
  If arrayBevelParmInfo(2).bOverridden = True Then
     dAntiMoldedFillet = arrayBevelParmInfo(2).vValue
  End If
    
  ' Calculate leg value if Fillets are measured as throat
  If arrayBevelParmInfo(11).vValue = 65537 Then
     dMoldedFillet = dMoldedFillet * Sqr(2)
     dAntiMoldedFillet = dAntiMoldedFillet * Sqr(2)
  End If
  
  If arrayBevelParmInfo(3).bOverridden = True Then
     dPitch = arrayBevelParmInfo(3).vValue
  End If
    
  If arrayBevelParmInfo(4).bOverridden = True Then
     dLength = arrayBevelParmInfo(4).vValue
  End If
    
  If arrayBevelParmInfo(5).bOverridden = True Then
     dRefSideFirstBevelDepth = arrayBevelParmInfo(5).vValue
  End If
    
  If arrayBevelParmInfo(6).bOverridden = True Then
     dRefSideFirstBevelAngle = arrayBevelParmInfo(6).vValue
  End If
    
  If arrayBevelParmInfo(7).bOverridden = True Then
     dAntiRefSideFirstBevelDepth = arrayBevelParmInfo(7).vValue
  End If
    
  If arrayBevelParmInfo(8).bOverridden = True Then
     dAntiRefSideFirstBevelAngle = arrayBevelParmInfo(8).vValue
  End If
    
  If arrayBevelParmInfo(9).bOverridden = True Then
     dRootGap = arrayBevelParmInfo(9).vValue
  End If
    
  If arrayBevelParmInfo(10).bOverridden = True Then
     If LCase(arrayBevelParmInfo(10).vValue) = "molded" Then
        bRefIsMoldedLocal = True
     Else
        bRefIsMoldedLocal = False
     End If
  End If
  
  ' set up some intermediate values used for the outputs
  Dim RefSideGroove As Long
  Dim AntiRefSideGroove As Long
  If (dRefSideFirstBevelDepth < 0.0001) Then
      ' there is no primary bevel, use the primary fillet and force bevel angle to 0 and groove type to none
      dRefSideFirstBevelAngle = 0#
      RefSideGroove = pcr_WG_None
  Else
      ' there is a primary bevel, use input angle and groove type
      RefSideGroove = GrooveType
  End If
  
  
  If (dAntiRefSideFirstBevelDepth < 0.0001) Then
      ' there is no secondary bevel, use the secondary fillet and force bevel angle to 0 and groove type to none
      dAntiRefSideFirstBevelAngle = 0#
      AntiRefSideGroove = pcr_WG_None
  Else
      ' there is a secondary bevel, use input angle and groove type
      AntiRefSideGroove = GrooveType
  End If
  
  ' now set the output values
  ' the proper values to use are as follows:
  ' NOTE: Ref/AntiRef side values are reversed from this if the ref side is not molded
  ' "PrimarySideLength" = dLength
  ' "PrimarySidePitch" = dPitch
  ' "SecondarySideLength" = dLength
  ' "SecondarySidePitch" = dPitch
  ' "PrimarySideGroove" = RefSideGroove
  ' "SecondarySideGroove" = AntiRefSideGroove
  ' "PrimarySideActualThroatThickness" = dRefSideFirstBevelDepth
  ' "PrimarySideNominalThroatThickness" = dMoldedFillet
  ' "PrimarySideRootOpening" = dRootGap
  ' "SecondarySideRootOpening" = dRootGap
  ' "PrimarySideGrooveAngle" = dRefSideFirstBevelAngle
  ' "SecondarySideActualThroatThickness" = dAntiRefSideFirstBevelDepth
  ' "SecondarySideNominalThroatThickness" = dAntiMoldedFillet
  ' "SecondarySideGrooveAngle" = dAntiRefSideFirstBevelAngle
  
  ' however, we need to write out all of these computed parameters into the parameter info
  ' structure by getting all of the possible computed values, updating the ones that are
  ' actually updated by this rule, and writing all of them out.  Otherwise, we would get
  ' errors caused by the ones that are declared as outputs, but that we are not updating
   
  Dim arrayIJWeldingSymbolParmInfo() As PARAMETER_INFO
  Dim nParmCount As Integer
  Dim nParmIndex As Integer
   
  ' get all of the IJWeldingSymbol parameters
  GetWeldParmInfo sWeldName, arrayIJWeldingSymbolParmInfo, nParmCount
  For nParmIndex = 1 To nParmCount
     Select Case arrayIJWeldingSymbolParmInfo(nParmIndex).sName
     
        Case PRIMARY_SIDE_LENGTH
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dLength
           
        Case PRIMARY_SIDE_PITCH
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dPitch
           
        Case SECONDARY_SIDE_LENGTH
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dLength
           
        Case SECONDARY_SIDE_PITCH
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dPitch
           
        Case PRIMARY_SIDE_GROOVE
           If bRefIsMoldedLocal Then
              ' ref side is molded, ref side is primary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGroove
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGroove
           End If
           
        Case PRIMARY_SIDE_ACTUAL_THROAT_THICKNESS
           If bRefIsMoldedLocal Then
              ' ref side is molded, ref side is primary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dRefSideFirstBevelDepth
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dAntiRefSideFirstBevelDepth
           End If
           
        Case PRIMARY_SIDE_NOMINAL_THROAT_THICKNESS
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dMoldedFillet
           
        Case PRIMARY_SIDE_ROOT_OPENING, SECONDARY_SIDE_ROOT_OPENING
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dRootGap
           
        Case PRIMARY_SIDE_GROOVE_ANGLE
           If bRefIsMoldedLocal Then
              ' ref side is molded, ref side is primary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dRefSideFirstBevelAngle
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dAntiRefSideFirstBevelAngle
           End If
           
        Case SECONDARY_SIDE_GROOVE
           If bRefIsMoldedLocal Then
              ' ref side is molded, anti-ref side is secondary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGroove
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGroove
           End If
           
        Case SECONDARY_SIDE_ACTUAL_THROAT_THICKNESS
           If bRefIsMoldedLocal Then
              ' ref side is molded, anti-ref side is secondary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dAntiRefSideFirstBevelDepth
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dRefSideFirstBevelDepth
           End If
           
        Case SECONDARY_SIDE_NOMINAL_THROAT_THICKNESS
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dAntiMoldedFillet
           
        Case SECONDARY_SIDE_GROOVE_ANGLE
           If bRefIsMoldedLocal Then
              ' ref side is molded, anti-ref side is secondary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dAntiRefSideFirstBevelAngle
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dRefSideFirstBevelAngle
           End If
           
        Case PRIMARY_SIDE_SYMBOL
           If dMoldedFillet > 0 Then
              If sWeldName = STAGGERED_WELD Then
                 arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_STAGGERED_FILLET ' 2
              Else
                 arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_FILLET ' 1
              End If
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_NONE ' 0
           End If
                      
        Case SECONDARY_SIDE_SYMBOL
           If dAntiMoldedFillet > 0 Then
              If sWeldName = STAGGERED_WELD Then
                 arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_STAGGERED_FILLET ' 2
              Else
                 arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_FILLET ' 1
              End If
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_NONE ' 0
           End If
           
        Case Else
     End Select
  Next
   
  ' Set IJWeldingSymbol parameter values
  SetWeldParmValues _
               pPRL, _
               nParmCount, _
               arrayIJWeldingSymbolParmInfo
  
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SetCalculatedTeeWeldParams").Number
End Sub

' Set the calculated weld parameter info for BUTT weld types
'
' The decisions can be made based on the specific weld type passed in, but all BUTT weld types
' can use the same logic, as defined here.  Customers may want to modify this behavior for
' specific classes of BUTT weld.
'
'This procedure will output the values provided by the user if they have been overriden.
'
' The input values, except for oPRL, sWeldName, and GrooveType correspond to the actual values stored by
' the parm rule, not the intermediate name used within the parm rule.  For example, some rules may carry
' a value as "refSide", but then store it as "antiRefSide" based on some other flag.  We need the value
' actually stored.
'
' For each of these input values, we will determine if the value was overriden by the user and use the
' value from the parameter grid if it has been overridden
'
' bRefIsMolded refers to whether the ref side of the reference part is molded, not the NR part
' this can be overridden if the user has overridden the "ReferenceSide" parameter
'
Public Sub SetCalculatedButtWeldParams( _
            ByVal pPRL As IJDParameterLogic, _
            ByVal sWeldName As String, _
            ByVal bRefIsMolded As Boolean, _
            ByVal dRefSideFirstBevelDepth As Double, _
            ByVal dAntiRefSideFirstBevelDepth As Double, _
            ByVal dRefSideFirstBevelAngle As Double, _
            ByVal dAntiRefSideFirstBevelAngle As Double, _
            ByVal dNRRefSideFirstBevelDepth As Double, _
            ByVal dNRAntiRefSideFirstBevelDepth As Double, _
            ByVal dNRRefSideFirstBevelAngle As Double, _
            ByVal dNRAntiRefSideFirstBevelAngle As Double, _
            ByVal GrooveType As Long, _
            ByVal dRootGap As Double, _
            ByVal dNRRootGap As Double)

  On Error GoTo ErrorHandler
  
  ' values set by inputs
  ' first, set the intermediate and input values based on common rules
  
  ' now check all symbol parameter values used in setting weld parameter values to see which are overridden
  ' there are a total of 10 values that can be overridden, but not all are valid for all weld types
  ' Create the array with 10 values and initialize them to "ParmUnsupported".  We will set the valid ones
  ' based on the cross section.  The following is the full set of values and proper alignment
  '
  ' arrayBevelParmInfo(1).sName = "RefSideFirstBevelDepth"
  ' arrayBevelParmInfo(2).sName = "RefSideFirstBevelAngle"
  ' arrayBevelParmInfo(3).sName = "AntiRefSideFirstBevelDepth"
  ' arrayBevelParmInfo(4).sName = "AntiRefSideFirstBevelAngle"
  ' arrayBevelParmInfo(5).sName = "NRRefSideFirstBevelDepth"
  ' arrayBevelParmInfo(6).sName = "NRRefSideFirstBevelAngle"
  ' arrayBevelParmInfo(7).sName = "NRAntiRefSideFirstBevelDepth"
  ' arrayBevelParmInfo(8).sName = "NRAntiRefSideFirstBevelAngle"
  ' arrayBevelParmInfo(9).sName = "RootGap"
  ' arrayBevelParmInfo(10).sName = "NRRootGap"
  ' arrayBevelParmInfo(11).sName = "ReferenceSide"
  
  ' local variable to handle possible override of reference side
  Dim bRefIsMoldedLocal As Boolean
  bRefIsMoldedLocal = bRefIsMolded
  
  ' initialize
  Dim arrayBevelParmInfo(1 To 11) As PARAMETER_INFO
  Dim index As Long
  For index = 1 To 11
     arrayBevelParmInfo(index).sName = "ParmUnsupported"
  Next index

  Select Case sWeldName
     
     Case BUTT_WELD_I
        arrayBevelParmInfo(9).sName = "RootGap"
        arrayBevelParmInfo(10).sName = "NRRootGap"
        arrayBevelParmInfo(11).sName = "ReferenceSide"
        
     Case BUTT_WELD_IV, BUTT_WELD_IX, BUTT_WELD_K, BUTT_WELD_V, BUTT_WELD_X, BUTT_WELD_Y
        arrayBevelParmInfo(1).sName = "RefSideFirstBevelDepth"
        arrayBevelParmInfo(2).sName = "RefSideFirstBevelAngle"
        arrayBevelParmInfo(3).sName = "AntiRefSideFirstBevelDepth"
        arrayBevelParmInfo(4).sName = "AntiRefSideFirstBevelAngle"
        arrayBevelParmInfo(5).sName = "NRRefSideFirstBevelDepth"
        arrayBevelParmInfo(6).sName = "NRRefSideFirstBevelAngle"
        arrayBevelParmInfo(7).sName = "NRAntiRefSideFirstBevelDepth"
        arrayBevelParmInfo(8).sName = "NRAntiRefSideFirstBevelAngle"
        arrayBevelParmInfo(9).sName = "RootGap"
        arrayBevelParmInfo(10).sName = "NRRootGap"
        arrayBevelParmInfo(11).sName = "ReferenceSide"
              
  End Select
    
  ' get any overridden values
  GetSymbolParameterInfo pPRL.SmartOccurrence, 11, arrayBevelParmInfo
  
  ' update any overridden values
  If arrayBevelParmInfo(1).bOverridden = True Then
     dRefSideFirstBevelDepth = arrayBevelParmInfo(1).vValue
  End If
    
  If arrayBevelParmInfo(2).bOverridden = True Then
     dRefSideFirstBevelAngle = arrayBevelParmInfo(2).vValue
  End If
    
  If arrayBevelParmInfo(3).bOverridden = True Then
     dAntiRefSideFirstBevelDepth = arrayBevelParmInfo(3).vValue
  End If
    
  If arrayBevelParmInfo(4).bOverridden = True Then
     dAntiRefSideFirstBevelAngle = arrayBevelParmInfo(4).vValue
  End If
    
  If arrayBevelParmInfo(5).bOverridden = True Then
     dNRRefSideFirstBevelDepth = arrayBevelParmInfo(5).vValue
  End If
    
  If arrayBevelParmInfo(6).bOverridden = True Then
     dNRRefSideFirstBevelAngle = arrayBevelParmInfo(6).vValue
  End If
    
  If arrayBevelParmInfo(7).bOverridden = True Then
     dNRAntiRefSideFirstBevelDepth = arrayBevelParmInfo(7).vValue
  End If
    
  If arrayBevelParmInfo(8).bOverridden = True Then
     dNRAntiRefSideFirstBevelAngle = arrayBevelParmInfo(8).vValue
  End If
    
  If arrayBevelParmInfo(9).bOverridden = True Then
     dRootGap = arrayBevelParmInfo(9).vValue
  End If
    
  If arrayBevelParmInfo(10).bOverridden = True Then
     dNRRootGap = arrayBevelParmInfo(10).vValue
  End If
  
  If arrayBevelParmInfo(11).bOverridden = True Then
     If LCase(arrayBevelParmInfo(11).vValue) = "molded" Then
        bRefIsMoldedLocal = True
     Else
        bRefIsMoldedLocal = False
     End If
  End If
  
  
  ' set the controlling values for IJWeldSymbol using the correct values of the input parameters
  ' the inputs from the model have the bevel values stored by part, we just need to keep the deepest
  ' one on each side.
  ' NOTE:  We could do some additional error checking here, such as making sure that both parts have the same
  ' bevel depth for a V or one is 0 for a BEVEL, but this will be skipped for performance.  Not sure that we
  ' would want to fail in these cases anyway.
  Dim RefSideGroove As Long
  Dim AntiRefSideGroove As Long
  Dim RefSideGrooveDepth As Double
  Dim AntiRefSideGrooveDepth As Double
  Dim RefSideGrooveAngle As Double
  Dim AntiRefSideGrooveAngle As Double

    
  If sWeldName = BUTT_WELD_I Then
     ' for a ButtWeldI, we know we have square and none for groove types
     RefSideGroove = pcr_WG_Square        ' in the final step, we will force PRIMARY_SIDE square
     AntiRefSideGroove = pcr_WG_Square    ' and SECONDARY_SIDE to none
     RefSideGrooveDepth = 0
     AntiRefSideGrooveDepth = 0
     RefSideGrooveAngle = 0
     AntiRefSideGrooveAngle = 0
     
  Else
     ' not a ButtWeldI, the input groove type controls the type for any side that has a bevel
     ' any side that has no bevel (depth = 0) will have no groove
     If dRefSideFirstBevelDepth < 0.0001 And dNRRefSideFirstBevelDepth < 0.0001 Then
        ' there is no primary bevel, force groove depth and bevel angle to 0 and groove type to none
        RefSideGroove = pcr_WG_None
        RefSideGrooveDepth = 0
        RefSideGrooveAngle = 0
     Else
        ' there is a primary bevel, use input groove type and set the proper depth
        RefSideGroove = GrooveType
        If dRefSideFirstBevelDepth > dNRRefSideFirstBevelDepth Then
           ' use the values from primary part
           RefSideGrooveDepth = dRefSideFirstBevelDepth
        Else
           ' use the values from NR part
           RefSideGrooveDepth = dNRRefSideFirstBevelDepth
        End If
        ' the angle is the sum of the part bevel angles
        RefSideGrooveAngle = dRefSideFirstBevelAngle + dNRRefSideFirstBevelAngle
     End If
    
     If dAntiRefSideFirstBevelDepth < 0.0001 And dNRAntiRefSideFirstBevelDepth < 0.0001 Then
        ' there is no secondary bevel, force groove depth and bevel angle to 0 and groove type to none
        AntiRefSideGroove = pcr_WG_None
        AntiRefSideGrooveDepth = 0
        AntiRefSideGrooveAngle = 0
     Else
        ' there is a primary bevel, use input groove type and set the proper depth
        AntiRefSideGroove = GrooveType
        If dAntiRefSideFirstBevelDepth > dNRAntiRefSideFirstBevelDepth Then
           ' use the values from primary part
           AntiRefSideGrooveDepth = dAntiRefSideFirstBevelDepth
        Else
           ' use the values from NR part
           AntiRefSideGrooveDepth = dNRAntiRefSideFirstBevelDepth
        End If
        ' the angle is the sum of the part bevel angles
        AntiRefSideGrooveAngle = dAntiRefSideFirstBevelAngle + dNRAntiRefSideFirstBevelAngle
     End If
     
  End If
  
  ' now set the output values
  ' the proper values to use are as follows:
  ' NOTE: Ref/AntiRef side values are reversed from this if the ref side is not molded
  ' "PrimarySideGroove" = RefSideGroove
  ' "PrimarySideActualThroatThickness" = RefSideGrooveDepth
  ' "PrimarySideRootOpening" = dRootGap + dNRRootGap
  ' "PrimarySideGrooveAngle" = RefSideGrooveAngle
  ' "SecondarySideGroove" = AntiRefSideGroove
  ' "SecondarySideActualThroatThickness" = AntiRefSideGrooveDepth
  ' "SecondarySideGrooveAngle" = AntiRefSideGrooveAngle
  
  ' however, we need to write out all of these computed parameters into the parameter info
  ' structure by getting all of the possible computed values, updating the ones that are
  ' actually updated by this rule, and writing all of them out.  Otherwise, we would get
  ' errors caused by the ones that are declared as outputs, but that we are not outputting
   
  Dim arrayIJWeldingSymbolParmInfo() As PARAMETER_INFO
  Dim nParmCount As Integer
  Dim nParmIndex As Integer
   
  ' get all of the IJWeldingSymbol parameters
  ' note the code below is intended to make the MOLDED side correspond to the PRIMARY side
  GetWeldParmInfo sWeldName, arrayIJWeldingSymbolParmInfo, nParmCount
  For nParmIndex = 1 To nParmCount
     Select Case arrayIJWeldingSymbolParmInfo(nParmIndex).sName
     
        Case PRIMARY_SIDE_GROOVE
           If bRefIsMoldedLocal Then
              ' ref side is molded, ref side is primary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGroove
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGroove
           End If
           
        Case PRIMARY_SIDE_ACTUAL_THROAT_THICKNESS
           If bRefIsMoldedLocal Then
              ' ref side is molded, ref side is primary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGrooveDepth
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGrooveDepth
           End If
           
        Case PRIMARY_SIDE_ROOT_OPENING, SECONDARY_SIDE_ROOT_OPENING
           arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dRootGap + dNRRootGap
           
        Case PRIMARY_SIDE_GROOVE_ANGLE
           If bRefIsMoldedLocal Then
              ' ref side is molded, ref side is primary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGrooveAngle
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGrooveAngle
           End If
           
        Case SECONDARY_SIDE_GROOVE
           If bRefIsMoldedLocal Then
              ' ref side is molded, anti-ref side is secondary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGroove
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGroove
           End If
           
           ' the secondary side should not be square.  If it is, force it to none.
           If arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = pcr_WG_Square Then
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = pcr_WG_None
           End If
           
        Case SECONDARY_SIDE_ACTUAL_THROAT_THICKNESS
           If bRefIsMoldedLocal Then
              ' ref side is molded, anti-ref side is secondary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGrooveDepth
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGrooveDepth
           End If
           
        Case SECONDARY_SIDE_GROOVE_ANGLE
           If bRefIsMoldedLocal Then
              ' ref side is molded, anti-ref side is secondary
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = AntiRefSideGrooveAngle
           Else
              arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = RefSideGrooveAngle
           End If
           
        Case Else
     End Select
  Next
   
  ' Set IJWeldingSymbol parameter values
  SetWeldParmValues _
               pPRL, _
               nParmCount, _
               arrayIJWeldingSymbolParmInfo
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SetCalculatedButtWeldParams").Number
End Sub

' Set the calculated weld parameter info for BUTT weld types
'
' The decisions can be made based on the specific weld type passed in, but all BUTT weld types
' can use the same logic, as defined here.  Customers may want to modify this behavior for
' specific classes of BUTT weld.
'
'This procedure will output the values provided by the user if they have been overriden.
'
' The input values, except for oPRL, sWeldName, and GrooveType correspond to the actual values stored by
' the parm rule, not the intermediate name used within the parm rule.  For example, some rules may carry
' a value as "refSide", but then store it as "antiRefSide" based on some other flag.  We need the value
' actually stored.
'
' For each of these input values, we will determine if the value was overriden by the user and use the
' value from the parameter grid if it has been overridden
'
Public Sub SetCalculatedLapWeldParams( _
            ByVal pPRL As IJDParameterLogic, _
            ByVal sWeldName As String, _
            ByVal dMoldedFillet As Double, _
            ByVal dAntiMoldedFillet As Double)

   On Error GoTo ErrorHandler
  
   ' values set by inputs
   ' first, set the intermediate and input values based on common rules
  
   ' now check all symbol parameter values used in setting weld parameter values to see which are overridden
   ' there are a total of 2 values that can be overridden.  The following is the full set of values
   ' and proper alignment.
   '
   ' arrayBevelParmInfo(1).sName = "MoldedFillet"
   ' arrayBevelParmInfo(2).sName = "AntiMoldedFillet"
   ' arrayBevelParmInfo(3).sName = "FilletMeasureMethod"
  
   ' For IJWeldingSymbol
   '   PRIMARY_SIDE_NOMINAL_THROAT_THICKNESS is driven by MoldedFillet
   '   SECONDARY_SIDE_NOMINAL_THROAT_THICKNESS is driven by AntiMoldedFillet
   '
   ' Get MoldedFillet and AntiMoldedFillet
   '
   Dim dPrimarySideNominalThroatThickness As Double
   Dim dSecondarySideNominalThroatThickness As Double
      
   ' Default to rule based bevel values
   dPrimarySideNominalThroatThickness = dMoldedFillet
   dSecondarySideNominalThroatThickness = dAntiMoldedFillet
   
   ' Check if user overrode them
   Dim arrayBevelParmInfo(1 To 3) As PARAMETER_INFO
   
   arrayBevelParmInfo(1).sName = "MoldedFillet"
   arrayBevelParmInfo(2).sName = "AntiMoldedFillet"
   arrayBevelParmInfo(3).sName = "FilletMeasureMethod"
   
   GetSymbolParameterInfo _
                    pPRL.SmartOccurrence, _
                    3, _
                    arrayBevelParmInfo
   
   If arrayBevelParmInfo(1).bOverridden = True Then
      dPrimarySideNominalThroatThickness = arrayBevelParmInfo(1).vValue
   End If
   
   If arrayBevelParmInfo(2).bOverridden = True Then
      dSecondarySideNominalThroatThickness = arrayBevelParmInfo(2).vValue
   End If

   ' Calculate leg value if Fillets are measured as throat
   If arrayBevelParmInfo(3).vValue = 65537 Then
      dPrimarySideNominalThroatThickness = dPrimarySideNominalThroatThickness * Sqr(2)
      dSecondarySideNominalThroatThickness = dSecondarySideNominalThroatThickness * Sqr(2)
   End If
   
   ' now we need to write out all of the computed parameters into the parameter info
   ' structure by getting all of the possible computed values, updating the ones that
   ' are actually updated by this rule, and writing all of them out
   
   Dim arrayIJWeldingSymbolParmInfo() As PARAMETER_INFO
   Dim nParmCount As Integer
   Dim nParmIndex As Integer
   
   ' get all of the IJWeldingSymbol parameters
   GetWeldParmInfo sWeldName, arrayIJWeldingSymbolParmInfo, nParmCount
   For nParmIndex = 1 To nParmCount
      Select Case arrayIJWeldingSymbolParmInfo(nParmIndex).sName
      
         Case PRIMARY_SIDE_NOMINAL_THROAT_THICKNESS
            arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dPrimarySideNominalThroatThickness
            
         Case SECONDARY_SIDE_NOMINAL_THROAT_THICKNESS
            arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = dSecondarySideNominalThroatThickness
            
         Case PRIMARY_SIDE_SYMBOL
            If dPrimarySideNominalThroatThickness > 0 Then
               arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_FILLET ' 1
            Else
               arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_NONE ' 0
            End If
            
         Case SECONDARY_SIDE_SYMBOL
            If dSecondarySideNominalThroatThickness > 0 Then
               arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_FILLET ' 1
            Else
               arrayIJWeldingSymbolParmInfo(nParmIndex).vValue = WELD_SYMBOL_NONE ' 0
            End If
         
         Case Else
      End Select
   Next

   ' Set IJWeldingSymbol parameter values
   SetWeldParmValues _
                pPRL, _
                nParmCount, _
                arrayIJWeldingSymbolParmInfo
                
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SetCalculatedLapWeldParams").Number
End Sub
Public Function GetMoldedSide(oPCObj As Object, Optional oConnectObj As Object) As String
On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim sMoldedSide As String
    
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = oPCObj
    
    If oConnectObj Is Nothing Then
        Set oConnectObj = oPhysConn.ConnectedObject1
    End If

    Dim pHelper As New StructDetailObjects.Helper

    Select Case pHelper.ObjectType(oConnectObj)
        Case SDOBJECT_STIFFENER
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = oConnectObj
            If oProfilePart.sectionType = "HalfR" Or oProfilePart.sectionType = "R" Then
            Exit Function
        End If
        
        Dim oSysChild As IJDesignChild
        Set oSysChild = oPhysConn.object
            
        If TypeOf oSysChild.GetParent Is IJStructFeature Then
            Dim objWebOrFlange As IJStructFeature
            Set objWebOrFlange = oSysChild.GetParent
            
            If objWebOrFlange.get_StructFeatureType = SF_WebCut Then
                    'Mountig face considered as bottom always
                 Select Case oProfilePart.loadPoint
                 Case 2, 3, 4, 5, 6
                    sMoldedSide = "WebLeft"
                 Case 20, 21, 22, 23, 24
                    sMoldedSide = "WebRight"
                 Case 1, 25 'centered
                    sMoldedSide = "WebLeft"
                 Case Else
                End Select
            ElseIf objWebOrFlange.get_StructFeatureType = SF_FlangeCut Then
                Dim oFlangeObj As New StructDetailObjects.FlangeCut
                Set oFlangeObj.object = objWebOrFlange
                Dim bIsTopFlange As Boolean
                bIsTopFlange = False
                bIsTopFlange = oFlangeObj.IsTopFlange
                If bIsTopFlange Then
                    Select Case oProfilePart.loadPoint
                             Case 2, 3, 4, 5, 6
                                sMoldedSide = "TopFlangeTopFace"
                             Case 20, 21, 22, 23, 24
                                sMoldedSide = "TopFlangeBottomFace"
                             Case 1, 25 'centered
                                sMoldedSide = "TopFlangeTopFace"
                             Case Else
                            End Select
                Else
                    Select Case oProfilePart.loadPoint
                             Case 2, 3, 4, 5, 6
                                sMoldedSide = "BottomFlangeBottomFace"
                             Case 20, 21, 22, 23, 24
                                sMoldedSide = "BottomFlangeTopFace"
                             Case 1, 25 'centered
                                sMoldedSide = "BottomFlangeBottomFace"
                             Case Else
                            End Select
                End If
            End If
        Else
                'Bounded object is a stiffener but its parent is neither a web-cut
                'nor a flange-cut
                 Select Case oProfilePart.loadPoint
                 Case 1 To 6, 25
                    sMoldedSide = "WebLeft"
                 Case 20 To 24
                    sMoldedSide = "WebRight"
                 Case Else
                End Select
        End If
       
    End Select
    GetMoldedSide = sMoldedSide
Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetRefSide", strError).Number
End Function

Public Sub CheckAndSetUpSplitOprn(oPhysCon As IJStructPhysicalConnection)
On Error GoTo ErrorHandler
    
    Dim oModelBody          As IJDModelBody
    Dim lLumpCount          As Long
    Dim lShellCount         As Long
    Dim lFaceCount          As Long
    Dim lWireCount          As Long
    Dim lLoopCount          As Long
    Dim lEdgeCount          As Long
    Dim lCoEdgeCount        As Long
    Dim lVertexCount        As Long
    Dim bIsBodyValid        As Boolean
    Dim oAssocRel           As IJDAssocRelation
    Dim oTargetObjCol       As IJDTargetObjectCol
    Dim oTargetObjCol2      As IJDTargetObjectCol
    Dim oEntityOprn         As IJDStructEntityOperation
    Dim oSplitPCAE          As IJPhyConSplitEntity
    Dim oOperColl           As IJElements
    Dim oSDOPhysConn        As StructDetailObjects.PhysicalConn
    
    Set oSDOPhysConn = New StructDetailObjects.PhysicalConn
    Set oSDOPhysConn.object = oPhysCon
    
    'At least one of the input objects must be a Plate
    If oSDOPhysConn.ConnectedObject1Type <> SDOBJECT_PLATE And _
        oSDOPhysConn.ConnectedObject2Type <> SDOBJECT_PLATE Then
        Exit Sub
    End If
    
    On Error Resume Next
    Set oModelBody = oPhysCon
    
    On Error GoTo ErrorHandler
    If Not oModelBody Is Nothing Then
        oModelBody.CheckTopology vbNullString, lLumpCount, lShellCount, lWireCount, lFaceCount, lLoopCount, _
                                    lCoEdgeCount, lEdgeCount, lVertexCount, bIsBodyValid
                                    
        If bIsBodyValid And lLumpCount > 0 Then
            Set oAssocRel = oPhysCon
            Set oTargetObjCol = oAssocRel.CollectionRelations _
                                            (IID_IJStructGeometry, "StructOperation_OPRND_DEST")
            
            If Not oTargetObjCol Is Nothing Then
            'We have an AE, find out if the split operation is a normal one
                If oTargetObjCol.Count > 0 Then
                    Set oSplitPCAE = oTargetObjCol.Item(1)
                    Set oAssocRel = oSplitPCAE
                    
                    Set oTargetObjCol2 = oAssocRel.CollectionRelations _
                                                    (IID_IJStructSplit, "StructSplit_OPER1_ORIG")
                                                    
                    If Not oTargetObjCol2 Is Nothing Then
                        If oTargetObjCol2.Count > 0 Then
                            oSplitPCAE.SplitType = STRDET_SPLITBY_BOTH
                        Else
                            If lLumpCount = 1 Then
                                Dim oIJDObject      As IJDObject
                                Set oIJDObject = oSplitPCAE
                                
                                oIJDObject.Remove
                            Else
                                oSplitPCAE.SplitType = STRDET_SPLITBY_INNERCONTOURS
                            End If
                        End If
                    End If
                Else
                    If lLumpCount > 1 Then
                        Set oEntityOprn = oPhysCon
                        Set oOperColl = New JObjectCollection
                        
                        oEntityOprn.SetEntityOperation "PhyConSplitEntity.CPhyConSplitEntity_AE.1", oOperColl
                        
                        Set oTargetObjCol = oAssocRel.CollectionRelations _
                                                        (IID_IJStructGeometry, "StructOperation_OPRND_DEST")
                                                        
                        If Not oTargetObjCol Is Nothing Then
                            If oTargetObjCol.Count > 0 Then
                                Set oSplitPCAE = oTargetObjCol.Item(1)
                                oSplitPCAE.SplitType = STRDET_SPLITBY_INNERCONTOURS
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CheckAndSetUpSplitOprn").Number
End Sub

'TR 171555 -- specific case to be handled to obtain MoldedOffset

Public Function IsConnectionBetweenTubeAndCAN(oPort1 As IJPort, oPort2 As IJPort, dOffset As Double) As Boolean
        
        IsConnectionBetweenTubeAndCAN = False
        dOffset = 0#
        
        Dim bBuiltUp1 As Boolean
        Dim bBuiltUp2 As Boolean
        Dim oBuiltUpMember1 As ISPSDesignedMember
        Dim oBuiltUpMember2 As ISPSDesignedMember
        Dim pPOM As IJDPOM
        Dim oLine As IJLine
        Dim oLineDir As IJDVector
        Dim pLine As IJLine
        Dim oLineDir1 As IJDVector
        Dim oSurfaceBody1 As IJSurfaceBody
        Dim oSurfaceBody2 As IJSurfaceBody
        Dim ppNormal1 As IJDVector
        Dim ppNormal2 As IJDVector
        Dim oMemberLine As IJLine
        
        Dim oMemberPart1 As New StructDetailObjects.MemberPart
        Dim oMemberPart2 As New StructDetailObjects.MemberPart
        Dim oModelBody1 As IJModelBody
        Dim oModelBody2 As IJModelBody
        Dim oCanRule As ISPSCanRule
        Dim oSpsCanRuleStatus As SPSCanRuleStatus
        Dim oPrimaryMemberSystem As ISPSMemberSystem
        Dim oShipGeomOps As GSCADShipGeomOps.SGOModelBodyUtilities
'        Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
        Dim ppPointOnFirstBody As IJDPosition
        Dim ppPointOnSecondBody As IJDPosition
        
        Dim oRefPlate As New StructDetailObjects.PlatePart
        Dim oNonRefPlate As New StructDetailObjects.PlatePart
        
        Set pPOM = Nothing
        Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
        Set oShipGeomOps = New GSCADShipGeomOps.SGOModelBodyUtilities
        
        If TypeOf oPort1.Connectable Is IJPlatePart And TypeOf oPort2.Connectable Is IJPlatePart Then
            Set oRefPlate.object = oPort1.Connectable
            Set oNonRefPlate.object = oPort2.Connectable
        Else
            Exit Function
        End If

        IsPortFromBuiltUpMember oPort1, bBuiltUp1, oBuiltUpMember1
        IsPortFromBuiltUpMember oPort2, bBuiltUp2, oBuiltUpMember2

        If bBuiltUp1 And bBuiltUp2 Then
            Set oMemberPart1.object = oBuiltUpMember1
            Set oMemberPart2.object = oBuiltUpMember2
            'check needed to ensure it is a can type.
            Dim oSectionName1 As String
            Dim oSectionName2 As String
            oSectionName1 = oMemberPart1.sectionType
            oSectionName2 = oMemberPart2.sectionType
            If oSectionName1 = "BUCan" And oSectionName2 = "BUTube" Then
                Set oCanRule = oMemberPart1.CanRule
                oSpsCanRuleStatus = oCanRule.GetPrimaryMemberSystem(oPrimaryMemberSystem)
            ElseIf oSectionName2 = "BUCan" And oSectionName1 = "BUTube" Then
                Set oCanRule = oMemberPart2.CanRule
                oSpsCanRuleStatus = oCanRule.GetPrimaryMemberSystem(oPrimaryMemberSystem)
            End If
            If Not oPrimaryMemberSystem Is Nothing Then
                'get the direction of the PrimaryMemberSystem
                Set pLine = oPrimaryMemberSystem.LogicalAxis.CurveGeometry
                Set oMemberLine = Line_FromPositions(pPOM, Position_FromLine(pLine, 0), Position_FromLine(pLine, 1))
                'This gives the Direction of the PrimaryMemberSystem along which the CAN is placed.
                Set oLineDir1 = Vector_FromLine(oMemberLine)
                
                Set oModelBody1 = oTopologyLocate.GetPlateParentBodyModel(oRefPlate.object)
                Set oModelBody2 = oTopologyLocate.GetPlateParentBodyModel(oNonRefPlate.object)

                oShipGeomOps.GetClosestPointsBetweenTwoBodies oModelBody1, oModelBody2, ppPointOnFirstBody, ppPointOnSecondBody, dOffset
         
                If TypeOf oModelBody1 Is IJSurfaceBody Then
                    Set oSurfaceBody1 = oModelBody1
                    oSurfaceBody1.GetNormalFromPosition ppPointOnFirstBody, ppNormal1
                End If
                If TypeOf oModelBody2 Is IJSurfaceBody Then
                    Set oSurfaceBody2 = oModelBody2
                    oSurfaceBody2.GetNormalFromPosition ppPointOnSecondBody, ppNormal2
                End If
                If oSectionName1 = "BUTube" Then
                    If Round(oLineDir1.Dot(ppNormal1), 3) = 0 Then
                        IsConnectionBetweenTubeAndCAN = True
                    End If
                ElseIf oSectionName2 = "BUTube" Then
                    If Round(oLineDir1.Dot(ppNormal2), 3) = 0 Then
                        IsConnectionBetweenTubeAndCAN = True
                    End If
                End If
                If IsConnectionBetweenTubeAndCAN Then
                    Set oLine = Line_FromPositions(pPOM, ppPointOnFirstBody, ppPointOnSecondBody)
                    Set oLineDir = Vector_FromLine(oLine)
                    'find out if the offset is +ve or -ve. with respect to molded surface normal of re
                    'plate, measure the offset.part 1 is always ref part
            
                    Dim ddotproduct As Double
                    ddotproduct = ppNormal1.Dot(oLineDir)
                    If ddotproduct > 0 Then
                        dOffset = dOffset
                    Else
                        dOffset = -dOffset
                    End If
                End If
                
                Set oBuiltUpMember1 = Nothing
                Set oBuiltUpMember2 = Nothing
                Set oRefPlate = Nothing
                Set oNonRefPlate = Nothing
                Set oMemberPart1 = Nothing
                Set oMemberPart2 = Nothing
                Set oSurfaceBody1 = Nothing
                Set oModelBody1 = Nothing
                Set oModelBody2 = Nothing
                Set oLine = Nothing
            End If
    End If
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetButtWeldOverlappingThicknessAndAdditions").Number
End Function


'***********************************************************************
' METHOD:  IsPortFromBuiltUpMember
'
' DESCRIPTION:  Compare the normal of the molded surface of two
'               connected plate parts. If they point to the same
'               direction, then return true. If not, return false.
'***********************************************************************
Public Sub IsPortFromBuiltUpMember(oPort As IJPort, _
                                   bFromBuiltUp As Boolean, _
                                   Optional oBuiltupMember As ISPSDesignedMember)
Const METHOD = "::IsPortFromBuiltUpMember"
On Error GoTo ErrorHandler
    
    bFromBuiltUp = False
    
    Dim oParentObject As Object
    Dim oPlateSystem As IJPlateSystem
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_PlateSystem As StructDetailObjects.PlateSystem
    
    ' Check if given port is from a PlatePart
    If TypeOf oPort.Connectable Is IJPlatePart Then
        ' Given Port's Connectable is IJplatePart
        ' Get the Plate Part's Parent Object
        Set oSDO_PlatePart = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart.object = oPort.Connectable
        Set oParentObject = oSDO_PlatePart.ParentSystem
        Set oSDO_PlatePart = Nothing
            
        ' Check if the Plate Part's Parent object is IJPlateSystem
        If TypeOf oParentObject Is IJPlateSystem Then
            ' Plate Part's Parent object is IJPlateSystem
            ' Get the Plate Systems's Parent object
            Set oSDO_PlateSystem = New StructDetailObjects.PlateSystem
            Set oSDO_PlateSystem.object = oParentObject
            Set oParentObject = oSDO_PlateSystem.ParentSystem
            Set oSDO_PlateSystem = Nothing
            
            ' Check if the Plate System's Parent object is IJPlateSystem
            If TypeOf oParentObject Is IJPlateSystem Then
                Set oPlateSystem = oParentObject
                bFromBuiltUp = oPlateSystem.IsBuiltupPlateSystem
                
                If bFromBuiltUp Then
                    Set oBuiltupMember = oPlateSystem.ParentBuiltup
                End If
            End If
            
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub


Function Position_FromLine(pLine As IJLine, iIndex As Integer) As IJDPosition
    ' create new position
    Dim pPosition As IJDPosition
    Set pPosition = New DPosition

    ' retrieve coordinates
    Dim x As Double, y As Double, z As Double
    If iIndex = 0 Then
        Call pLine.GetStartPoint(x, y, z)
    Else
        Call pLine.GetEndPoint(x, y, z)
    End If
    
    ' set coordinates
    Call pPosition.Set(x, y, z)
    
    ' return result
    Set Position_FromLine = pPosition
End Function


Function Vector_FromLine(pLine As IJLine) As IJDVector
    ' get vector
    Dim u As Double, v As Double, w As Double
    Call pLine.GetDirection(u, v, w)
    
    ' create result
    Dim pVector As New DVector
    Call pVector.Set(u, v, w)
    
    ' return result
    Set Vector_FromLine = pVector
End Function

Public Function Line_FromPositions(pPOM As IJDPOM, pPositionOfStartPoint As IJDPosition, pPositionOfEndPoint As IJDPosition) As IJLine
    ' create new line
    Dim pGeometryFactory As New GeometryFactory
    Set Line_FromPositions = pGeometryFactory.Lines3d.CreateBy2Points(pPOM, _
        pPositionOfStartPoint.x, pPositionOfStartPoint.y, pPositionOfStartPoint.z, _
        pPositionOfEndPoint.x, pPositionOfEndPoint.y, pPositionOfEndPoint.z)
End Function
