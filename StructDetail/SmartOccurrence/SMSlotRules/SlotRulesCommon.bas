Attribute VB_Name = "SlotRulesCommon"
' Each individual rule should refer to project name to make customization easier
' DI-CP-235957  StrDet: Invalid DB Col Name issues (31 Char limit) view gen. on Oracle DB
Public Const m_sProjectName As String = CUSTOMERID + "SlotRules"
Public Const m_sProjectPath As String = "S:\StructDetail\Data\SmartOccurrence\" + m_sProjectName + "\"

Public Const PARAM_TOP_FLANGE_LEFT_TOP_CORNER_RADIUS As String = "TopFlangeLeftTopCornerRadius"
Public Const PARAM_TOP_FLANGE_LEFT_TOP_CLEARANCE As String = "TopFlangeLeftClearance"

Public Const PARAM_TOP_CLEARANCE As String = "TopClearance"

Public Const PARAM_TOP_FLANGE_RIGHT_TOP_CLEARANCE As String = "TopFlangeRightTopClearance"
Public Const PARAM_TOP_FLANGE_RIGHT_TOP_CORNER_RADIUS As String = "TopFlangeRightTopCornerRadius"
Public Const PARAM_TOP_FLANGE_RIGHT_CLEARANCE As String = "TopFlangeRightClearance"
Public Const PARAM_TOP_FLANGE_RIGHT_BOTTOM_CORNER_RADIUS As String = "TopFlgRightBtmCnrRadius"
Public Const PARAM_TOP_FLANGE_RIGHT_BOTTOM_CLEARANCE As String = "TopFlangeRightBottomClearance"

Public Const PARAM_WEB_RIGHT_TOP_CORNER_RADIUS As String = "WebRightTopCornerRadius"
Public Const PARAM_WEB_RIGHT_CLEARANCE As String = "WebRightClearance"
Public Const PARAM_WEB_RIGHT_BOTTOM_CORNER_RADIUS As String = "WebRightBottomCornerRadius"

Public Const PARAM_BOTTOM_FLANGE_RIGHT_TOP_CLEARANCE As String = "BottomFlangeRightTopClearance"
Public Const PARAM_BOTTOM_FLANGE_RIGHT_TOP_CORNER_RADIUS As String = "BtmFlgRightTopCnrRadius"
Public Const PARAM_BOTTOM_FLANGE_RIGHT_CLEARANCE = "BottomFlangeRightClearance"

Public Const PARAM_BOTTOM_FLANGE_LEFT_CLEARANCE = "BottomFlangeLeftClearance"
Public Const PARAM_BOTTOM_FLANGE_LEFT_TOP_CORNER_RADIUS As String = "BtmFlgLeftTopCnrRadius"
Public Const PARAM_BOTTOM_FLANGE_LEFT_TOP_CLEARANCE As String = "BottomFlangeLeftTopClearance"

Public Const PARAM_WEB_LEFT_BOTTOM_CORNER_RADIUS As String = "WebLeftBottomCornerRadius"
Public Const PARAM_WEB_LEFT_CLEARANCE As String = "WebLeftClearance"
Public Const PARAM_WEB_LEFT_TOP_CORNER_RADIUS As String = "WebLeftTopCornerRadius"

Public Const PARAM_TOP_FLANGE_LEFT_BOTTOM_CLEARANCE As String = "TopFlangeLeftBottomClearance"
Public Const PARAM_TOP_FLANGE_LEFT_BOTTOM_CORNER_RADIUS As String = "TopFlgLeftBtmCnrRadius"
Public Const PARAM_TOP_FLANGE_LEFT_CLEARANCE As String = "TopFlangeLeftClearance"

Public Const PARAM_FLANGE_CLEARANCE As String = "FlangeClearance"
Public Const PARAM_WEB_CLEARANCE As String = "WebClearance"
Public Const PARAM_BACK_CLEARANCE As String = "BackClearance"

Public Const PARAM_CORNER_RADIUS As String = "CornerRadius"
Public Const PARAM_CORNER_RADIUS_LOWER As String = "CornerRadiusLower"
Public Const PARAM_SET_BACK As String = "SetBack"

Public Const PARAM_SLOT_ANGLE As String = "SlotAngle"
Public Const PARAM_SLOT_ANGLE_LEFT As String = "SlotAngleLeft"
Public Const PARAM_SLOT_WIDTH As String = "SlotWidth"

Public Const ONE_DEGREE_AS_RADIANT As Double = 0.0174532925

' This method sets the values of the IJUASlotAssyOrientation custom
' atrributes that have been placed on the slot.
' Version 1
' Note: The IJAssemblyOrientation interface is in Level 4 (Planning)
' while this rule is in level 3.  Therefore the reference to the
' interface cannot be made in code that will be delivered with
' IntelliShip.  This code is provided as an example only and requires
' a reference to 'Ingr GSCAD Planning Assembly 1.0 Type Library'.
Public Sub SetSlotAssyOrientation(ByVal pPLH As IJDParameterLogic)
  On Error GoTo ErrorHandler

  ' Use penetrated object to get assembly
  Dim oAssemblyChild As IJAssemblyChild
  Dim oAssembly As IJAssembly
  Dim oSlotWrapper As New Structdetailobjects.Slot
  
  Set oSlotWrapper.object = pPLH.SmartOccurrence
  Set oAssemblyChild = oSlotWrapper.Penetrated
  
  ' Get the Assembly in which the slot exists.
  Set oAssembly = oAssemblyChild.Parent
  
  If oAssembly Is Nothing Then
     ' The part has not been assigned to an assembly yet(so is the slot)
     GoTo Cleanup
  End If
  
  If Not TypeOf oAssembly Is IJLocalCoordinateSystem Then
      'This assembly does not support IJLocalCoordinateSystem,
      'no orientation associated with it,so it won't affect slot
      GoTo Cleanup
  End If
  
  ' Get the orientation of the Assembly.
  Dim oAssemblyOrientation As IJLocalCoordinateSystem
  Dim vector As IJDVector
  
  Set oAssemblyOrientation = oAssembly
  Set vector = oAssemblyOrientation.ZAxis
  
  ' Set the values on the slot attributes.
  Dim oSlot As IJSmartOccurrence
  Dim oAttrHelper As IJDAttributes
  
  Set oSlot = pPLH.SmartOccurrence
  Set oAttrHelper = oSlot

  Dim oAttributeCol As IMSAttributes.IJDAttributesCol
  Dim oAttr As IJDAttribute
  
  Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUASlotAssyOrientation")
  Set oAttr = oAttributeCol.Item("AssyOrientationX")
  oAttr.Value = vector.x
  Set oAttr = oAttributeCol.Item("AssyOrientationY")
  oAttr.Value = vector.y
  Set oAttr = oAttributeCol.Item("AssyOrientationZ")
  oAttr.Value = vector.z
    
  GoTo Cleanup
  
ErrorHandler:
  Err.Raise LogError(Err, "SlotRulesCommon", "SetSlotAssyOrientation").Number
  
Cleanup:
  Set oAssemblyChild = Nothing
  Set oAssembly = Nothing
  Set oAssemblyOrientation = Nothing
  Set vector = Nothing
  Set oSlot = Nothing
  Set oAttrHelper = Nothing
  Set oAttributeCol = Nothing
  Set oAttr = Nothing
  
End Sub
Public Function GetSlotAngle(ByVal oPLH As IJDParameterLogic) As Double
  On Error GoTo ErrorHandler

  Dim strAssyMethod As String
  Dim dSlotAngle As Double
  
  dSlotAngle = ONE_DEGREE_AS_RADIANT / 10
  
  'TR-169311 Trying to retrieve the question's answer without hardcoding
  'the path by caliling GetSelectorAnswer() method
  GetSelectorAnswer oPLH, "AssyMethod", strAssyMethod
  
  Select Case strAssyMethod
      Case "Drop", "Drop at angle", "Vertical drop"
          Dim oSlotWrapper As New Structdetailobjects.Slot
          Dim oAssembly As IJAssembly
          Dim oAssemblyChild As IJAssemblyChild
          
          Set oSlotWrapper.object = oPLH.SmartOccurrence
          On Error Resume Next
          
          If TypeOf oSlotWrapper.Penetrated Is IJAssemblyChild Then
            Set oAssemblyChild = oSlotWrapper.Penetrated
            
            If TypeOf oAssemblyChild.Parent Is IJAssembly Then
                Set oAssembly = oAssemblyChild.Parent
            End If
         End If
         
          Set oAssemblyChild = Nothing
          
          On Error GoTo ErrorHandler
          
          If Not oAssembly Is Nothing Then
            '
            ' Call PlanningObjects.PlnAssembly.BuildMethod to get build method.
            ' PlanningObjects is in Middle, while this probject is in Content which
            ' is built after PlanningObjects during Smart 3D internal build process.
              
              Dim oPlnAssemblyHelper As PlanningObjects.PlnAssembly
              Set oPlnAssemblyHelper = New PlanningObjects.PlnAssembly
              
              Set oPlnAssemblyHelper.object = oAssembly
              
              Dim dAssyAngle As Double
              dAssyAngle = oPlnAssemblyHelper.SlotOpenAngle(oSlotWrapper.object)
              
              Set oAssembly = Nothing
              Set oPlnAssemblyHelper = Nothing

              If GreaterThan(dAssyAngle, ONE_DEGREE_AS_RADIANT) And LessThan(dAssyAngle, 1.570079633) Then
                  dSlotAngle = dAssyAngle
              Else
                  dSlotAngle = ONE_DEGREE_AS_RADIANT / 10
              End If
          End If
          Set oSlotWrapper = Nothing
          
  End Select
  
  GetSlotAngle = dSlotAngle
  
  Exit Function
  
ErrorHandler:
    Err.Raise LogError(Err, "SlotRulesCommon", "GetSlotAngle").Number
End Function
