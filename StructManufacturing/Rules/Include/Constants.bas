Attribute VB_Name = "Constants"
Public Const COMPANY_ID As String = "TSN"

'*** CONSTANTS FOR PLATE/PROFILE DECLIVITY MARKS ***'
Public Const DECLIVITY_MARKING_LINE_LENGTH As Double = 0.125
Public Const DECLIVITY_CONNECTED_LENGTH As Double = 3#
Public Const DECLIVITY_CONN_END_OFFSET_ABS As Double = 0.3
Public Const DECLIVITY_CONN_END_OFFSET_REL As Double = 0.1 ' 10 %
Public Const DECLIVITY_OFFSET_FROM_LOCATION As Double = 0.035
Public Const DECLIVITY_MARK_OBTUSE_ANGLE As Boolean = False
Public Const DECLIVITY_MARK_MINIMUM_TEE_LENGTH As Double = 0.2 ' in meters
Public Const DECLIVITY_SHOW_TOLERANCE As Double = 0.1 ' in degrees
Public Const DECLIVITY_TWIST_TOLERANCE As Double = 0.5 ' in degrees
' Below constant controls whether Declivity mark "points towards" location
' mark (SKDY) or if one of its arms is parallel to location mark (TSN)
Public Const DECLIVITY_POINTS_TO_LOCATION As Boolean = True ' False for TSN
' Below constant used in SKDY rules only.  For TSN, it
' is implicitly driven by DECLIVITY_MARKING_LINE_LENGTH
Public Const DECLIVITY_OFFSET_FROM_FRAME As Double = 0.01

Public Const SPLIT_KNUCKLE_OFFSET_FROM_EDGE As Double = 0.04
Public Const SPLIT_KNUCKLE_MARKING_LINE_LENGTH As Double = 0.15

Public Const BEVEL_MARK_LONGI_LENGTH As Double = 0.1
Public Const BEVEL_MARK_TRANS_LENGTH As Double = 0.075

'*** CONSTANTS FOR PLATE/PROFILE LOCATION MARKING LINES ***'
Public Const LOCATION_MARKING_LENGTH As Double = 5#
Public Const LOCATION_MARKING_SEG_MAX As Double = 0.16
Public Const LOCATION_MARKING_SEG_MIN As Double = 0.075
Public Const THICK_BUBBLE_RADIUS As Double = 0.047
Public Const THIN_BUBBLE_RADIUS As Double = 0.02
Public Const THICK_ARC_RADIUS As Double = 0.04
Public Const THIN_ARC_RADIUS As Double = 0.0175
Public Const OFFSET_DISTANCE As Double = 0.9
Public Const OFFSET_CONDITION As Double = 0.01
'**********************************************************'

Public Const FRAME_MARKING_DELIMITER As String = ","


'*** CONSTANTS FOR END FITTING MARKS ***'
Public Const END_FITTING_MARK_LENGTH As Double = 0.015
Public Const BRACKETSTIFF_FITTING_MARK_LENGTH As Double = 0.03
'***************************************'

'*** CONSTANTS FOR PROFILE TO PLATE PENETRATION MARK ***'
Public Const PROFILE_TO_PLATE_PEN_MARK_LENGTH As Double = 0.05
Public Const PROFILE_TO_PLATE_PEN_STRETCH_LENGTH As Double = 0.001

'***************************************'

'*** CONSTANTS FOR END OF FACE PLATE FITTING MARKS ***'
Public Const FACE_PLATE_BUBBLE_LENGTH As Double = 0#
'***************************************'

Public Const IDS_PROFMARKING_RULE_MSG_PROJN_FAILED = 1

' Constants for margin types, these value match codelist value
Public Const MARGIN_TYPE_SHIAGE As Long = 10001
Public Const MARGIN_TYPE_ARA    As Long = 10002

' Fitting mark length for REND Marks
Public Const REND_FIT_MARKING_LINE_LENGTH As Double = 0.05

' Ship direction mark length
Public Const SHIP_DIRECTION_MARKING_LENGTH_DEFAULT As Double = 0.07
Public Const SHIP_DIRECTION_MARKING_LENGTH_UPSIDE As Double = 0.005

' Collar mark length
Public Const COLLAR_LOCATION_MARKING_LENGTH_DEFAULT As Double = 0.005

'HoleMark Length
Public Const HOLE_MINIMUM_DIAMETER As Double = 0.035

'Constants for End Connection mark
Public Const LEFT_EXTENDED_LENGTH = 0.05
Public Const RIGHT_EXTENDED_LENGTH = 0.1
Public Const ENDCONN_EXTENSION = 0.001

' Offset Edge Mark
Public Const OFFSET_MARK_PREFIX As String = "L2"

' Connected Part Condition for Location Mark Label
'Public Const CONN_PART_CONDITION As Long = 0

' Profile Flange Location Mark
Public Const FLANGE_MARK_PREFIX As String = "-"

Public Const TEMPLATE_BCL_MARK_LENGTH = 0.05
Public Const TEMPLATE_SEAM_MARK_LENGTH = 0.05
Public Const TEMPLATE_REF_MARK_LENGTH = 0.05
Public Const TEMPLATE_KNUCKLE_MARK_LENGTH = 0.05
Public Const TEMPLATE_SIGHT_LINE_OFFSET = 0.05
Public Const TEMPLATE_SHIP_DIR_PRIMARY_LENGTH = 0.2
Public Const TEMPLATE_SHIP_DIR_SECONDARY_LENGTH = 0.1
Public Const PR_TEMPLATE_SHIP_DIR_PRIMARY_LENGTH = 0.1
Public Const PR_TEMPLATE_SHIP_DIR_SECONDARY_LENGTH = 0.05
Public Const TEMPLATE_FITTING_MARK_OFFSET = 0.3

' Shell Plate Roll Line Mark
Public Const SHELL_PLATE_ROLL_LINE_LENGTH As Double = 2.5   'Units in Meters
Public Const SHELL_PLATE_ROLL_LINE_SHORTLENGTH As Double = 1   'Units in Meters


Public Const MARGIN_VALUE As Double = 0.01
Public Const DEGREES_PER_RADIAN = 57.2957795130823
Public Const PI = 3.14159265358979
