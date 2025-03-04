Attribute VB_Name = "CodelistKeys"
Option Explicit

'Codelist tables
Public Const gsBooleanCol = "Boolean"
Public Const gsShapeAtEdgeCol = "ShapeAtEdgeCol"
Public Const gsShapeAtFaceCol = "ShapeAtFaceCol"
Public Const gsShapeAtEdgeOverlapCol = "ShapeAtEdgeOverlapCol"
Public Const gsShapeAtEdgeOutsideCol = "ShapeAtEdgeOutsideCol"
Public Const gsConnectionTypeCol = "ConnectionTypeCol"
Public Const gsExtendOrOffsetCol = "ExtendOrOffsetCol"
Public Const gsShapeOutsideCol = "ShapeOutsideCol"
Public Const gsInsideCornerCol = "InsideCornerCol"
Public Const gsEndCutShapeCol = "EndCutShapeCol"
'25/Aug/2011   - Addedd new Codelist to create InsetBrace
Public Const gsBraceTypeCol = "BraceTypeCol"
Public Const gsBraceConnTypeCol = "BraceConnTypeCol"
Public Const gsChamferMeasurementCol = "ChamferMeasurementCodeList"

'Commonly used codelist values
Public Const gsNone = "None"
Public Const gsYes = "Yes"
Public Const gsNo = "No"

'ShapeAtEdgeCol codelist options
Public Const gsFaceToCorner = "FaceToCorner"
Public Const gsFaceToEdge = "FaceToEdge"
Public Const gsFaceToFlange = "FaceToFlange"
Public Const gsFaceToOutside = "FaceToOutside"
Public Const gsInsideToEdge = "InsideToEdge"
Public Const gsInsideToFlange = "InsideToFlange"
Public Const gsInsideToOutside = "InsideToOutside"
Public Const gsCornerToFlange = "CornerToFlange"
Public Const gsCornerToOutside = "CornerToOutside"
Public Const gsEdgeToFlange = "EdgeToFlange"
Public Const gsEdgeToOutside = "EdgeToOutside"
Public Const gsOutsideToEdge = "OutsideToEdge"
Public Const gsOutsideToFlange = "OutsideToFlange"
Public Const gsOutsideToOutside = "OutsideToOutside"

'ShapeAtEdgeOverlapCol codelist options
Public Const gsFaceToInsideCorner = "FaceToInsideCorner"
'Public Const gsFaceToEdge = "FaceToEdge"
Public Const gsFaceToOutsideCorner = "FaceToOutsideCorner"
'Public Const gsFaceToOutside = "FaceToOutside"
'Public Const gsInsideToEdge = "InsideToEdge"
Public Const gsInsideToOutsideCorner = "InsideToOutsideCorner"
'Public Const gsInsideToOutside = "InsideToOutside"
Public Const gsInsideCornerToOutside = "InsideCornerToOutside"
'Public Const gsEdgeToOutside = "EdgeToOutside"

'ShapeAtFace codelist
Public Const gsCope = "Cope"
Public Const gsInside = "Inside"

'ExtendOrOffsetCol codelist options
Public Const gsExtendFarCorner = "ExtendFarCorner"
Public Const gsOffsetFarCorner = "OffsetFarCorner"
Public Const gsExtendFarSide = "ExtendFarSide"
Public Const gsOffsetFarSide = "OffsetFarSide"
Public Const gsExtendNearSide = "ExtendNearSide"
Public Const gsOffsetNearSide = "OffsetNearSide"
Public Const gsExtendNearCorner = "ExtendNearCorner"
Public Const gsOffsetNearCorner = "OffsetNearCorner"

'ShapeOutsideCol And EndCutShapeCol codelist options
Public Const gsStraight = "Straight"
Public Const gsSniped = "Sniped"
  '----only exisits for EndCutShapeCol
Public Const gsNotApplied = "NotApplied"

'InsideCornerCol codelist options
Public Const gsSnipe = "Snipe"
Public Const gsScallop = "Scallop"
Public Const gsFillet = "Fillet"

'ChamferMeasurement Codelist
Public Const gsSlope = "Slope"
Public Const gsAngle = "Angle"
