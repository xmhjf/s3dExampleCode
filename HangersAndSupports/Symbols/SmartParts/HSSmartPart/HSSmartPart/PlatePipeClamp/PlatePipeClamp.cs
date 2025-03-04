//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PipeClamp.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PlatePipeClamp
//   Author       :  Chethan
//   Creation Date:  02-04-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   02-04-2015   Chethan   CR-CP-270224  Create and Add PlatePipeClamp SmartPart (.NET)  
//   10-06-2015   PVK	    TR-CP-274155	SmartPart TDL Errors should be corrected.
//   16-07-2015   PVK       Resolve coverity issues found in July 2015 report
//   30-11=2015   VDP       Integrate the newly developed SmartParts into Product(DI-CP-282644)
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class PlatePipeClamp : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PlatePipeClamp"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        
        // Basic inputs
        [InputDouble(2, "PipeOD", "PipeOD", 0)]        
        public InputDouble m_PipeOD;
        [InputDouble(3, "TrunnionDiameter", "TrunnionDiameter", 0)]
        public InputDouble m_Trunniondiameter;
        
        //Pipe clamp aplte inputs
        [InputString(4, "ClampPlateShape", "ClampPlateShape", "No Value")]
        public InputString m_sClampPlateShape;
        [InputDouble(5, "Offset1X", "Offset1X", 0)]
        public InputDouble m_dOffset1X;
        [InputDouble(6, "Offset1Y", "Offset1Y", 0)]
        public InputDouble m_dOffset1Y;
        [InputDouble(7, "Offset1Z", "Offset1Z", 0)]
        public InputDouble m_dOffset1Z;

        //Vertical Plate inputs
        [InputString(8, "VerticalPlateShape", "VerticalPlateShape", "No Value")]
        public InputString m_sVerticalPlateShape;
        [InputDouble(9, "VerticalPlateGap", "VerticalPlateGap", 0)]
        public InputDouble m_dVerticalPlateGap;
        [InputDouble(10, "VerticalPlateLength", "VerticalPlateLength", 0)]
        public InputDouble m_dVerticalPlateLength;
        [InputDouble(11, "VerticalPlateHeight", "VerticalPlateHeight", 0)]
        public InputDouble m_dVerticalPlateHeight;

        //Middle Plate inputs
        [InputString(12, "MiddlePlateShape", "MiddlePlateShape", "No Value")]
        public InputString m_sMiddlePlateShape;
        [InputDouble(13, "MiddlePlateGap", "MiddlePlateGap", 0)]
        public InputDouble m_dMiddlePlateGap;
        [InputDouble(14, "Offset2X", "Offset2X", 0)]
        public InputDouble m_dOffset2X;
        [InputDouble(15, "Offset2Y", "Offset2Y", 0)]
        public InputDouble m_dOffset2Y;
        
        //End Plate inputs
        [InputString(16, "EndPlateShape", "EndPlateShape", "No Value")]
        public InputString m_sEndPlateShape;
        [InputDouble(17, "EndPlateLocateBy", "EndPlateLocateBy", 0)]
        public InputDouble m_dEndPlateLocateBy;
        [InputDouble(18, "EndPlateRotation", "EndPlateRotation", 0)]
        public InputDouble m_dEndPlateRotation;
        [InputDouble(19, "Offset3X", "Offset3X", 0)]
        public InputDouble m_dOffset3X;
        [InputDouble(20, "Offset3Y", "Offset3Y", 0)]
        public InputDouble m_dOffset3Y;
        [InputDouble(21, "Offset3Z", "Offset3Z", 0)]
        public InputDouble m_dOffset3Z;

        //Nut1 Inputs
        [InputString(22, "Nut1Shape", "Nut1Shape", "No Value")]
        public InputString m_sNut1Shape;
        [InputDouble(23, "Offset4X", "Offset4X", 0)]
        public InputDouble m_dOffset4X;
        [InputDouble(24, "Offset4Z", "Offset4Z", 0)]
        public InputDouble m_dOffset4Z;
        [InputDouble(25, "Nut1Quantity", "Nut1Quantity", 0)]
        public InputDouble m_dNut1Quantity;
        [InputDouble(26, "Nut1Spacing", "Nut1Spacing", 0)]
        public InputDouble m_dNut1Spacing;

        //Nut2 Inputs
        [InputString(27, "Nut2Shape", "Nut2Shape", "No Value")]
        public InputString m_sNut2Shape;
        [InputDouble(28, "Offset5X", "Offset5X", 0)]
        public InputDouble m_dOffset5X;
        [InputDouble(29, "Offset5Z", "Offset5Z", 0)]
        public InputDouble m_dOffset5Z;
        [InputDouble(30, "Nut2Quantity", "Nut2Quantity", 0)]
        public InputDouble m_dNut2Quantity;
        [InputDouble(31, "Nut2Spacing", "Nut2Spacing", 0)]
        public InputDouble m_dNut2Spacing;

        //Nut3 Inputs
        [InputString(32, "Nut3Shape", "Nut3Shape", "No Value")]
        public InputString m_sNut3Shape;
        [InputDouble(33, "Offset6X", "Offset6X", 0)]
        public InputDouble m_dOffset6X;
        [InputDouble(34, "Offset6Z", "Offset6Z", 0)]
        public InputDouble m_dOffset6Z;
        [InputDouble(35, "Nut3Quantity", "Nut3Quantity", 0)]
        public InputDouble m_dNut3Quantity;
        [InputDouble(36, "Nut3Spacing", "Nut3Spacing", 0)]
        public InputDouble m_dNut3Spacing;

        //Hole Ports offset
        [InputDouble(37, "Offset7X", "Offset7X", 0)]
        public InputDouble m_dOffset7X;
        [InputDouble(38, "Offset8X", "Offset8X", 0)]
        public InputDouble m_dOffset8X;

        //Miscellaneous
        [InputDouble(39, "CCValue", "CCValue", 0)]
        public InputDouble m_dCCValue;
        [InputDouble(40, "CCMinimumValue", "CCMinimumValue", 0)]
        public InputDouble m_dCCMinimumValue;
        [InputDouble(41, "CCMaximumValue", "CCMaximumValue", 0)]
        public InputDouble m_dCCMaximumValue;
        [InputDouble(42, "MinimumHeight", "MinimumHeight", 0)]
        public InputDouble m_dMinimumHeight;
        [InputDouble(43, "MaximumHeight", "MaximumHeight", 0)]
        public InputDouble m_dMaximumHeight;

        #endregion "Definition of Inputs"

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
        [SymbolOutput("ClampPlate", "ClampPlate")]
        [SymbolOutput("VerticalPlate1", "VerticalPlate1")]
        [SymbolOutput("VerticalPlate2", "VerticalPlate2")]
        [SymbolOutput("MiddlePlate1", "MiddlePlate1")]
        [SymbolOutput("MiddlePlate2", "MiddlePlate2")]
        [SymbolOutput("EndPlate1", "EndPlate1")]
        [SymbolOutput("EndPlate2", "EndPlate2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construct Outputs"
        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {            
            try
            {
                Part part = (Part)m_PartInput.Value;


                PlateInputs VerticalplateshapeInputs , ClampPlateInputs ,MiddlePlateInputs, EndPlateInputs;
                NutInputs Nut1Inputs, Nut2Inputs, Nut3Inputs;

                Double pipeOD = m_PipeOD.Value;
                Double trunnionDiameter = m_Trunniondiameter.Value;

                String clampPlateShape = m_sClampPlateShape.Value;
                Double offset1X = m_dOffset1X.Value;
                Double offset1Y = m_dOffset1Y.Value;
                Double offset1Z = m_dOffset1Z.Value;

                String verticalPlateShape = m_sVerticalPlateShape.Value;
                Double verticalPlateGap = m_dVerticalPlateGap.Value;
                Double verticalPlateLength = m_dVerticalPlateLength.Value;
                Double verticalPlateHeight = m_dVerticalPlateHeight.Value;

                String middlePlateShape = m_sMiddlePlateShape.Value;
                Double middlePlateGap = m_dMiddlePlateGap.Value;
                Double offset2X = m_dOffset2X.Value;
                Double offset2Y = m_dOffset2Y.Value;

                String endPlateShape = m_sEndPlateShape.Value;
                int endPlateLocateBy = (int)m_dEndPlateLocateBy.Value;
                int endPlateRotation = (int)m_dEndPlateRotation.Value;
                Double offset3X = m_dOffset3X.Value;
                Double offset3Y = m_dOffset3Y.Value;
                Double offset3Z = m_dOffset3Z.Value;

                String nut1Shape = m_sNut1Shape.Value;
                Double offset4X = m_dOffset4X.Value;
                Double offset4Z = m_dOffset4Z.Value;
                int nut1Quantity = (int)m_dNut1Quantity.Value;
                Double nut1Spacing = m_dNut1Spacing.Value;

                String nut2Shape = m_sNut2Shape.Value;
                Double offset5X = m_dOffset5X.Value;
                Double offset5Z = m_dOffset5Z.Value;
                int nut2Quantity = (int)m_dNut2Quantity.Value;
                Double nut2Spacing = m_dNut2Spacing.Value;

                String nut3Shape = m_sNut3Shape.Value;
                Double offset6X = m_dOffset6X.Value;
                Double offset6Z = m_dOffset6Z.Value;
                int nut3Quantity = (int)m_dNut3Quantity.Value;
                Double nut3Spacing = m_dNut3Spacing.Value;

                Double offset7X =  m_dOffset7X.Value;
                Double offset8X = m_dOffset8X.Value;

                Double CCValue = m_dCCValue.Value;
                Double CCMinValue = m_dCCMinimumValue.Value;
                Double CCMaxValue = m_dCCMaximumValue.Value;
                Double minHeight = m_dMinimumHeight.Value;
                Double maxHeight = m_dMaximumHeight.Value;
                
                //Error Checking
                
                if(pipeOD<=0)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrInvalidPipeOD,"Pipe Outer Diameter is required.");
                }
                if(verticalPlateShape == "" || verticalPlateShape=="No Value")
                {
                    ToDoListMessage=new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrVerticalPlateShapeRequired,"Vertical Plate Shape Data is required."));
                }
                if(verticalPlateHeight<minHeight)
                {
                    ToDoListMessage=new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidMinVerticalPLateHeight, "Total height of vertical plate is less than the minimum allowable height."));
                }
                if(verticalPlateHeight>maxHeight)
                {
                    ToDoListMessage=new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning,SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidMaxVerticalPLateHeight,"Total height of vertical plate is greater than the maximum allowable height."));
                }
                


                //Initializing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                 //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Route"] = port1;

                //Add Vertical PLates

                 VerticalplateshapeInputs= LoadHolePlateDataByQuery(verticalPlateShape);

                //Override Shape data

                if (trunnionDiameter>0)
                {
                    VerticalplateshapeInputs.Hole1Diameter=trunnionDiameter;
                }
                if(verticalPlateLength>0)
                {
                    VerticalplateshapeInputs.length1=verticalPlateLength;
                }
                if(verticalPlateHeight>0)
                {
                    VerticalplateshapeInputs.width1=verticalPlateHeight;
                }
                
                Matrix4X4 matrix =new Matrix4X4();
                matrix.Rotate(Math.PI/2,new Vector(1,0,0));
                matrix.Translate(new Vector(-VerticalplateshapeInputs.width1 / 2,(verticalPlateGap / 2) + VerticalplateshapeInputs.thickness1,-VerticalplateshapeInputs.length1/2));
                AddPlateShapeWithHole(VerticalplateshapeInputs,matrix,m_PhysicalAspect.Outputs,"VerticalPlate1");
                
                matrix =new Matrix4X4();
                matrix.Rotate(Math.PI/2,new Vector(1,0,0));
                matrix.Translate(new Vector(-VerticalplateshapeInputs.width1 / 2,-(verticalPlateGap / 2) ,-VerticalplateshapeInputs.length1/2));
                AddPlateShapeWithHole(VerticalplateshapeInputs,matrix,m_PhysicalAspect.Outputs,"VerticalPlate2");

                if (HgrCompareDoubleService.cmpdbl(endPlateLocateBy,2))
                {
                    offset3Z=offset3Z+VerticalplateshapeInputs.length1/2;
                
                }
                
                //Error Checking

                if (Math.Round((2*offset3Z), 6) < Math.Round(CCMinValue,6))
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrIvalidminCCValue, "The C-C Value is less than the minimum allowable C-C Value.");
                }
                if(Math.Round((2*offset3Z),6) > Math.Round(CCMaxValue,6))
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrIvalidmaxCCValue, "The C-C Value is greater than the maximum allowable C-C Value.");
                }

                //Add Clamp Plate
                if(clampPlateShape !="" && clampPlateShape!="No Value")
                {
                    ClampPlateInputs = LoadHolePlateDataByQuery(clampPlateShape);

                    if(pipeOD>0)
                    {
                        ClampPlateInputs.Hole1Diameter = pipeOD;                        
                    }
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI/2,new Vector(0,1,0));                    
                    matrix.Translate(new Vector(offset1X+ (VerticalplateshapeInputs.width1 /2), offset1Z-(ClampPlateInputs.length1 /2), -(offset1Y-(ClampPlateInputs.width1 /2))));                    
                    AddPlateShapeWithHole(ClampPlateInputs, matrix, m_PhysicalAspect.Outputs, "ClampPlate");
                }

                //Add Middle Plates
                if (middlePlateShape != "" && middlePlateShape != "No Value")
                {
                    MiddlePlateInputs = LoadHolePlateDataByQuery(middlePlateShape);

                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(offset2X - (MiddlePlateInputs.width1 / 2), offset2Y - (MiddlePlateInputs.length1 / 2), (middlePlateGap / 2)));
                    AddPlate(MiddlePlateInputs,matrix,m_PhysicalAspect.Outputs,"MiddlePlate1");

                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(offset2X - (MiddlePlateInputs.width1 / 2), offset2Y - (MiddlePlateInputs.length1 / 2), -((middlePlateGap / 2) + MiddlePlateInputs.thickness1)));
                    AddPlate(MiddlePlateInputs, matrix, m_PhysicalAspect.Outputs, "MiddlePlate2");
                }

                //Add End Plates

                EndPlateInputs = LoadHolePlateDataByQuery(endPlateShape);

                if (endPlateShape != "" && endPlateShape != "No Value")
                {                    
                    if(endPlateRotation==1)
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI/2,new Vector(1,0,0));
                        matrix.Translate(new Vector(offset3X - (EndPlateInputs.width1 / 2), offset3Y + (EndPlateInputs.thickness1 / 2), offset3Z - (EndPlateInputs.length1 / 2)));
                        AddPlate(EndPlateInputs,matrix,m_PhysicalAspect.Outputs,"EndPlate1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset3X - (EndPlateInputs.width1 / 2), offset3Y + (EndPlateInputs.thickness1 / 2), -offset3Z - (EndPlateInputs.length1 / 2)));
                        AddPlate(EndPlateInputs, matrix, m_PhysicalAspect.Outputs, "EndPlate2");
                    }
                    else
                    {
                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(offset3X - (EndPlateInputs.width1 / 2), offset3Y - (EndPlateInputs.length1 / 2), offset3Z - (EndPlateInputs.thickness1 / 2)));
                        AddPlate(EndPlateInputs, matrix, m_PhysicalAspect.Outputs, "EndPlate1");

                        matrix = new Matrix4X4();
                        matrix.Translate(new Vector(offset3X - (EndPlateInputs.width1 / 2), offset3Y - (EndPlateInputs.length1 / 2), -offset3Z - (EndPlateInputs.thickness1 / 2)));
                        AddPlate(EndPlateInputs, matrix, m_PhysicalAspect.Outputs, "EndPlate2");
                    }
                }

                //Add Input1
                if (nut1Shape != "" && nut1Shape != "No Value")
                {
                    Nut1Inputs = LoadNutDataByQuery(nut1Shape, 1);

                    if (HgrCompareDoubleService.cmpdbl(nut1Quantity, 1))
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X, Nut1Inputs.ShapeLength / 2, offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X, Nut1Inputs.ShapeLength / 2, -offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"2");
                    }
                    else if (HgrCompareDoubleService.cmpdbl(nut1Quantity, 2))
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X+(nut1Spacing/2), Nut1Inputs.ShapeLength / 2, offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X - (nut1Spacing / 2), Nut1Inputs.ShapeLength / 2, offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"2");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X + (nut1Spacing / 2), Nut1Inputs.ShapeLength / 2, -offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"3");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X - (nut1Spacing / 2), Nut1Inputs.ShapeLength / 2, -offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"4");
                    }
                    else if (HgrCompareDoubleService.cmpdbl(nut1Quantity, 3))
                    {
                        //Left Side
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X + (nut1Spacing / 2), Nut1Inputs.ShapeLength / 2, offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X, Nut1Inputs.ShapeLength / 2, offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"2");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X + (nut1Spacing / 2), Nut1Inputs.ShapeLength / 2, offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"3");

                        //Right Side
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X + (nut1Spacing / 2), Nut1Inputs.ShapeLength / 2, -offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"4");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X, Nut1Inputs.ShapeLength / 2, -offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"5");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset4X + (nut1Spacing / 2), Nut1Inputs.ShapeLength / 2, -offset4Z));
                        AddNut(Nut1Inputs, matrix, m_PhysicalAspect.Outputs, "Nut1"+"6");
                    }
                }
                //Add nut2
                if (nut2Shape != "" && nut2Shape != "No Value")
                {
                    Nut2Inputs = LoadNutDataByQuery(nut2Shape, 1);

                    if (HgrCompareDoubleService.cmpdbl(nut2Quantity, 1))
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X, Nut2Inputs.ShapeLength / 2, offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X, Nut2Inputs.ShapeLength / 2, -offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"2");
                    }
                    else if (HgrCompareDoubleService.cmpdbl(nut2Quantity, 2))
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X + (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"3");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X - (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"4");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X + (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, -offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"5");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X - (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, -offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"6");
                    }
                    else if (HgrCompareDoubleService.cmpdbl(nut2Quantity, 3))
                    {
                        //Left Side
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X + (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X, Nut2Inputs.ShapeLength / 2, offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"2");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X + (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"3");

                        //Right Side
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X + (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, -offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"4");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X, Nut2Inputs.ShapeLength / 2, -offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"5");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X + (nut2Spacing / 2), Nut2Inputs.ShapeLength / 2, -offset5Z));
                        AddNut(Nut2Inputs, matrix, m_PhysicalAspect.Outputs, "Nut2"+"6");
                    }
                }
                //Add Nut3
                if (nut3Shape != "" && nut3Shape != "No Value")
                {
                    Nut3Inputs = LoadNutDataByQuery(nut3Shape, 1);

                    if (nut3Quantity == 1)
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X, Nut3Inputs.ShapeLength / 2, offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X, Nut3Inputs.ShapeLength / 2, -offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"2");
                    }
                    else if (nut3Quantity == 2)
                    {
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X + (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X - (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"2");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X + (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, -offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"3");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset5X - (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, -offset5Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"4");
                    }
                    else if (HgrCompareDoubleService.cmpdbl(nut3Quantity, 3))
                    {
                        //Left Side
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X + (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"1");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X, Nut3Inputs.ShapeLength / 2, offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"2");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X + (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"3");

                        //Right Side
                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X + (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, -offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"4");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X, Nut3Inputs.ShapeLength / 2, -offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"5");

                        matrix = new Matrix4X4();
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        matrix.Translate(new Vector(offset6X + (nut3Spacing / 2), Nut3Inputs.ShapeLength / 2, -offset6Z));
                        AddNut(Nut3Inputs, matrix, m_PhysicalAspect.Outputs, "Nut3"+"6");
                    }
                }
                //Add Hole Ports

                if ((endPlateShape != "" && endPlateShape != "No Value") & (endPlateRotation == 1))
                {
                    Port Hole1 = new Port(OccurrenceConnection, part, "Hole1", new Position(offset7X + offset3X + (EndPlateInputs.width1 / 2), 0, offset3Z), new Vector(0, 0, 1), new Vector(1, 0, 0));
                    m_PhysicalAspect.Outputs["Hole1"] = Hole1;

                    Port Hole2 = new Port(OccurrenceConnection, part, "Hole2", new Position(offset8X + offset3X + (EndPlateInputs.width1 / 2), 0, -offset3Z), new Vector(0, 0, -1), new Vector(1, 0, 0));
                    m_PhysicalAspect.Outputs["Hole2"] = Hole2;                   
                }
                else if ((endPlateShape != "" && endPlateShape != "No Value") & (endPlateRotation == 2))
                {
                    Port Hole1 = new Port(OccurrenceConnection, part, "Hole1", new Position(offset7X + offset3X + (EndPlateInputs.width1 / 2), 0, offset3Z), new Vector(0, 1, 0), new Vector(1, 0, 0));
                    m_PhysicalAspect.Outputs["Hole1"] = Hole1;

                    Port Hole2 = new Port(OccurrenceConnection, part, "Hole2", new Position(offset8X + offset3X + (EndPlateInputs.width1 / 2), 0, -offset3Z), new Vector(0, -1, 0), new Vector(1, 0, 0));
                    m_PhysicalAspect.Outputs["Hole2"] = Hole2;
                }
                else
                {
                    Port Hole1 = new Port(OccurrenceConnection, part, "Hole1", new Position(offset7X  , 0, VerticalplateshapeInputs.width1/2), new Vector(0, 1, 0), new Vector(1, 0, 0));
                    m_PhysicalAspect.Outputs["Hole1"] = Hole1;

                    Port Hole2 = new Port(OccurrenceConnection, part, "Hole2", new Position(offset8X , 0, -VerticalplateshapeInputs.width1/2), new Vector(0, -1, 0), new Vector(1, 0, 0));
                    m_PhysicalAspect.Outputs["Hole2"] = Hole2;
                }

                CCValue = 2 * offset3Z;
            }
            catch (SmartPartSymbolException hgrEx)
            {
                throw hgrEx;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PlatePipeClamp"));
                }
            }           
        }

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {

            string materialType, materialGrade;
            SupportComponent supportComponent = (SupportComponent)supportComponentBO;
            VolumeCG volumeCG = supportComponent.GetVolumeAndCOG();

            Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
            CatalogStructHelper catalogStructHelper = new CatalogStructHelper();

            if (supportComponentBO.SupportsInterface("IJOAhsMaterialEx"))
            {
                try
                {
                    materialType = (((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue);
                    materialGrade = (((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue);
                }
                catch
                {
                    materialType = String.Empty;
                    materialGrade = String.Empty;
                }

            }
            else if (catalogPart.SupportsInterface("IJOAhsMaterialEx"))
            {
                try
                {
                    materialType = (((PropertyValueString)catalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue);
                    materialGrade = (((PropertyValueString)catalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue);
                }
                catch
                {
                    materialType = String.Empty;
                    materialGrade = String.Empty;
                }

            }
            else
            {
                materialType = String.Empty;
                materialGrade = String.Empty;
            }


            Material material;
            double materialDensity;
            try
            {
                material = catalogStructHelper.GetMaterial(materialType, materialGrade);
                materialDensity = material.Density;
            }
            catch
            {
                // the specified MaterialType is not available.refdata needs to be checked.
                // so assigning 0 to materialDensity.
                materialDensity = 0;
            }


            double weight, cogx, cogy, cogz;
             try
             {
                 weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
             }
             catch
             {
                 weight = volumeCG.Volume * materialDensity;
             }

             try
             {
                 cogx = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
             }
             catch
             {
                 cogx = volumeCG.COGX;
             }

             try
             {
                 cogy = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
             }
             catch
             {
                 cogy = volumeCG.COGY;
             }

             try
             {
                 cogz = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
             }
             catch
             {
                 cogz = volumeCG.COGZ;
             }
             supportComponent.SetWeightAndCOG(weight, cogx, cogy, cogz);
        }
        #endregion
    }
}
