//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   BeamClamp.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.BeamClamp
//   Author       :  Vijay
//   Creation Date: 06.Feb.2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   06.Feb.2013     Vijay    CR-CP-222474 Initial Creation 
//	 25.Mar.2013     Vijay 	  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   11/11/2013     Rajeswari    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   12-12-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   10-06-2015      PVK	  TR-CP-274155	SmartPart TDL Errors should be corrected.
//   16-07-2015      PVK       Resolve coverity issues found in July 2015 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class BeamClamp : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.BeamClamp"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                //Beam Clamp Inputs
                AddBeamClampInputs(2, out endIndex, additionalInputs);

                //Botttom Dimensions
                additionalInputs.Add(new InputDouble(++endIndex, "Width1", "Width1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Thickness1", "Thickness1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Height1", "Height1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Height2", "Height2", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Length1", "Length1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Length2", "Length2", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "RodTakeOut", "RodTakeOut", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Angle1", "Angle1", 0, false));

                //Bottom Shape Bolt Dimensions
                additionalInputs.Add(new InputDouble(++endIndex, "Pin1Diameter", "Pin1Diameter", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Pin1Length", "Pin1Length", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Height3", "Height3", 0, false));

                //Left and Right Bolt Dimensions
                AddBoltInputs(endIndex, 2, out endIndex, additionalInputs);

                //UBolt/JBolt Dimensions
                additionalInputs.Add(new InputDouble(++endIndex, "UBoltWidth", "UBoltWidth", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "UBoltLength", "UBoltLength", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "UBoltRodDia", "UBoltRodDia", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "UBoltFlatSpot", "UBoltFlatSpot", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "IsJBolt", "IsJBolt", 1, false));

                //Struct Dimensions
                additionalInputs.Add(new InputDouble(++endIndex, "ClampThickness", "ClampThickness", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "StructWidth", "StructWidth", 0, false));
                return additionalInputs;
            }
        }
        #endregion
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("Top", "Top")]
        [SymbolOutput("Bottom", "Bottom")]
        [SymbolOutput("RodEnd1", "RodEnd1")]
        [SymbolOutput("RodEnd2", "RodEnd2")]
        [SymbolOutput("RodEnd", "RodEnd")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddBeamClampOutputs(2, additionalOutputs);
            }
            return additionalOutputs;
        }
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
                int endIndex, startIndex;
                Matrix4X4 matrix = new Matrix4X4();
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //Load Beam Clamp Data
                BeamClampInputs beamClamp = LoadBeamClampData(2, out endIndex);
                startIndex = endIndex;

                //Load Bottom Shape Data
                Double width1 = GetDoubleInputValue(++startIndex);
                Double thickness1 = GetDoubleInputValue(++startIndex);
                Double height1 = GetDoubleInputValue(++startIndex);
                Double height2 = GetDoubleInputValue(++startIndex);
                Double length1 = GetDoubleInputValue(++startIndex);
                Double length2 = GetDoubleInputValue(++startIndex);
                Double rodTakeOut = GetDoubleInputValue(++startIndex);
                Double angle1 = GetDoubleInputValue(++startIndex);

                //Load Pin Bolt Data
                Double pin1Diameter = GetDoubleInputValue(++startIndex);
                Double pin1Length = GetDoubleInputValue(++startIndex);
                Double height3 = GetDoubleInputValue(++startIndex);

                //Load Left and Right Bolt Data
                BoltInputs leftBolt = LoadBoltData(++startIndex, out endIndex);
                startIndex = endIndex;
                BoltInputs rightBolt = LoadBoltData(++startIndex, out endIndex);
                startIndex = endIndex;

                //Load U Bolt/J Bolt Dimensions
                Double uBoltWidth = GetDoubleInputValue(++startIndex);
                Double uBoltLength = GetDoubleInputValue(++startIndex);
                Double uBoltRodDia = GetDoubleInputValue(++startIndex);
                Double uBoltFlatSpot = GetDoubleInputValue(++startIndex);
                int isJBolt = (int)GetDoubleInputValue(++startIndex);

                //Load Struct Attributes
                Double flangeThickness = GetDoubleInputValue(++startIndex);
                Double flangeWidth = GetDoubleInputValue(++startIndex);
                Double vertOffset1 = 0;
                Double vertOffset2 = 0;
                BeamClipInputs leftBeamClip = new BeamClipInputs();
                BeamClipInputs rightBeamClip = new BeamClipInputs();

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                //Load BeamClip Data By Query
                if (beamClamp.LeftClipShape != "")
                {
                    leftBeamClip = LoadBeamClipDataByQuery(beamClamp.LeftClipShape);
                }
                if (beamClamp.RightClipShape != "")
                {
                    rightBeamClip = LoadBeamClipDataByQuery(beamClamp.RightClipShape);
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Structure"] = port1;

                if ((beamClamp.LeftClipShape != "") && (beamClamp.LeftClipShape != "No Value"))
                {
                    if ((leftBeamClip.ShapeType == 5) || (leftBeamClip.ShapeType == 6))         //For P and Q
                    {
                        if (HgrCompareDoubleService.cmpdbl(leftBeamClip.ClipHeight1, 0) == true)
                            leftBeamClip.ClipHeight1 = flangeThickness + 2 * leftBeamClip.ClipThickness;
                        if (leftBeamClip.ShapeType == 5)
                        {
                            if (HgrCompareDoubleService.cmpdbl(leftBeamClip.ClipHeight2 , 0) == true)
                                leftBeamClip.ClipHeight2 = height3 + 2 * leftBeamClip.ClipThickness;
                        }
                    }
                    if ((HgrCompareDoubleService.cmpdbl(leftBeamClip.Bolt.BoltLength , 0) == false)  && (leftBeamClip.ShapeType != 7))
                    {
                        vertOffset2 = 0;
                        if (leftBeamClip.ShapeType == 1)
                            vertOffset1 = leftBeamClip.ClipHeight1 - leftBeamClip.ClipThickness1 - leftBeamClip.ClipThickness2 - flangeThickness;
                        else if ((leftBeamClip.ShapeType == 2) || (leftBeamClip.ShapeType == 3))
                            vertOffset1 = leftBeamClip.ClipHeight1 - leftBeamClip.ClipThickness1 - flangeThickness;
                        else if (leftBeamClip.ShapeType == 4)
                            vertOffset1 = leftBeamClip.ClipHeight1 - leftBeamClip.ClipThickness - leftBeamClip.ClipHeight2 - flangeThickness;
                        else if (leftBeamClip.ShapeType == 5)
                            vertOffset1 = leftBeamClip.ClipHeight1 - 2 * leftBeamClip.ClipThickness - flangeThickness;
                        else if (leftBeamClip.ShapeType == 6)
                            vertOffset1 = leftBeamClip.ClipHeight1 - 2 * leftBeamClip.ClipThickness - flangeThickness;
                    }
                    else
                    {
                        vertOffset1 = 0;
                        if (leftBeamClip.ShapeType == 1)
                            vertOffset2 = leftBeamClip.ClipHeight1 - leftBeamClip.ClipThickness1 - leftBeamClip.ClipThickness2 - flangeThickness;
                        else if ((leftBeamClip.ShapeType == 2) || (leftBeamClip.ShapeType == 3))
                            vertOffset2 = leftBeamClip.ClipHeight1 - leftBeamClip.ClipThickness1 - flangeThickness;
                        else if (leftBeamClip.ShapeType == 4)
                            vertOffset2 = leftBeamClip.ClipHeight1 - leftBeamClip.ClipThickness - leftBeamClip.ClipHeight2 - flangeThickness;
                        else if (leftBeamClip.ShapeType == 5)
                            vertOffset2 = leftBeamClip.ClipHeight1 - 2 * leftBeamClip.ClipThickness - flangeThickness;
                        else if (leftBeamClip.ShapeType == 6)
                            vertOffset2 = leftBeamClip.ClipHeight1 - 2 * leftBeamClip.ClipThickness - flangeThickness;
                    }

                    Port port2 = new Port(OccurrenceConnection, part, "Top", new Position(0, flangeWidth / 2, vertOffset1 + flangeThickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Top"] = port2;

                    if ((leftBeamClip.ShapeType == 8) && (uBoltWidth != 0))
                    {
                        Port port3 = new Port(OccurrenceConnection, part, "Bottom", new Position(0, flangeWidth / 2 + beamClamp.LeftClipGap + beamClamp.LeftOffsetH, beamClamp.LeftOffsetV - flangeThickness / 2 + uBoltRodDia), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Bottom"] = port3;
                    }
                    else
                    {
                        Port port3 = new Port(OccurrenceConnection, part, "Bottom", new Position(0, flangeWidth / 2, -vertOffset2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Bottom"] = port3;
                    }
                    if (!(leftBeamClip.ShapeType >= 5))
                    {
                        Port port4 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, flangeWidth / 2 + beamClamp.LeftClipGap + beamClamp.LeftOffsetH, beamClamp.LeftOffsetV), new Vector(1, 0, 0), new Vector(0, 0, -1));
                        m_PhysicalAspect.Outputs["RodEnd1"] = port4;
                    }
                    if ((leftBeamClip.ShapeType == 5) || (leftBeamClip.ShapeType == 6))
                    {
                        if (leftBeamClip.ShapeType == 5)
                        {
                            if ((HgrCompareDoubleService.cmpdbl(pin1Diameter, 0) == false) || (HgrCompareDoubleService.cmpdbl(pin1Length, 0) == false))
                            {
                                Port port4 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, 0, -height3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                                m_PhysicalAspect.Outputs["RodEnd1"] = port4;
                            }
                        }
                        if (leftBeamClip.ShapeType == 6)
                        {
                            if ((HgrCompareDoubleService.cmpdbl(pin1Diameter, 0) == false) || (HgrCompareDoubleService.cmpdbl(pin1Length, 0) == false))
                            {
                                Port port4 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, 0, -height3), new Vector(1, 0, 0), new Vector(0, 0, -1));
                                m_PhysicalAspect.Outputs["RodEnd1"] = port4;

                                Port port5 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, 0, -rodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                                m_PhysicalAspect.Outputs["RodEnd2"] = port5;
                            }
                            else
                            {
                                Port port4 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, 0, -rodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                                m_PhysicalAspect.Outputs["RodEnd1"] = port4;
                            }
                        }
                    }
                }

                if ((beamClamp.RightClipShape != "") && (beamClamp.RightClipShape != "No Value"))
                {
                    if ((beamClamp.LeftClipShape == "") || (beamClamp.LeftClipShape == "No Value"))
                    {
                        if ((rightBeamClip.ShapeType == 5) || (rightBeamClip.ShapeType == 6))
                        {
                            if (HgrCompareDoubleService.cmpdbl(rightBeamClip.ClipHeight1 , 0) == true)
                                rightBeamClip.ClipHeight1 = flangeThickness + 2 * rightBeamClip.ClipThickness;
                            if (rightBeamClip.ShapeType == 5)
                            {
                                if (HgrCompareDoubleService.cmpdbl(rightBeamClip.ClipHeight2 , 0) == true)
                                    rightBeamClip.ClipHeight2 = height3 + 2 * rightBeamClip.ClipThickness;
                            }
                        }
                        if ((HgrCompareDoubleService.cmpdbl(rightBeamClip.Bolt.BoltLength , 0) == false) && (rightBeamClip.ShapeType != 7))
                        {
                            vertOffset2 = 0;
                            if (rightBeamClip.ShapeType == 1)
                                vertOffset1 = rightBeamClip.ClipHeight1 - rightBeamClip.ClipThickness1 - rightBeamClip.ClipThickness2 - flangeThickness;
                            else if ((rightBeamClip.ShapeType == 2) || (rightBeamClip.ShapeType == 3))
                                vertOffset1 = rightBeamClip.ClipHeight1 - rightBeamClip.ClipThickness1 - flangeThickness;
                            else if (rightBeamClip.ShapeType == 4)
                                vertOffset1 = rightBeamClip.ClipHeight1 - rightBeamClip.ClipThickness - rightBeamClip.ClipHeight2 - flangeThickness;
                            else if (rightBeamClip.ShapeType == 5)
                                vertOffset1 = rightBeamClip.ClipHeight1 - 2 * rightBeamClip.ClipThickness - flangeThickness;
                            else if (rightBeamClip.ShapeType == 6)
                                vertOffset1 = rightBeamClip.ClipHeight1 - 2 * rightBeamClip.ClipThickness - flangeThickness;
                        }
                        else
                        {
                            vertOffset1 = 0;
                            if (rightBeamClip.ShapeType == 1)
                                vertOffset2 = rightBeamClip.ClipHeight1 - rightBeamClip.ClipThickness1 - rightBeamClip.ClipThickness2 - flangeThickness;
                            else if ((rightBeamClip.ShapeType == 2) || (rightBeamClip.ShapeType == 3))
                                vertOffset2 = rightBeamClip.ClipHeight1 - rightBeamClip.ClipThickness1 - flangeThickness;
                            else if (rightBeamClip.ShapeType == 4)
                                vertOffset2 = rightBeamClip.ClipHeight1 - rightBeamClip.ClipThickness - rightBeamClip.ClipHeight2 - flangeThickness;
                            else if (rightBeamClip.ShapeType == 5)
                                vertOffset2 = rightBeamClip.ClipHeight1 - 2 * rightBeamClip.ClipThickness - flangeThickness;
                            else if (rightBeamClip.ShapeType == 6)
                                vertOffset2 = rightBeamClip.ClipHeight1 - 2 * rightBeamClip.ClipThickness - flangeThickness;
                        }
                        Port port2 = new Port(OccurrenceConnection, part, "Top", new Position(0, -flangeWidth / 2, vertOffset1 + flangeThickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_PhysicalAspect.Outputs["Top"] = port2;
                        if ((rightBeamClip.ShapeType == 8) && (HgrCompareDoubleService.cmpdbl(uBoltWidth , 0) == false))
                        {
                            Port port3 = new Port(OccurrenceConnection, part, "Bottom", new Position(0, -(flangeWidth / 2 + beamClamp.RightClipGap + beamClamp.RightOffsetH), beamClamp.RightOffsetV - flangeThickness / 2 + uBoltRodDia), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_PhysicalAspect.Outputs["Bottom"] = port3;
                        }
                        else
                        {
                            Port port3 = new Port(OccurrenceConnection, part, "Bottom", new Position(0, -flangeWidth / 2, -vertOffset2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_PhysicalAspect.Outputs["Bottom"] = port3;
                        }
                        if (rightBeamClip.ShapeType > 0 && rightBeamClip.ShapeType < 7)
                        {
                            Port port4 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, flangeWidth / 2 + beamClamp.RightClipGap + beamClamp.RightOffsetH, beamClamp.RightOffsetV), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_PhysicalAspect.Outputs["RodEnd1"] = port4;
                        }
                    }
                    else
                    {
                        if (rightBeamClip.ShapeType > 0 && rightBeamClip.ShapeType < 5)
                        {
                            Port port5 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, -(flangeWidth / 2.0 + beamClamp.RightClipGap + beamClamp.RightOffsetH), beamClamp.RightOffsetV), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_PhysicalAspect.Outputs["RodEnd2"] = port5;
                        }
                    }
                }

                if ((HgrCompareDoubleService.cmpdbl(width1 , 0) == false) && (HgrCompareDoubleService.cmpdbl(length1 , 0) == false) && (HgrCompareDoubleService.cmpdbl(thickness1 , 0) == false))
                {
                    if (!(beamClamp.BottomShape == 1))
                    {
                        if (angle1 != 0)
                        {
                            Port port6 = new Port(OccurrenceConnection, part, "RodEnd", new Position(Math.Sin(Math.PI + angle1) * (rodTakeOut), 0, Math.Cos(Math.PI + angle1) * (rodTakeOut)), new Vector(Math.Cos(angle1), 0, -Math.Sin(angle1)), new Vector(-Math.Sin(angle1), 0, -Math.Cos(angle1)));
                            m_PhysicalAspect.Outputs["RodEnd"] = port6;
                        }
                        else
                        {
                            Port port6 = new Port(OccurrenceConnection, part, "RodEnd", new Position(0, 0, -rodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                            m_PhysicalAspect.Outputs["RodEnd"] = port6;
                        }
                    }
                }

                //Add Beam Clamp
                matrix = new Matrix4X4();
                matrix.Origin = new Position(0, 0, 0);
                AddBeamClamp(beamClamp, flangeThickness, flangeWidth, height3, matrix, m_PhysicalAspect.Outputs, "BeamClamp");

                //Add Bottom Shape

                if ((HgrCompareDoubleService.cmpdbl(width1 , 0) == false) && (HgrCompareDoubleService.cmpdbl(length1 , 0) == false) && (HgrCompareDoubleService.cmpdbl(thickness1 , 0) == false))
                {
                    switch (beamClamp.BottomShape)
                    {
                        case 1:     //I shape
                            //Throw a warning for Length is not enough
                            if (length1 < flangeWidth + beamClamp.LeftClipGap + beamClamp.RightClipGap)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrStructureWidthGaps, "Length1 is less than Structure Width with Gaps");
                                return;
                            }
                            symbolGeometryHelper.ActivePosition = new Position(-width1 / 2, -length1 / 2, -thickness1 - vertOffset2);
                            symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                            Projection3d iShapeBox = symbolGeometryHelper.CreateBox(null, width1, length1, thickness1, 9);
                            m_PhysicalAspect.Outputs["BeamClampBottom"] = iShapeBox;
                            break;
                        case 2:
                            WBAHoleInputs wbaHole = new WBAHoleInputs();
                            Double length;
                            //Throw a warning for Length is not enough
                            if ((2 * length1 + length2) < flangeWidth + beamClamp.LeftClipGap + beamClamp.RightClipGap)
                            {
                                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrStructureWidthGapsClips, "Length1 is less than Structure Width with Gaps and Clip Widths. Resetting the value"));
                                length = (flangeWidth + beamClamp.LeftClipGap + beamClamp.RightClipGap + leftBeamClip.ClipWidth2 + rightBeamClip.ClipWidth2);
                            }
                            else
                                length = 2 * length1 + length2;

                            //Map the attributes with WBA Hole attributes
                            wbaHole.WBAHoleConfig = 2;
                            wbaHole.Thickness1 = thickness1;
                            wbaHole.Gap1 = length1;
                            wbaHole.Thickness3 = thickness1;
                            wbaHole.Height1 = height1;
                            wbaHole.Width1 = width1;
                            wbaHole.Length1 = length1;
                            wbaHole.Length2 = length;

                            matrix = new Matrix4X4();
                            matrix.Origin = new Position(0, 0, -vertOffset2);
                            AddWBAHole(wbaHole, matrix, m_PhysicalAspect.Outputs, "BeamClampBottom");
                            break;
                        case 3:
                            ClevisHangerInputs clevisHgr = new ClevisHangerInputs();

                            //Throw a warning for Length is not enough
                            if (length1 < flangeWidth + beamClamp.LeftClipGap + beamClamp.RightClipGap)
                            {
                                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrStructureWidthGaps, "Length1 is less than Structure Width with Gaps. Resetting the value"));
                                length1 = flangeWidth + beamClamp.LeftClipGap + beamClamp.RightClipGap;
                            }

                            //Map the attributes with Clevis Hanger attributes
                            clevisHgr.ClevisTopShp = 1;
                            clevisHgr.ClevisBotShp = 3;
                            clevisHgr.Width1 = width1;
                            clevisHgr.Thickness1 = thickness1;
                            clevisHgr.Height3 = height1;
                            clevisHgr.Height4 = height2;
                            clevisHgr.Length1 = length2;
                            clevisHgr.Length2 = length1 - 2 * thickness1;

                            clevisHgr.Height1 = 0.002;
                            clevisHgr.RodTakeOut = 0.001;

                            matrix = new Matrix4X4();
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                            matrix.Rotate(angle1, new Vector(0, 1, 0));
                            matrix.Translate(matrix.Transform(new Vector(0, 0, height1 - thickness1)));
                            AddClevisHanger(clevisHgr, 0, 0, matrix, m_PhysicalAspect.Outputs, "BeamClampBottom");
                            break;
                    }
                }
                if ((beamClamp.LeftClipShape != "") && (beamClamp.LeftClipShape != "No Value") && (beamClamp.RightClipShape != "") && (beamClamp.RightClipShape != "No Value"))
                {
                    //Add Pin
                    if ((HgrCompareDoubleService.cmpdbl(pin1Diameter , 0) == false) || (HgrCompareDoubleService.cmpdbl(pin1Length , 0) == false))
                    {
                        Matrix4X4 cylinderMatrix = new Matrix4X4();
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                        cylinderMatrix.SetIdentity();
                        cylinderMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                        if ((beamClamp.BottomShape == 3) && (HgrCompareDoubleService.cmpdbl(width1 , 0) == false) && (HgrCompareDoubleService.cmpdbl(length1 , 0) == false) && (HgrCompareDoubleService.cmpdbl(thickness1 , 0) == false) && (HgrCompareDoubleService.cmpdbl(angle1 , 0) == false))
                        {
                            cylinderMatrix.Translate(cylinderMatrix.Transform(new Vector(0, 0, -pin1Length / 2)));
                            Projection3d pinCylinder = symbolGeometryHelper.CreateCylinder(null, pin1Diameter / 2, pin1Length);
                            pinCylinder.Transform(cylinderMatrix);
                            m_PhysicalAspect.Outputs["BeamClampPin"] = pinCylinder;
                        }
                        else
                        {
                            cylinderMatrix.Translate(cylinderMatrix.Transform(new Vector(0, -height3, -pin1Length / 2)));
                            Projection3d pinCylinder = symbolGeometryHelper.CreateCylinder(null, pin1Diameter / 2, pin1Length);
                            pinCylinder.Transform(cylinderMatrix);
                            m_PhysicalAspect.Outputs["BeamClampPin"] = pinCylinder;
                        }
                    }

                    //Add Left Bolt
                    if ((leftBolt.BoltLength > 0) && (leftBolt.BoltDiameter > 0))
                    {
                        matrix = new Matrix4X4();
                        matrix.Origin = new Position(0, flangeWidth / 2 + beamClamp.LeftClipGap + beamClamp.LeftOffsetH, -leftBolt.BoltLength + beamClamp.LeftOffsetV);
                        AddBoltWithHead(leftBolt, matrix, m_PhysicalAspect.Outputs, "BeamClampBolt1");
                    }

                    //Add Right Bolt
                    if ((rightBolt.BoltLength > 0) && (rightBolt.BoltDiameter > 0))
                    {
                        matrix = new Matrix4X4();
                        matrix.Origin = new Position(0, -flangeWidth / 2 - beamClamp.RightClipGap - beamClamp.RightOffsetH, -rightBolt.BoltLength + beamClamp.RightOffsetV);
                        AddBoltWithHead(rightBolt, matrix, m_PhysicalAspect.Outputs, "BeamClampBolt2");
                    }
                }

                //Add J Bolt
                if ((HgrCompareDoubleService.cmpdbl(uBoltWidth, 0) == false) && (HgrCompareDoubleService.cmpdbl(uBoltLength, 0) == false) && (HgrCompareDoubleService.cmpdbl(uBoltRodDia, 0) == false))
                {
                    UBoltInputs uBolt = new UBoltInputs();
                    //Throw a warning for Ubolt Length
                    if (((beamClamp.LeftClipShape != "") && (beamClamp.LeftClipShape != "No Value")) && ((beamClamp.RightClipShape == "") || (beamClamp.RightClipShape == "No Value")))
                    {
                        if (!(leftBeamClip.ShapeType >= 7))
                        {
                            if (uBoltLength < flangeWidth + beamClamp.LeftClipGap)
                            {
                                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrUBoltStructureWidthGaps, "UBolt Length is less than Structure Width with Gaps. Resetting the value"));
                                uBoltLength = flangeWidth + beamClamp.LeftClipGap + 2 * uBoltRodDia;
                            }
                        }
                        else    //For I and 2I
                        {
                            if (uBoltLength < beamClamp.LeftOffsetV)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrUBoltLengthLeftOffSet, "UBolt Length is less than Left Vertical Offset");
                                return;
                            }
                        }
                    }
                    else if (((beamClamp.RightClipShape != "") && (beamClamp.RightClipShape != "No Value")) && ((beamClamp.LeftClipShape == "") || (beamClamp.LeftClipShape == "No Value")))
                    {
                        if (!(rightBeamClip.ShapeType >= 7))
                        {
                            if (uBoltLength < flangeWidth + beamClamp.RightClipGap)
                            {
                                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrUBoltStructureWidthGaps, "UBolt Length is less than Structure Width with Gaps. Resetting the value"));
                                uBoltLength = flangeWidth + beamClamp.RightClipGap + 2 * uBoltRodDia;
                            }
                        }
                        else    //For I and 2I
                        {
                            if (uBoltLength < beamClamp.RightOffsetV)
                            {
                                RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrUBoltLengthRightOffSet, "UBolt Length is less than Right Vertical Offset");
                                return;
                            }
                        }
                    }
                    uBolt.UBoltWidth = uBoltWidth;
                    uBolt.UBoltCenterToEnd = uBoltLength;
                    uBolt.UBoltRodDia = uBoltRodDia;
                    uBolt.UBoltFlatSpot = uBoltFlatSpot;
                    uBolt.UBoltOneSided = isJBolt;

                    if (((beamClamp.LeftClipShape != "") && (beamClamp.LeftClipShape != "No Value")) && ((beamClamp.RightClipShape == "") || (beamClamp.RightClipShape == "No Value")))
                    {
                        matrix = new Matrix4X4();
                        if (leftBeamClip.ShapeType >= 7)        //I or 2I
                        {
                            matrix.Translate(new Vector(flangeWidth / 2 + beamClamp.LeftClipGap + beamClamp.LeftOffsetH, 0, -beamClamp.LeftOffsetV - flangeThickness / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddUBolt(uBolt, flangeThickness, matrix, m_PhysicalAspect.Outputs, "BeamClampJBolt");
                        }
                        else
                        {
                            matrix.Translate(new Vector(0, uBolt.UBoltWidth / 2 - uBolt.UBoltRodDia / 2 - flangeThickness, flangeWidth / 2 - flangeThickness / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddUBolt(uBolt, flangeThickness, matrix, m_PhysicalAspect.Outputs, "BeamClampJBolt");
                        }
                    }
                    else if (((beamClamp.RightClipShape != "") && (beamClamp.RightClipShape != "No Value")) && ((beamClamp.LeftClipShape == "") || (beamClamp.LeftClipShape == "No Value")))
                    {
                        matrix = new Matrix4X4();
                        if (rightBeamClip.ShapeType >= 7)       //I or 2I
                        {
                            matrix.Translate(new Vector(-(flangeWidth / 2 + beamClamp.RightClipGap + beamClamp.RightOffsetH), 0, -beamClamp.RightOffsetV - flangeThickness / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddUBolt(uBolt, flangeThickness, matrix, m_PhysicalAspect.Outputs, "BeamClampJBolt");
                        }
                        else
                        {
                            matrix.Translate(new Vector(0, uBolt.UBoltWidth / 2 - uBolt.UBoltRodDia / 2 - flangeThickness, flangeWidth / 2 - flangeThickness / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddUBolt(uBolt, flangeThickness, matrix, m_PhysicalAspect.Outputs, "BeamClampJBolt");
                        }
                    }
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of BeamClamp"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch (SmartPartSymbolException hgrEx)
            {
                throw hgrEx;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of BeamClamp"));
                    return;
                }
            }
        }
        #endregion
    }

}
