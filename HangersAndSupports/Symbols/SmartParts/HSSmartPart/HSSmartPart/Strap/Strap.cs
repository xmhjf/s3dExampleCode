//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Strap.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strap
//   Author       :  Hema
//   Creation Date:  25.Jan.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   25.Jan.2013     Hema    CR-CP-222465  Initial Creation 
//   25/Mar/2013     Vijay   DI-CP-228142  Modify the error handling for delivered H&S symbols
//   17-Oct-2014     Chethan CR-CP-253371  Add Maintenance Aspect to PipeClamp smartpart and Strap smartpart
//   11-02-2013      Chethan DI-CP-263820  Fix priority 3 items to .net SmartParts as a result of new testing  
//   13.08.2015      PR      TR 276067	Few smart parts are having cache option as non-cached when they should be cached
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    [SymbolVersion("1.0.0.1")]
    [VariableOutputs]
    public class Strap : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strap"
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
                AddStrapInputs(2, out endIndex, additionalInputs);
                additionalInputs.Add(new InputDouble(++endIndex, "Pin1Diameter", "Pin1Diameter", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Pin1Length", "Pin1Length", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Offset1", "Offset1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Offset2", "Offset2", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi1Qty", "Multi1Qty", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi1LocateBy", "Multi1LocateBy", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi1Location", "Multi1Location", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Gap1", "Gap1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Height1", "Height1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Width1", "Width1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Length1", "Length1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Height2", "Height2", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Width2", "Width2", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Length2", "Length2", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Width4", "Width4", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Height4", "Height4", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Thickness4", "Thickness4", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Width5", "Width5", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Width6", "Width6", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Pin2Diameter", "Pin2Diameter", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Pin2Length", "Pin2Length", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Offset3", "Offset3", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi2Qty", "Multi2Qty", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi2LocateBy", "Multi2LocateBy", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi2Location", "Multi2Location", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Diameter3", "Diameter3", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Length3", "Length3", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Height3", "Height3", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Width3", "Width3", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Thickness3", "Thickness3", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi3Qty", "Multi3Qty", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi3LocateBy", "Multi3LocateBy", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Multi3Location", "Multi3Location", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "MinPipeToSteel", "MinPipeToSteel", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "MaxPipeToSteel", "MaxPipeToSteel", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "PipeOD", "PipeOD", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "PipeToSteel", "PipeToSteel", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "RotY", "RotY", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "MaintenanceThickness", "MaintenanceThickness", 0, true));
                additionalInputs.Add(new InputDouble(++endIndex, "CreateMaintenanceAspect", "CreateMaintenanceAspect", 0, true));
                return additionalInputs;
            }
        }
        #endregion

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Steel", "Steel")]
        public AspectDefinition m_PhysicalAspect;

        [Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)]
        public AspectDefinition m_Maintenance;

        #endregion


        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddStrapOutputs(additionalOutputs);
                additionalOutputs.Add(new OutputDefinition("Bolts", "Bolts"));
                additionalOutputs.Add(new OutputDefinition("Blocks", "Blocks"));
                additionalOutputs.Add(new OutputDefinition("Wrap", "Wrap"));
                additionalOutputs.Add(new OutputDefinition("StripBolt", "StripBolt"));
                additionalOutputs.Add(new OutputDefinition("Gussets", "Gussets"));
                additionalOutputs.Add(new OutputDefinition("Bolts", "Bolts"));
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
                int startIndex = 12;

                StrapInputs strap = LoadStrapData(2);
                Double pin1Diameter = GetDoubleInputValue(startIndex);
                Double pin1Length = GetDoubleInputValue(++startIndex);
                Double offset1 = GetDoubleInputValue(++startIndex);
                Double offset2 = GetDoubleInputValue(++startIndex);
                Double multi1Qty = GetDoubleInputValue(++startIndex);
                Double multi1LocateBy = GetDoubleInputValue(++startIndex);
                Double multi1Location = GetDoubleInputValue(++startIndex);
                Double gap1 = GetDoubleInputValue(++startIndex);
                Double height1 = GetDoubleInputValue(++startIndex);
                Double width1 = GetDoubleInputValue(++startIndex);
                Double length1 = GetDoubleInputValue(++startIndex);
                Double height2 = GetDoubleInputValue(++startIndex);
                Double width2 = GetDoubleInputValue(++startIndex);
                Double length2 = GetDoubleInputValue(++startIndex);
                Double width4 = GetDoubleInputValue(++startIndex);
                Double height4 = GetDoubleInputValue(++startIndex);
                Double thickness4 = GetDoubleInputValue(++startIndex);
                Double width5 = GetDoubleInputValue(++startIndex);
                Double width6 = GetDoubleInputValue(++startIndex);
                Double pin2Diameter = GetDoubleInputValue(++startIndex);
                Double pin2Length = GetDoubleInputValue(++startIndex);
                Double offset3 = GetDoubleInputValue(++startIndex);
                Double multi2Qty = GetDoubleInputValue(++startIndex);
                Double multi2LocateBy = GetDoubleInputValue(++startIndex);
                Double multi2Location = GetDoubleInputValue(++startIndex);
                Double diameter3 = GetDoubleInputValue(++startIndex);
                Double length3 = GetDoubleInputValue(++startIndex);
                Double height3 = GetDoubleInputValue(++startIndex);
                Double width3 = GetDoubleInputValue(++startIndex);
                Double thickness3 = GetDoubleInputValue(++startIndex);
                Double multi3Qty = GetDoubleInputValue(++startIndex);
                Double multi3LocateBy = GetDoubleInputValue(++startIndex);
                Double multi3Location = GetDoubleInputValue(++startIndex);
                Double minPipeToSteel = GetDoubleInputValue(++startIndex);
                Double maxPipeToSteel = GetDoubleInputValue(++startIndex);
                Double pipeOD = GetDoubleInputValue(++startIndex);
                Double pipeToSteel = GetDoubleInputValue(++startIndex);
                Double rotY = GetDoubleInputValue(++startIndex);
                Double maintenanceThickness = GetDoubleInputValue(++startIndex);
                bool createMaintenanceAspect = false;
                createMaintenanceAspect = Convert.ToBoolean(GetDoubleInputValue(++startIndex));
                Matrix4X4 matrix = new Matrix4X4();
                rotY = rotY * 180 / Math.PI;

                //Initializing symbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(Math.Cos(rotY * Math.PI / 180), 0, Math.Sin(rotY * Math.PI / 180)), new Vector(Math.Cos((270 + rotY) * Math.PI / 180), 0, Math.Sin((270 + rotY) * Math.PI / 180)));
                m_PhysicalAspect.Outputs["Route"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, -pipeOD / 2.0 - pipeToSteel), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Steel"] = port2;

                //Strap
                AddStrap(strap, pipeOD, matrix, m_PhysicalAspect.Outputs, "Strap");

                //Hold Down Bolts
                if (pin1Diameter > 0 && pin1Length > 0 && multi1Qty > 0)
                {
                    //Bolts
                    AddBoltsByRow(strap.StrapStockWidth, multi1Qty, multi1LocateBy, multi1Location, pin1Diameter, pin1Length, offset1, 90, 1, m_PhysicalAspect.Outputs, "LeftBolts", "LeftBolts", 0, false, -pin1Length / 2 + pipeOD / 2.0 + strap.StrapTopGap - strap.StrapHeightInside + offset2);

                    if (strap.StrapOneSided == 2)
                        AddBoltsByRow(strap.StrapStockWidth, multi1Qty, multi1LocateBy, multi1Location, pin1Diameter, pin1Length, -offset1, 90, 1, m_PhysicalAspect.Outputs, "RightBolts", "RightBolts", 0, false, -pin1Length / 2 + pipeOD / 2.0 + strap.StrapTopGap - strap.StrapHeightInside + offset2);
                }

                //Blocks
                if (height1 > 0 && width1 > 0 && length1 > 0)
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, -pipeOD / 2.0 - gap1 - height1);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                    Projection3d block1 = (Projection3d)symbolGeometryHelper.CreateBox(null, height1, width1, length1);
                    m_PhysicalAspect.Outputs["Block1"] = block1;
                }
                if (height2 > 0 && width2 > 0 && length2 > 0)
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, -pipeOD / 2.0 - gap1);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                    Projection3d block2 = (Projection3d)symbolGeometryHelper.CreateBox(null, height2, width2, length2);
                    m_PhysicalAspect.Outputs["Block2"] = block2;
                }

                //Rectangular Wrap/Strap (Either a StrapShape or a Plate depending on the inputs.
                if (width4 > 0 && height4 > 0 && width5 > 0)
                {
                    if (thickness4 <= 0 || thickness4 >= width4 / 2 || thickness4 >= height4)
                    {
                        //it is a block, use the Plate Shape Function
                        PlateInputs plate = new PlateInputs();
                        plate.width1 = width4;
                        plate.length1 = height4;
                        plate.thickness1 = width5;
                        plate.tlCornerType = 7;
                        plate.tlCornerRad = (width4 - width6) / 2;
                        plate.trCornerType = 7;
                        plate.trCornerRad = (width4 - width6) / 2;
                        matrix.SetIdentity();
                        matrix.Rotate((Math.PI) / 2, new Vector(1, 0, 0));
                        matrix.Rotate(3 * ((Math.PI) / 2), new Vector(0, 0, 1));

                        matrix.Translate(new Vector(width5 / 2, width4 / 2, pipeOD / 2 + strap.StrapTopGap - height4));

                        AddPlate(plate, matrix, m_PhysicalAspect.Outputs, "RectangularWrap");
                    }
                    else
                    {
                        //It is a strap, use the strap shape function.
                        StrapInputs rectStrap = new StrapInputs();
                        rectStrap.StrapWidthInside = width4 - 2 * thickness4;
                        rectStrap.StrapHeightInside = height4 - thickness4;
                        rectStrap.StrapThickness = thickness4;
                        rectStrap.StrapStockWidth = width5;
                        rectStrap.StrapFlatSpot = width6;
                        rectStrap.StrapWidthWings = 0;
                        rectStrap.StrapOneSided = 2;
                        rectStrap.StrapSplitGap = 0;
                        rectStrap.StrapSplitExtension = 0;
                        rectStrap.StrapTopGap = strap.StrapTopGap + (width4 - strap.StrapWidthInside) / 2 - thickness4;
                        AddStrap(rectStrap, pipeOD, matrix, m_PhysicalAspect.Outputs, "StrapWrap");
                    }

                }

                //Split Strap Bolts
                if (pin2Diameter > 0 && pin2Length > 0 && multi2Qty > 0)
                    AddBoltsByRow(strap.StrapStockWidth, multi2Qty, multi2LocateBy, multi2Location, pin2Diameter, pin2Length, pipeOD / 2 + strap.StrapTopGap + offset3, 0, 1, m_PhysicalAspect.Outputs, "StrapBolts", "StrapBolts", 0, false, 0);
                //Full Circle Wrap
                if (diameter3 > 0 && length3 > 0)
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-length3 / 2, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Projection3d circularWrap = new Projection3d(symbolGeometryHelper.CreateCylinder(null, diameter3 / 2.0, length3));
                    m_PhysicalAspect.Outputs["CircularWrap"] = circularWrap;
                }
                //Gussets
                double wingOffset;

                if (height3 > 0 && width3 > 0 && thickness3 > 0 && multi3Qty > 0)
                {
                    if (strap.StrapWidthWings > (strap.StrapWidthInside + 2 * strap.StrapThickness))
                        wingOffset = strap.StrapThickness;
                    else
                        wingOffset = 0;

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -strap.StrapWidthInside / 2 - strap.StrapThickness, pipeOD / 2 + strap.StrapTopGap - strap.StrapHeightInside + wingOffset));
                    AddGussetsByRow(strap.StrapStockWidth, multi3Qty, multi3LocateBy, multi3Location, width3, height3, thickness3, matrix, m_PhysicalAspect.Outputs, "Gussets", 1);

                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                    Vector vector = matrix.Transform(new Vector(0, -strap.StrapWidthInside / 2 - strap.StrapThickness, pipeOD / 2 + strap.StrapTopGap - strap.StrapHeightInside + wingOffset));
                    matrix.Translate(vector);
                    AddGussetsByRow(strap.StrapStockWidth, multi3Qty, multi3LocateBy, multi3Location, width3, height3, thickness3, matrix, m_PhysicalAspect.Outputs, "Gussets", 2);
                }

                // Construction of Maintenance Aspect

                if (createMaintenanceAspect)
                {
                    if (strap.StrapThickness > 0)
                    {
                        strap.StrapThickness = strap.StrapThickness + maintenanceThickness;
                    }
                    if (strap.StrapWidthWings > 0)
                    {
                        strap.StrapWidthWings = strap.StrapWidthWings + 2 * maintenanceThickness;
                    }
                    if (strap.StrapSplitExtension > 0)
                    {
                        strap.StrapSplitExtension = strap.StrapSplitExtension + maintenanceThickness;
                    }
                    if (height1 > 0)
                    {
                        width1 = strap.StrapWidthWings;
                    }

                    //Strap
                    AddStrap(strap, pipeOD, matrix, m_Maintenance.Outputs, "Strap1");

                    //Hold Down Bolts

                    //Blocks
                    if (height1 > 0 && width1 > 0 && length1 > 0)
                    {
                        symbolGeometryHelper.ActivePosition = new Position(0, 0, -pipeOD / 2.0 - gap1 - height1);
                        symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                        Projection3d block1 = (Projection3d)symbolGeometryHelper.CreateBox(OccurrenceConnection, height1, width1, length1);
                        m_Maintenance.Outputs["Block1"] = block1;
                    }
                    if (height2 > 0 && width2 > 0 && length2 > 0)
                    {
                        symbolGeometryHelper.ActivePosition = new Position(0, 0, -pipeOD / 2.0 - gap1);
                        symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                        Projection3d block2 = (Projection3d)symbolGeometryHelper.CreateBox(OccurrenceConnection, height2, width2, length2);
                        m_Maintenance.Outputs["Block2"] = block2;
                    }

                    //Rectangular Wrap/Strap (Either a StrapShape or a Plate depending on the inputs.
                    if (width4 > 0 && height4 > 0 && width5 > 0)
                    {
                        if (thickness4 <= 0 || thickness4 >= width4 / 2 || thickness4 >= height4)
                        {
                            //it is a block, use the Plate Shape Function
                            PlateInputs plate = new PlateInputs();
                            plate.width1 = width4;
                            plate.length1 = height4;
                            plate.thickness1 = width5;
                            plate.tlCornerType = 7;
                            plate.tlCornerRad = (width4 - width6) / 2;
                            plate.trCornerType = 7;
                            plate.trCornerRad = (width4 - width6) / 2;
                            matrix.SetIdentity();
                            matrix.Rotate((Math.PI) / 2, new Vector(1, 0, 0));
                            matrix.Rotate(3 * ((Math.PI) / 2), new Vector(0, 0, 1));

                            matrix.Translate(new Vector(width5 / 2, width4 / 2, pipeOD / 2 + strap.StrapTopGap - height4));

                            AddPlate(plate, matrix, m_Maintenance.Outputs, "RectangularWrap");
                        }
                        else
                        {
                            //It is a strap, use the strap shape function.
                            StrapInputs rectStrap = new StrapInputs();
                            rectStrap.StrapWidthInside = width4 - 2 * thickness4;
                            rectStrap.StrapHeightInside = height4 - thickness4;
                            rectStrap.StrapThickness = thickness4;
                            rectStrap.StrapStockWidth = width5;
                            rectStrap.StrapFlatSpot = width6;
                            rectStrap.StrapWidthWings = 0;
                            rectStrap.StrapOneSided = 2;
                            rectStrap.StrapSplitGap = 0;
                            rectStrap.StrapSplitExtension = 0;
                            rectStrap.StrapTopGap = strap.StrapTopGap + (width4 - strap.StrapWidthInside) / 2 - thickness4;
                            AddStrap(rectStrap, pipeOD, matrix, m_Maintenance.Outputs, "StrapWrap");
                        }

                    }

                    //Full Circle Wrap
                    if (diameter3 > 0 && length3 > 0)
                    {
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(-length3 / 2, 0, 0);
                        symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                        Projection3d circularWrap = new Projection3d(symbolGeometryHelper.CreateCylinder(null, diameter3 / 2.0, length3));
                        m_Maintenance.Outputs["CircularWrap"] = circularWrap;
                    }
                    //Gussets

                    if (height3 > 0 && width3 > 0 && thickness3 > 0 && multi3Qty > 0)
                    {
                        if (strap.StrapWidthWings > (strap.StrapWidthInside + 2 * strap.StrapThickness))
                            wingOffset = strap.StrapThickness;
                        else
                            wingOffset = 0;

                        matrix.SetIdentity();
                        matrix.Translate(new Vector(0, -strap.StrapWidthInside / 2 - strap.StrapThickness, pipeOD / 2 + strap.StrapTopGap - strap.StrapHeightInside + wingOffset));
                        AddGussetsByRow(strap.StrapStockWidth, multi3Qty, multi3LocateBy, multi3Location, width3, height3, thickness3, matrix, m_Maintenance.Outputs, "Gussets", 1);

                        matrix.SetIdentity();
                        matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                        Vector vector = matrix.Transform(new Vector(0, -strap.StrapWidthInside / 2 - strap.StrapThickness, pipeOD / 2 + strap.StrapTopGap - strap.StrapHeightInside + wingOffset));
                        matrix.Translate(vector);
                        AddGussetsByRow(strap.StrapStockWidth, multi3Qty, multi3LocateBy, multi3Location, width3, height3, thickness3, matrix, m_Maintenance.Outputs, "Gussets", 2);
                    }
                }
            }

            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Strap"));
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
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Strap"));
                    return;
                }
            }
        }
        #endregion
    }

}
