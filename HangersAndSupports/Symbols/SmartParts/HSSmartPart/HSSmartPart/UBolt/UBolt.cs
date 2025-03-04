//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UBolt.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.UBolt
//   Author       :  Hema
//   Creation Date:  01-Feb-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01-Feb-2013    Hema     CR-CP-222468  Converted UBolt VB Project to C# .Net
//   25/Mar/2013    Vijay    DI-CP-228142  Modify the error handling for delivered H&S symbols
//   31/10/2013     Rajeswari    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   11/Aug/2014    Ramya    TR-CP-256377  Additional input values are retrieved from catalog part in smart parts
//   12-12-2014     PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report  
//   11-02-2013       Chethan DI-CP-263820  Fix priority 3 items to .net SmartParts as a result of new testing  
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
    public class UBolt : SmartPartComponentDefinition, ICustomWeightCG
    {

        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.UBolt"
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
                AddUBoltInputs(2, out endIndex, additionalInputs);
                AddNutInputs(endIndex, 4, out endIndex, additionalInputs);
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Gap1", "Gap1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Height1", "Height1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Width1", "Width1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Length1", "Length1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Height2", "Height2", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Width2", "Width2", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Length2", "Length2", 0, false));
                AddStrapInputs(++endIndex, out endIndex, additionalInputs);
                additionalInputs.Add((Input)new InputDouble(++endIndex, "MinPipeToSteel", "MinPipeToSteel", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "MaxPipeToSteel", "MaxPipeToSteel", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "PipeOD", "PipeOD", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "PipeToSteel", "PipeToSteel", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "SteelThickness", "SteelThickness", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "RotY", "RotY", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Diameter3", "Diameter3", 0, true));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Length3", "Length3", 0, true));
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

        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddUBoltOutputs(additionalOutputs);
                AddNutOutputs(4, additionalOutputs);
                AddStrapOutputs(additionalOutputs);
                additionalOutputs.Add(new OutputDefinition("Blocks", "Blocks"));
                additionalOutputs.Add(new OutputDefinition("Wrap", "Wrap"));
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
                UBoltInputs ubolt = LoadUBoltData(2, out endIndex);
                if (base.ToDoListMessage != null)
                {
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }
                startIndex = endIndex;

                NutInputs[] nutCollection = new NutInputs[4];
                for (int k = 0; k <= 3; k++)
                {
                    nutCollection[k] = LoadNutData(++startIndex, out endIndex);
                    startIndex = endIndex;
                }
                Double gap1 = GetDoubleInputValue(++startIndex);
                Double height1 = GetDoubleInputValue(++startIndex);
                Double width1 = GetDoubleInputValue(++startIndex);
                Double length1 = GetDoubleInputValue(++startIndex);
                Double height2 = GetDoubleInputValue(++startIndex);
                Double width2 = GetDoubleInputValue(++startIndex);
                Double length2 = GetDoubleInputValue(++startIndex);
                StrapInputs strap = LoadStrapData(++startIndex, out endIndex);
                startIndex = endIndex;
                Double minPipeToSteel = GetDoubleInputValue(++startIndex);
                Double maxPipeToSteel = GetDoubleInputValue(++startIndex);
                Double pipeOD = GetDoubleInputValue(++startIndex);
                Double pipeToSteel = GetDoubleInputValue(++startIndex);
                Double steelThickness = GetDoubleInputValue(++startIndex);
                Double rotY = GetDoubleInputValue(++startIndex);
                Double diameter3 = GetDoubleInputValue(++startIndex);
                Double length3 = GetDoubleInputValue(++startIndex);

                rotY = rotY * 180 / Math.PI;

                Matrix4X4 matrix = new Matrix4X4();

                //Initializing symbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                /*
                try
                {
                    diameter3 = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsUBoltDiameter3", "Diameter3")).PropValue;
                }
                catch
                {
                    diameter3 = 0;
                }
                try
                {
                    length3 = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsUBoltLength3", "Length3")).PropValue;
                }
                catch
                {
                    length3 = 0;
                }
                */
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(Math.Cos(rotY * Math.PI / 180), 0, Math.Sin(rotY * Math.PI / 180)), new Vector(Math.Cos((270 + rotY) * Math.PI / 180), 0, Math.Sin((270 + rotY) * Math.PI / 180)));
                m_PhysicalAspect.Outputs["Route"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, -pipeOD / 2.0 - gap1 - height1), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Steel"] = port2;

                //'U Bolt
                AddUBolt(ubolt, pipeOD, matrix, m_PhysicalAspect.Outputs, "UBolt");
                //'Nuts 1-4 Left Side
                matrix.Translate(new Vector(0, ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1 - height1 - steelThickness - nutCollection[0].ShapeLength));
                AddNut(nutCollection[0], matrix, m_PhysicalAspect.Outputs, "NutL1");

                matrix.SetIdentity();
                matrix.Translate(new Vector(0, ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1 - height1 - steelThickness - nutCollection[0].ShapeLength - nutCollection[1].ShapeLength));
                AddNut(nutCollection[1], matrix, m_PhysicalAspect.Outputs, "NutL2");

                matrix.SetIdentity();
                matrix.Translate(new Vector(0, ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1));
                AddNut(nutCollection[2], matrix, m_PhysicalAspect.Outputs, "NutL3");

                matrix.SetIdentity();
                matrix.Translate(new Vector(0, ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1 + nutCollection[2].ShapeLength));
                AddNut(nutCollection[3], matrix, m_PhysicalAspect.Outputs, "NutL4");

                //Nuts 1-4 Right Side
                if (ubolt.UBoltOneSided == 2)
                {
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1 - height1 - steelThickness - nutCollection[0].ShapeLength));
                    AddNut(nutCollection[0], matrix, m_PhysicalAspect.Outputs, "NutR1");

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1 - height1 - steelThickness - nutCollection[0].ShapeLength - nutCollection[1].ShapeLength));
                    AddNut(nutCollection[1], matrix, m_PhysicalAspect.Outputs, "NutR2");

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1));
                    AddNut(nutCollection[2], matrix, m_PhysicalAspect.Outputs, "NutR3");

                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -ubolt.UBoltWidth / 2, -pipeOD / 2.0 - gap1 + nutCollection[2].ShapeLength));
                    AddNut(nutCollection[3], matrix, m_PhysicalAspect.Outputs, "NutR4");
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
                //Strap
                matrix.SetIdentity();
                AddStrap(strap, pipeOD, matrix, m_PhysicalAspect.Outputs, "Strap");

                if (pipeToSteel < minPipeToSteel || pipeToSteel > maxPipeToSteel)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidPipeToSteel, "Distance from Pipe to Steel (PipeToSteel) is outside the acceptable range"));
                //Full Circle Wrap
                if (diameter3 > 0 && length3 > 0)
                {
                    matrix = new Matrix4X4();
                    Vector normal = new Vector(0, 0, 1);
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
                    matrix.Translate(new Vector(-length3 / 2, 0, 0));
                    Projection3d wrap = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter3 / 2.0, length3);
                    wrap.Transform(matrix);
                    m_PhysicalAspect.Outputs["Wrap"] = wrap;
                }
            }
            catch//General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of UBolt"));
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

                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of UBolt"));
                }
            }
        }
        #endregion
    }

}




