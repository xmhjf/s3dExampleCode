//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2015, Intergraph Corporation. All rights reserved.
//
//   Insulation.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Insulation
//   Author       :  JRM
//   Creation Date:  17-Dec-2015
//   Description: CR-CP-281046 New Insulation SmartPart

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   28.01.2016     JRM    CR-CP-281046 New Insulation SmartPart
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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
    public class Insulation : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Shield"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int startIndex = 2;
                List<Input> additionalInputs = new List<Input>();

                //Botttom Dimensions
                additionalInputs.Add(new InputDouble(startIndex, "PipeOD", "PipeOD", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Diameter1", "Diameter1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Length", "Length", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Width1", "Width1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Height1", "Height1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Length1", "Length1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Width2", "Width2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Thickness2", "Thickness2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Offset2", "Offset2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "LayerQty", "LayerQty", 1, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Offset1", "Offset1", 0, false));
                return additionalInputs;
            }
        }
        #endregion
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        //Ports as Outputs
        [SymbolOutput("Port1", "Port1")]
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
                InsulationInputs InsulatInputs = LoadInsulatInputsData(2);

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Matrix4X4 matrix = new Matrix4X4();

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Port1"] = port1;

                //Insualtion Layers
                double layerLength, layerDiam;
                string layerName;
                double outerLayerLength = InsulatInputs.Length;

                for (int i = 1; i <= InsulatInputs.LayerQty; i++)
                {
                    layerLength = InsulatInputs.Length - (2 * (i - 1)) * InsulatInputs.Offset1;
                    layerDiam = InsulatInputs.PipeOD + i * (InsulatInputs.Diameter1 - InsulatInputs.PipeOD) / InsulatInputs.LayerQty;
                    layerName = "InsulationLayer" + i;

                    symbolGeometryHelper.ActivePosition = new Position(-layerLength / 2, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, -1));
                    Projection3d InsulationLayer = symbolGeometryHelper.CreateCylinder(null, layerDiam / 2, layerLength);
                    m_PhysicalAspect.Outputs[layerName] = InsulationLayer;

                    if (i == InsulatInputs.LayerQty)
                    {
                        outerLayerLength = layerLength;
                    }

                }

                //Strap Graphics
                if (InsulatInputs.Width2 > 0)
                {
                    symbolGeometryHelper.ActivePosition = new Position(-outerLayerLength / 2 + InsulatInputs.Offset2, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, -1));
                    Projection3d Strap1 = symbolGeometryHelper.CreateCylinder(null, (InsulatInputs.Diameter1 + 2 * InsulatInputs.Thickness2) / 2, InsulatInputs.Width2);
                    m_PhysicalAspect.Outputs["Strap1"] = Strap1;

                    symbolGeometryHelper.ActivePosition = new Position(outerLayerLength / 2 - InsulatInputs.Width2 - InsulatInputs.Offset2, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, -1));
                    Projection3d Strap2 = symbolGeometryHelper.CreateCylinder(null, (InsulatInputs.Diameter1 + 2 * InsulatInputs.Thickness2) / 2, InsulatInputs.Width2);
                    m_PhysicalAspect.Outputs["Strap2"] = Strap2;
                }

                //Cradle Graphics
                if (InsulatInputs.Width1 > 0)
                {
                    double cradleLength = 0;
                    if (InsulatInputs.Length1 == 0)
                    {
                        cradleLength = InsulatInputs.Length;
                    }
                    else
                    {
                        cradleLength = InsulatInputs.Length1;
                    }


                    if (InsulatInputs.Height1 == 0)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Height 1 is required for the cradle graphics."));
                        return;
                    }

                    symbolGeometryHelper.ActivePosition = new Position(-cradleLength / 2, 0, InsulatInputs.Height1 / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, -1));
                    Projection3d Cradle = symbolGeometryHelper.CreateBox(null, cradleLength, InsulatInputs.Height1, InsulatInputs.Width1);
                    m_PhysicalAspect.Outputs["Cradle"] = Cradle;
                }

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Shield"));
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
                //System WCG Attributes
                double length, pipeOD, diameter1;
                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                try
                {
                    length = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                }
                catch
                {
                    length = 0;
                }
                try
                {
                    pipeOD = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJOAhsPipeOD", "PipeOD")).PropValue;
                }
                catch
                {
                    pipeOD = 0;
                }
                try
                {
                    diameter1 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsDiameter1", "Diameter1")).PropValue;
                }
                catch
                {
                    diameter1 = 0;
                }

                Double totalVolume;
                Double materialDensity;

                String materialType;
                String materialGrade;

                //Custom Part Attributes
                materialType = ((PropertyValueString)CatalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue;
                materialGrade = ((PropertyValueString)CatalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue;

                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                Material material;
                material = catalogStructHelper.GetMaterial(materialType, materialGrade);

                try
                {
                    materialDensity = material.Density;
                }
                catch
                {
                    materialDensity = 0.0;
                }

                totalVolume = ((Math.PI * Math.Pow(diameter1 / 2, 2)) - (Math.PI * Math.Pow(pipeOD / 2, 2))) * length;

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = totalVolume * materialDensity;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Shield"));
                }
            }
        }
        #endregion
    }
}