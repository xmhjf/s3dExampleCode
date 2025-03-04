//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   ElbowRPad.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ElbowRPad
//   Author       :  RCM
//   Creation Date:  28-Apr-2014
//   Description: Elliptical Reinforcement Pad for Elbow

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   28/04/2014     RCM     CR-CP-252368 Elliptical Repad Parts for Straight Pipe and Elbow 
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
    public class ElbowRPad : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ElbowRPad"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        //Ports as Outputs
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddElbowRPadInputs(2, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddElbowRPadOutputs("RPad", additionalOutputs);
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
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                ElbowRPadInputs rpad = LoadElbowRPadData(2);

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                // Create the ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = port1;


                double weldPortAngle;
                double padLengthAngle;

                padLengthAngle = rpad.Length/(rpad.ElbowRadius + rpad.PipeOD/2 + rpad.Thickness);
                weldPortAngle = (rpad.BendAngle - padLengthAngle) / 2;

                double weldPortX;
                double weldPortZ;

                weldPortX = (rpad.ElbowRadius + rpad.PipeOD / 2) * Math.Cos(weldPortAngle) - rpad.ElbowRadius;
                weldPortZ = (rpad.ElbowRadius + rpad.PipeOD / 2) * Math.Sin(weldPortAngle) - (rpad.ElbowRadius) * Math.Tan(rpad.BendAngle / 2);

                Port port2 = new Port(OccurrenceConnection, part, "Weld1", new Position(weldPortX, 0, weldPortZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = port2;

                // Add the Pipe R-Pad Geometry Outputs
                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                AddElbowRPad(rpad, matrix, m_PhysicalAspect.Outputs, "RPad");
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ElbowRPad"));
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
                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double length;
                Double width;
                Double thinkness;

                if (supportComponentBO.SupportsInterface("IJUAhsLength1"))
                {
                    length = (Double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                }
                else
                {
                    length = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                }

                if (supportComponentBO.SupportsInterface("IJUAhsWidth1"))
                {
                    width = (Double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                }
                else
                {
                    width = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                }

                if (supportComponentBO.SupportsInterface("IJUAhsThickness1"))
                {
                    thinkness = (Double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsThickness1", "Thickness1")).PropValue;
                }
                else
                {
                    thinkness = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsThickness1", "Thickness1")).PropValue;
                }

                Double totalVolume;
                Double materialDensity;

                String materialType;
                String materialGrade;

                //Custom Part Attributes
                if (supportComponentBO.SupportsInterface("IJOAhsMaterialEx"))
                {
                    materialType = ((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue;
                    materialGrade = ((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue;
                }
                else
                {
                    materialType = ((PropertyValueString)CatalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue;
                    materialGrade = ((PropertyValueString)CatalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue;
                }

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

                totalVolume = Math.PI * length / 2 * width / 2 * thinkness;

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Elbow R-Pad"));
                }
            }
        }
        #endregion
    }
}