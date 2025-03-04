//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2015, Intergraph Corporation. All rights reserved.
//
//   Bolt.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Bolt
//   Author       : N Devender 
//   Creation Date: 30.11.2015
//   Description  : Creating Bolt Class
//   DI-CP-282644  Integrate the newly developed SmartParts into Product 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-12-2015      VDP     DI-CP-284831  Integrate newly developed Hilti Anchors into the product (Adjust Port Location)
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.Generic;
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
    public class Bolt : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Bolt"
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
                AddBoltSmartpartInputs(2, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }
        #endregion

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("RodEnd1", "RodEnd1")]
        [SymbolOutput("RodEnd2", "RodEnd2")]
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
                Part part =(Part) m_PartInput.Value;
                //int startIndex = 13;
                BoltSmartpartInputs BoltSmartpartInputs = LoadBoltSmartpartData(2);

                if (BoltSmartpartInputs.length - BoltSmartpartInputs.thickness1 < (BoltSmartpartInputs.tWPlateCount + BoltSmartpartInputs.bWPlateCount) * BoltSmartpartInputs.thickness3 + BoltSmartpartInputs.gap1 + BoltSmartpartInputs.thickness2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidGap1OfBoltSmartpart, "Gap1 is more than the correct expected value. Decrease Gap1 or Increase Length of the Bolt "));
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, 0, BoltSmartpartInputs.tWPlateCount*BoltSmartpartInputs.thickness3), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["RodEnd1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, 0, BoltSmartpartInputs.tWPlateCount * BoltSmartpartInputs.thickness3+BoltSmartpartInputs.gap1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["RodEnd2"] = port2;

                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                AddBoltSmarpart(BoltSmartpartInputs, matrix, m_PhysicalAspect.Outputs, "Bolt");

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of BoltSmartpart.");
                }
            }
        }
        #endregion

        #region "Calculating WeightCG"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                VolumeCG volumeCG = supportComponent.GetVolumeAndCOG();
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();

                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                string materialType;
                string materialGrade;
                double materialDensity;
                double weight, cogX, cogY, cogZ;

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

                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = volumeCG.Volume * materialDensity;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = volumeCG.COGX;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = volumeCG.COGY;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = volumeCG.COGZ;
                }

                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }

            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Bolt"));
                }
            }
        }
        #endregion
    }

}
