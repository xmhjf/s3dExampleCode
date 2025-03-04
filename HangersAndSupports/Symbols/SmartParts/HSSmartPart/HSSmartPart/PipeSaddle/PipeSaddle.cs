//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PipeSaddle.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PipeSaddle
//   Author       :  Havish Garimella
//   Creation Date:  30-Nov-2015
//   DI-CP-282644  Integrate the newly developed SmartParts into Product 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------ 
//   18.12.2015      VDP     DI-CP-285917 	Rename PipeStanchion SmartPart to PipeSaddle 
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
    public class PipeSaddle : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PipeSaddle"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion


         #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Steel", "Steel")]

        public AspectDefinition m_Symbolic;
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddPipeSaddleInputs(1, out endIndex, additionalInputs);
                return additionalInputs;
            }
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

                PipeSaddleInputs PipeSaddle = LoadPipeSaddleData(1);

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Route"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, PipeSaddle.Height), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Steel"] = port2;

                Matrix4X4 PipeSaddleMatrix = new Matrix4X4();
                PipeSaddleMatrix.Origin = new Position(0, 0, 0);
                AddPipeSaddle(PipeSaddle, PipeSaddleMatrix, m_Symbolic.Outputs, "Pipe Stanchion");
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Pipe Saddle.");
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Pipe Saddle"));
                }
            }
        }
        #endregion
    }
}
