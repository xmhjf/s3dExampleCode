//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HorizontalTraveler.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.HorizontalTraveler
//   Author       :  Havish Garimella
//   Creation Date:  30-Nov-2015
//   DI-CP-282644  Integrate the newly developed SmartParts into Product 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------ 
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
    public class HorizontalTraveler : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Witzenmann,Ingr.SP3D.Content.Support.Symbols.HorizontalTraveler"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("Hole", "Hole")]

        public AspectDefinition m_Symbolic;
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                List<Input> additionalInputs = new List<Input>();

                int startIndex = 2;
                additionalInputs.Add(new InputDouble(startIndex, "Thickness1", "Thickness1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Height1", "Height1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Width1", "Width1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Length1", "Length1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Thickness2", "Thickness2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Height2", "Height2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Width2", "Width2", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Width3", "Width3", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Height3", "Height3", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Thickness3", "Thickness3", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "Offset1", "Offset1", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "HoleDiameter", "HoleDiameter", 0, false));
                additionalInputs.Add(new InputDouble(++startIndex, "SimpShapeType", "SimpShapeType", 2, false));
                additionalInputs.Add(new InputDouble(++startIndex, "LugOffset", "LugOffset", 0, true));

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

                int startindex = 2, endIndex;
                HorizontalTravelerInputs horizontaTraveler = LoadHorizontalTravelerData(startindex, out endIndex);

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                Port Structure = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Structure"] = Structure;
               
                Port Hole = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, -horizontaTraveler.Offset1), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Hole"] = Hole;

                AddHorizontalTraveler(horizontaTraveler, new Matrix4X4(), m_Symbolic.Outputs, "HorizontalTraveler");
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Horizontal Traveler.");
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Horizontal Traveler"));
                }
            }
        }
        #endregion

    }
}
