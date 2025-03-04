//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   ThreeHolePort.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ThreeHolePort
//   Author       :  BS
//   Creation Date: 07.Nov.2012 
//   Description:

//   RefData is not avilable.

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.Nov.2012     BS   CR-CP-216332 Initial Creation
//   30-11=2015      VDP  Integrate the newly developed SmartParts into Product(DI-CP-282644)
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]

    public class ThreeHolePort : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ThreeHolePort"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;       
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]        
        public AspectDefinition m_PhysicalAspect;
        #endregion
        
        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddPlateInputs(2, out endIndex, additionalInputs);
                AddHolePortInputs(endIndex, 3, out endIndex, additionalInputs);//+ 1
                return additionalInputs;
            }
        }
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            AddPlateOutputs(aspectName, additionalOutputs);
            AddHolePortOutputs(aspectName, 3, additionalOutputs);
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
                int numberOfHolePorts = 3;
                PlateInputs plate = LoadPlateData(2);
                HolePortInputs[] holePorts = LoadHolePortData(24, numberOfHolePorts);
                Matrix4X4 matrix = new Matrix4X4();
                Double thicknessPos = 0;

                for (int i = 0; i < numberOfHolePorts; i++)
                {
                    if (i > 0)
                        thicknessPos = plate.thickness1;
                    matrix.Rotate(holePorts[i].Rotx, new Vector(1, 0, 0));
                    matrix.Rotate(holePorts[i].Roty, new Vector(0, 1, 0));
                    matrix.Rotate(holePorts[i].Rotz, new Vector(0, 0, 1));
                    //ports
                    Port port1 = new Port(OccurrenceConnection, part, "Hole" + (i + 1), new Position(holePorts[i].PosX, holePorts[i].PosY, plate.thickness1 / 2 + holePorts[i].Size), new Vector(matrix.GetIndexValue(0), matrix.GetIndexValue(2), matrix.GetIndexValue(1)), new Vector(matrix.GetIndexValue(8), matrix.GetIndexValue(10), matrix.GetIndexValue(9)));
                    m_PhysicalAspect.Outputs["Hole" + (i + 1)] = port1;

                }

                matrix.SetIdentity();
                matrix.Translate(new Vector(0, 0, 0));
                AddPlate(plate, matrix, m_PhysicalAspect.Outputs, "Plate");

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ThreeHolePort"));
                    return;
                }
            }
        }
        #endregion

        #region "Caluculating WeightCG"
        public void EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                string materialType, materialGrade;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                VolumeCG volumeCG = supportComponent.GetVolumeAndCOG();
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
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
                double materialDensity, weight;
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
                supportComponent.SetWeightAndCOG(weight, volumeCG.COGX, volumeCG.COGY, volumeCG.COGZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of ThreeHolePort"));
                }
            }
        }
        #endregion
    }
}
