//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   TwoHolePort.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.TwoHolePort
//   Author       :  BS
//   Creation Date: 07.Nov.2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.Nov.2012     BS      CR-CP-216332 Initial Creation
//	 07.Sep.2015     Vinay	 DI-CP-279038	Replace Anvil parts used in HS_Assembly with Anvil2010 Parts
//   30-11=2015      VDP     Integrate the newly developed SmartParts into Product(DI-CP-282644)
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System.Collections.Generic;
using System;
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

    public class TwoHolePort : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.TwoHolePort"
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
                AddHolePortInputs(endIndex, 2, out endIndex, additionalInputs);//+ 1
                additionalInputs.Add(new InputDouble(endIndex, "CurvedEndType", "CurvedEndType", 1, true));
                return additionalInputs;
            }
        }
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            AddPlateOutputs(aspectName, additionalOutputs);
            AddHolePortOutputs(aspectName, 2, additionalOutputs);
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
                int numberOfHolePorts = 2;
                PlateInputs plate = LoadPlateData(2);
                HolePortInputs[] holePorts = LoadHolePortData(24, numberOfHolePorts);
                if (part.SupportsInterface("IJUAhsCurvedEndType"))
                    plate.curvedEndType = (int)GetDoubleInputValue(36);
                Matrix4X4 matrix = new Matrix4X4();

                for (int i = 0; i < numberOfHolePorts; i++)
                {

                    Vector xaxis = new Vector();
                    Vector zaxis = new Vector();
                    HoleRotatePort(holePorts[i], xaxis, zaxis);
                    Port port1 = new Port(OccurrenceConnection, part, "Hole" + (i + 1), new Position(holePorts[i].PosX, holePorts[i].PosY, plate.thickness1 / 2 + holePorts[i].Size), new Vector(xaxis.X, xaxis.Z, xaxis.Y), new Vector(zaxis.X, zaxis.Z, zaxis.Y));
                    m_PhysicalAspect.Outputs["Hole" + (i + 1)] = port1;

                }
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, 0));
                AddPlate(plate, matrix, m_PhysicalAspect.Outputs, "Plate");

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of TwoHolePort"));
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
                
                //Center of Gravity
                if (double.IsNaN(volumeCG.COGX))
                {
                    try
                    {
                        volumeCG.COGX = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                    }
                    catch
                    {
                        volumeCG.COGX = 0;
                    }
                }
                if (double.IsNaN(volumeCG.COGY))
                {
                    try
                    {
                        volumeCG.COGY = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                    }
                    catch
                    {
                        volumeCG.COGY = 0;
                    }
                }
                if (double.IsNaN(volumeCG.COGZ))
                {
                    try
                    {
                        volumeCG.COGZ = (double)((PropertyValueDouble)supportComponent.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                    }
                    catch
                    {
                        volumeCG.COGZ = 0;
                    }
                }
                supportComponent.SetWeightAndCOG(weight, volumeCG.COGX, volumeCG.COGY, volumeCG.COGZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of TwoHolePort"));
                }
            }
        }
        #endregion
    }
}
