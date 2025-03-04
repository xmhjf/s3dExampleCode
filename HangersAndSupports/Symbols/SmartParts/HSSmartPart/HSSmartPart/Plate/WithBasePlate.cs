//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   WithBasePlate.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.WithBasePlate
//   Author       :  
//   Creation Date: 07.Nov.2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.Nov.2012     BS   CR-CP-216332 Initial Creation
//   31/10/2013     Rajeswari    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   30-11=2015      VDP         Integrate the newly developed SmartParts into Product(DI-CP-282644)
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
    public class WithBasePlate : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.WithBasePlate"
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
                AddStandardPortInputs(endIndex, 1, out endIndex, additionalInputs);
                AddHolePortInputs(endIndex, 1, out endIndex, additionalInputs);//+1 replace
                additionalInputs.Add(new InputDouble(endIndex, "Width2", "Width2", 0, false));
                additionalInputs.Add(new InputDouble(endIndex + 1, "Length2", "Length2", 0, false));//+ 1
                additionalInputs.Add(new InputDouble(endIndex + 2, "Thickness2", "Thickness2", 0, false));//+ 2
                return additionalInputs;
            }
        }
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            AddPlateOutputs(aspectName, additionalOutputs);
            AddStandardPortOutputs(aspectName, 1, additionalOutputs);
            AddHolePortOutputs(aspectName, 1, additionalOutputs);
            additionalOutputs.Add(new OutputDefinition("BasePlate", "BasePlate"));

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
                int numberOfStandardPorts = 1;
                PlateInputs plate = LoadPlateData(2);
                StandardPortInputs[] standardports = LoadStandardPortData(24, numberOfStandardPorts);
                Matrix4X4 matrix = new Matrix4X4();
                Double thicknessPos = 0;

                for (int i = 0; i < numberOfStandardPorts; i++)
                {
                    if (i > 0)
                        thicknessPos = plate.thickness1;
                    Double[] dPortDir = new Double[6];
                    Vector xaxis = new Vector();
                    Vector zaxis = new Vector();
                    RotatePort(standardports[i], xaxis, zaxis);
                    Port port1 = new Port(OccurrenceConnection, part, "Port" + (i + 1), new Position(plate.width1 / 2 + standardports[i].XOffset, plate.length1 / 2 + standardports[i].YOffset, thicknessPos + standardports[i].ZOffset), new Vector(xaxis.X, xaxis.Y, xaxis.Z), new Vector(zaxis.X, zaxis.Y, zaxis.Z));
                    m_PhysicalAspect.Outputs["Port" + (i + 1)] = port1;

                }
                int numberOfHolePorts = 1;
                HolePortInputs[] holePorts = LoadHolePortData(30, numberOfHolePorts);
                for (int i = 0; i < numberOfHolePorts; i++)
                {
                    Double[] dPortDir = new Double[6];
                    Vector xaxis = new Vector();
                    Vector zaxis = new Vector();
                    HoleRotatePort(holePorts[i], xaxis, zaxis);

                    Port port1 = new Port(OccurrenceConnection, part, "Hole" + (i + 1), new Position(holePorts[i].PosX, holePorts[i].PosY, plate.thickness1 / 2 + holePorts[i].Size), new Vector(xaxis.X, xaxis.Z, xaxis.Y), new Vector(zaxis.X, zaxis.Z, zaxis.Y));
                    m_PhysicalAspect.Outputs["Hole" + (i + 1)] = port1;

                }
                Double Width2 = GetDoubleInputValue(36);
                Double Length2 = GetDoubleInputValue(37);
                Double Thickness2 = GetDoubleInputValue(38);

                SymbolGeometryHelper symbolGeomHlpr = new SymbolGeometryHelper();
                symbolGeomHlpr.ActivePosition = new Position(plate.width1 / 2, -Thickness2, plate.thickness1 / 2);
                symbolGeomHlpr.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                BusinessObject b1 = symbolGeomHlpr.CreateBox(null, Thickness2, Length2, Width2);
                m_PhysicalAspect.Outputs["BasePlate"] = b1;

                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, 0));
                AddPlate(plate, matrix, m_PhysicalAspect.Outputs, "Plate");

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of WithBasePlate"));
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
                double weight;
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
                double materialDensity;
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
                supportComponent.SetWeightAndCOG(weight , volumeCG.COGX, volumeCG.COGY, volumeCG.COGZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of WithBasePlate"));
                }
            }
        }
        #endregion
    }
}
