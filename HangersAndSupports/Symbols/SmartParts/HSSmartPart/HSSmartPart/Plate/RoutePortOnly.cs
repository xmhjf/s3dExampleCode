//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   RoutePortOnly.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.RoutePortOnly
//   Author       :  BS
//   Creation Date: 07.Nov.2012 
//   Description:

//   RefData is not avilable.

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07.Nov.2012     BS   CR-CP-216332 Initial Creation
//   11.Aug.2014     Ramya   TR-CP-256377  Additional input values are retrieved from catalog part in smart parts  
//   12-12-2014      PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   30-11=2015      VDP     Integrate the newly developed SmartParts into Product(DI-CP-282644)
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
    public class RoutePortOnly : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.RoutePortOnly"
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
                AddRoutePortInputs(endIndex, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            AddPlateOutputs(aspectName, additionalOutputs);
            AddRoutePortOutputs(aspectName, additionalOutputs);

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
                PlateInputs plate = LoadPlateData(2);
                RoutePortInputs routePort = LoadRoutePortData(24);
                Matrix4X4 matrix = new Matrix4X4();
                Double actualOffsetFromCL, portAdjustment;

                PropertyValueCodelist propCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJUAhsRoutePort", "RPAdjustPlateLength");
                CodelistItem codeList = propCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)routePort.RPAdjustPlateLength);
                String adjustPlateLength = codeList.DisplayName;

                //The route port is always WITH BOTTOM EDGE of the plate


                if (HgrCompareDoubleService.cmpdbl(routePort.RPOffsetfromPipeCL, 0) == false)
                {
                    if (routePort.RPOffsetfromPipeCL < 0)
                        actualOffsetFromCL = routePort.RPOffsetfromPipeCL + plate.thickness1;
                    else
                        actualOffsetFromCL = routePort.RPOffsetfromPipeCL;
                    if (actualOffsetFromCL > routePort.RPDiameter / 2)
                        portAdjustment = routePort.RPDiameter / 2;
                    else
                        portAdjustment = routePort.RPDiameter / 2 - Math.Sqrt(Math.Abs(0.5 * routePort.RPDiameter * 0.5 * routePort.RPDiameter - routePort.RPOffsetfromPipeCL * routePort.RPOffsetfromPipeCL));
                }
                else
                {
                    actualOffsetFromCL = plate.thickness1 / 2;
                    portAdjustment = 0;
                }
                routePort.XOffset = plate.width1 / 2 + routePort.RPOffsetAlongPipeCL;
                routePort.YOffset = -routePort.RPDiameter / 2 + portAdjustment;
                routePort.ZOffset = actualOffsetFromCL;
                if (HgrCompareDoubleService.cmpdbl(routePort.RPOffsetAlongPipeCL, 0) == false)
                    if (adjustPlateLength.Equals("YES", StringComparison.OrdinalIgnoreCase))
                        plate.length1 = plate.length1 + portAdjustment;
                Port routePort1 = null;
                if (HgrCompareDoubleService.cmpdbl(routePort.RPRotationAroundPipe, 0) == false)
                    routePort1 = new Port(OccurrenceConnection, part, "Route", new Position(routePort.XOffset, routePort.YOffset, routePort.ZOffset), new Vector(1, 0, 0), new Vector(0, Math.Cos(routePort.RPRotationAroundPipe) / (Math.Pow(Math.Cos(routePort.RPRotationAroundPipe), 2) + Math.Pow(Math.Sin(routePort.RPRotationAroundPipe), 2)), Math.Sin(routePort.RPRotationAroundPipe) / (Math.Pow(Math.Cos(routePort.RPRotationAroundPipe), 2) + Math.Pow(Math.Sin(routePort.RPRotationAroundPipe), 2))));
                else
                    routePort1 = new Port(OccurrenceConnection, part, "Route", new Position(routePort.XOffset, routePort.YOffset, routePort.ZOffset), new Vector(1, 0, 0), new Vector(0, 1, 0));

                matrix.SetIdentity();
                matrix.Translate(new Vector(0, 0, 0));
                AddPlate(plate, matrix, m_PhysicalAspect.Outputs, "Plate");

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of RouteportOnly"));
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
                double materialDensity,weight;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of RouteportOnly"));
                }
            }
        }
        #endregion
    }
}
