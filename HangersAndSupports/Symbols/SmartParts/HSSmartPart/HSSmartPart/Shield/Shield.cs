//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Shield.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Shield
//   Author       :  Vijaya
//   Creation Date:  7-Feb-2013
//   Description: CR-CP-222488-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   7-Feb-2013     Vijaya    CR-CP-222488-Initial Creation
//   25/Mar/2013    Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//   30/10/2013     Hema      CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   12-12-2014     PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   10-06-2015     PVK	      TR-CP-274155	TR-CP-273182	GetPropertyValue in HSSmartPart should handle null value exception thrown by CLR
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
    public class Shield : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Shield"
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
        [SymbolOutput("Port3", "Port3")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddShieldInputs(2, out endIndex, additionalInputs);
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
                AddShieldOutputs(additionalOutputs);
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
                ShieldInputs shield = LoadShieldData(2);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                Matrix4X4 matrix = new Matrix4X4();
                Double[] sidePlateXLocation = new Double[1];
                double shieldWeld2Y = 0;
                double shieldWeld2Z = 0;
                int index;
               
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Weld1", new Position(shield.Length1 / 2 + shield.Offset1, 0, -shield.PipeOD / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Port2"] = port2;

                Array.Resize(ref sidePlateXLocation, shield.Multi1Qty);

                if (shield.thickness3 > 0)
                {
                    for (index = 0; index < shield.Multi1Qty; index++)
                    {
                        sidePlateXLocation[index] = MultiPosition(shield.Length1, shield.Multi1Qty, shield.Multi1LocateBy, shield.Multi1Location, shield.Length3)[index];
                    }
                    if (shield.Angle5 > 0)
                    {
                        shieldWeld2Y = Math.Cos(shield.Angle5 / 2) * (shield.PipeOD / 2 + shield.Thickness1);
                        shieldWeld2Z = Math.Sin(shield.Angle5 / 2) * (shield.PipeOD / 2 + shield.Thickness1);
                    }
                    else
                    {
                        shieldWeld2Y = Math.Sqrt(((shield.PipeOD / 2 + shield.Thickness1) * (shield.PipeOD / 2 + shield.Thickness1)) - ((shield.Height1 / 2) * (shield.Height1 / 2)));
                        shieldWeld2Z = shield.Height1 / 2;
                    }
                    Port port3 = new Port(OccurrenceConnection, part, "Weld2", new Position(-shield.Length1 / 2 + sidePlateXLocation[index - 1] + shield.Length3 / 2 + shield.Offset1, shieldWeld2Y, shieldWeld2Z), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_PhysicalAspect.Outputs["Port3"] = port3;
                }
                //call method that make graphics
                AddShield(shield, matrix, m_PhysicalAspect.Outputs, "shield1");
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
                double length1, length2, length3, width1, width2, width3, width4, angle1, angle2, angle3, angle4, angle5, thickness1, thickness2, thickness3, height1, pipeOD, diameter1;
                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                try
                {
                    length1 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                }
                catch
                {
                    length1 = 0;
                }
                try
                {
                    length2 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsLength2", "Length2")).PropValue;
                }
                catch
                {
                    length2 = 0;
                }
                try
                {
                    length3 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsLength3", "Length3")).PropValue;
                }
                catch
                {
                    length3 = 0;
                }
                try
                {
                    width1 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                }
                catch
                {
                    width1 = 0;
                }
                try
                {
                    width2 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsWidth2", "Width2")).PropValue;
                }
                catch
                {
                    width2 = 0;
                }
                try
                {
                    width3 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsWidth3", "Width3")).PropValue;
                }
                catch
                {
                    width3 = 0;
                }
                try
                {
                    width4 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsWidth4", "Width4")).PropValue;
                }
                catch
                {
                    width4 = 0;
                }
                try
                {
                    angle1 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsAngle1", "Angle1")).PropValue;
                }
                catch
                {
                    angle1 = 0;
                }
                try
                {
                    angle2 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsAngle2", "Angle2")).PropValue;
                }
                catch
                {
                    angle2 = 0;
                }
                try
                {
                    angle3 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsAngle3", "Angle3")).PropValue;
                }
                catch
                {
                    angle3 = 0;
                }
                try
                {
                    angle4 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsAngle4", "Angle4")).PropValue;
                }
                catch
                {
                    angle4 = 0;
                }
                try
                {
                    angle5 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsAngle5", "Angle5")).PropValue;
                }
                catch
                {
                    angle5 = 0;
                }
                try
                {
                    thickness1 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsThickness1", "Thickness1")).PropValue;
                }
                catch
                {
                    thickness1 = 0;
                }
                try
                {
                    thickness2 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsThickness2", "Thickness2")).PropValue;
                }
                catch
                {
                    thickness2 = 0;
                }
                try
                {
                    thickness3 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsThickness3", "Thickness3")).PropValue;
                }
                catch
                {
                    thickness3 = 0;
                }
                try
                {
                    height1 = (Double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsHeight1", "Height1")).PropValue;
                }
                catch
                {
                    height1 = 0;
                }
                int multQty = (int)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJUAhsMulti1", "Multi1Qty")).PropValue;
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

                if (diameter1 > pipeOD && diameter1 > 0)
                    pipeOD = diameter1;

                if (HgrCompareDoubleService.cmpdbl(angle1, 0) == true && HgrCompareDoubleService.cmpdbl(angle2, 0) == true)
                {
                    if (width1 < pipeOD / 2 && width1 > 0 && width2 < pipeOD / 2 && width2 > 0)
                    {
                        angle1 = Math.Asin(width1 / (pipeOD / 2));
                        angle2 = Math.Asin(width2 / (pipeOD / 2));
                    }
                }
                if (HgrCompareDoubleService.cmpdbl(angle3, 0) == true && HgrCompareDoubleService.cmpdbl(angle4, 0) == true)
                {
                    if (width3 < pipeOD / 2 && width3 > 0 && width4 < pipeOD / 2 && width4 > 0)
                    {
                        angle3 = Math.Asin(width3 / (pipeOD / 2));
                        angle4 = Math.Asin(width4 / (pipeOD / 2));
                    }
                }
                double upperPlateArea = 0;
                double upperPLVol = 0;
                double weldPLArea = 0;
                double weldPLVol = 0;
                double lowInRadius = pipeOD / 2;
                double lowOutRadius = pipeOD / 2 + thickness1;
                double upperInRadius = pipeOD / 2;
                double upperOutRadius = pipeOD / 2 + thickness2;
                double weldPLOutRadius = lowOutRadius + thickness3;
                double lowerAngle = angle1 + angle2;
                double upperAngle = angle3 + angle4;
                double lowerPlateArea = Math.PI * ((lowOutRadius * lowOutRadius) - (lowInRadius * lowInRadius)) * (lowerAngle / (2 * Math.PI));
                double lowerPLVol = lowerPlateArea * length1;

                if (thickness2 > 0)
                {
                    upperPlateArea = Math.PI * ((upperOutRadius * upperOutRadius) - (upperInRadius * upperInRadius)) * (upperAngle / (2 * Math.PI));
                    upperPLVol = upperPlateArea * length2;
                }

                if (thickness3 > 0)
                {
                    if (length3 <= 0)
                    {
                        if (length1 < length2)
                            length3 = length1;
                        else
                            length3 = length2;
                    }
                    if (angle5 > 0)
                        weldPLArea = Math.PI * ((weldPLOutRadius * weldPLOutRadius) - (lowOutRadius * lowOutRadius)) * (angle5 / (2 * Math.PI));
                    else if (height1 > 0)
                        weldPLArea = height1 * thickness3;

                    weldPLVol = (weldPLArea * length3) * multQty * 2;
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

                totalVolume = upperPLVol + lowerPLVol + weldPLVol;

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