//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Rod.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Rod
//   Author       : Manikanth 
//   Creation Date: 12.2.2013
//   Description  : CR-CP-222487 Creating Rod Class

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12.2.2013    Manikanth  CR-CP-222487 Creating Rod Class
//   25/Mar/2013  Rajeswari  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   11/Aug/2014    Ramya    TR-CP-256377  Additional input values are retrieved from catalog part in smart parts
//   12-12-2014     PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report  
//   10-06-2015     PVK	     TR-CP-274155	TR-CP-273182	GetPropertyValue in HSSmartPart should handle null value exception thrown by CLR
//   30-11-2015     VDP      Integrate the newly developed SmartParts into Product(DI-CP-282644)
//   6/Apr/2016     Ramya    TR 292072  Bug in Weight Calculation for .Net Rod SmartPart  
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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Rod : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Rod"
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
                AddRod1Inputs(2, out endIndex, additionalInputs);
                additionalInputs.Add(new InputDouble(++endIndex, "WtPerLen", "WtPerLen", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "Weight", "Weight", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "MinLen", "MinLen", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "MaxLen", "MaxLen", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "RepOverLen1", "RepOverLen1", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "BOMLenUnits", "BOMLenUnits", 0, false));

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

        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddRod1Outputs(additionalOutputs);
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
                Part part =(Part) m_PartInput.Value;
                int startIndex = 13;
                Rod1Inputs rodInputs = LoadRod1Data(2);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                Double wtPerLen = GetDoubleInputValue(startIndex);
                Double weight = GetDoubleInputValue(++startIndex);
                Double minLen = GetDoubleInputValue(++startIndex);
                Double maxLen = GetDoubleInputValue(++startIndex);
                Double repOverLen1 = GetDoubleInputValue(++startIndex);
                Double bomLenUnits = GetDoubleInputValue(++startIndex);

                string maxLength = "";
                string minLength = "";
                string value="";
                try
                {
                    PropertyValueCodelist bomList = (PropertyValueCodelist)part.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                    value = bomList.PropertyInfo.CodeListInfo.GetCodelistItem((int)bomLenUnits).DisplayName;
                }
                catch
                {
                    value = "in";
                }
                if (value.ToUpper() == "IN")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_INCH);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_INCH);
                }
                else if (value.ToUpper() == "FT")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_FOOT);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_FOOT);
                }
                else if (value.ToUpper() == "MM")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_MILLIMETER);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_MILLIMETER);
                }
                else if (value.ToUpper() == "M")
                {
                    minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minLen, UnitName.DISTANCE_METER);
                    maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxLen, UnitName.DISTANCE_METER);
                }

                if (rodInputs.length < minLen || rodInputs.length > maxLen)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidLengthOfRod, "Lenth of the rod must be between" + "" + minLength + "and" + maxLength));
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "RodEnd1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["RodEnd1"] = port1;
                if (rodInputs.rodEnd2Type == 5)
                {
                    Port port2 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, 0, rodInputs.length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["RodEnd2"] = port2;
                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "RodEnd2", new Position(0, 0, rodInputs.length), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_PhysicalAspect.Outputs["RodEnd2"] = port2;
                }

                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                AddRod(rodInputs, matrix, m_PhysicalAspect.Outputs, "Rod");

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Rod"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double length;
                double rodDiameter;
                string materialType;
                string materialGrade;
                double materialDensity;
                double weightPerUnitlength;
                double additionalWeight;
                double overLength1;
                double overLength2;
                double diameter1;
                double length1;
                double thickness1;
                int rodEnd1Type;
                int rodEnd2Type;
                int rodCenterType;
                double dryWeight;
                double dryCogX;
                double dryCogY;
                double dryCogZ;
                double weight, cogX, cogY, cogZ;

                if (supportComponentBO.SupportsInterface("IJOAHgrOccLength"))
                {
                    try
                    {
                        length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;
                    }
                    catch
                    {
                        length = 0;
                    }
                }
                else if (supportComponentBO.SupportsInterface("IJUAhsLength"))
                {
                    try
                    {
                        length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                    }
                    catch
                    {
                        length = 0;
                    }
                }
		        else if (catalogPart.SupportsInterface("IJOAHgrOccLength"))
                {
                    try
                    {
                        length = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;
                    }
                    catch
                    {
                        length = 0;
                    }
                }
                else if (catalogPart.SupportsInterface("IJUAhsLength"))
                {
                    try
                    {
                        length = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                    }
                    catch
                    {
                        length = 0;
                    }
                }
                else
                    length = 0;
                try
                {

                    overLength1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                }
                catch
                {
                    overLength1 = 0;
                }

                try
                {
                    overLength2 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsOverLength2", "Overlength2")).PropValue;
                }
                catch
                {
                    overLength2 = 0;
                }
                try
                {
                    diameter1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsDiameter1", "Diameter1")).PropValue;
                }
                catch
                {
                    diameter1 = 0;
                }
                try
                {
                    thickness1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsThickness1", "Thickness1")).PropValue;
                }
                catch
                {
                    thickness1 = 0;
                }
                try
                {
                    length1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                }
                catch
                {
                    length1 = 0;
                }

                rodEnd1Type = (int)((PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsRodEnd1", "RodEnd1Type")).PropValue;
                rodEnd2Type = (int)((PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsRodEnd2", "RodEnd2Type")).PropValue;
                rodCenterType = (int)((PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsRodCenterType", "RodCenterType")).PropValue;
                try
                {
                    dryWeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    dryWeight = 0;
                }
                //Center of Gravity
                try
                {
                    dryCogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    dryCogX = 0;
                }
                try
                {
                    dryCogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    dryCogY = 0;
                }
                try
                {
                    dryCogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    dryCogZ = 0;
                }

                
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

                try
                {
                    rodDiameter = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsRodDiameter", "RodDiameter")).PropValue;
                }
                catch
                {
                    rodDiameter = 0;
                }
                try
                {
                    weightPerUnitlength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWtPerLen", "WtPerLen")).PropValue;

                }
                catch
                {
                    weightPerUnitlength = 0;
                }
                try
                {
                    additionalWeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWeight", "Weight")).PropValue;
                }
                catch
                {
                    additionalWeight = 0;
                }
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
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

                if (HgrCompareDoubleService.cmpdbl(dryWeight, 0)==false )
                {
                    weight = dryWeight;
                }
                else if (HgrCompareDoubleService.cmpdbl(weightPerUnitlength, 0)==false && HgrCompareDoubleService.cmpdbl(additionalWeight, 0)==false)
                {
                    weight = weightPerUnitlength * length + additionalWeight;
                }
                else
                {
                    double weight1 = 0;
                    double weight2 = 0;
                    double weight3 = 0;
                    double volume1 = 0;
                    double volume2 = 0;
                    double volume3 = 0;
                    double getPi = (Math.Atan(1) * 4.0);
                    weight = ((getPi * (Math.Pow((rodDiameter / 2), 2) * length)) + (getPi * (Math.Pow((rodDiameter / 2), 2) * overLength1)) + (getPi * (Math.Pow((rodDiameter / 2), 2) * overLength2))) * materialDensity;
                    if (rodEnd1Type > 1)
                    {
                        if (rodEnd1Type == 2)
                        {
                            volume1 = getPi * Math.Pow((diameter1 / 2), 2) * rodDiameter;
                            volume2 = getPi * Math.Pow((rodDiameter / 2), 2) * (diameter1 / 2);
                            weight1 = (volume1 - volume2) * materialDensity;
                        }
                        else if (rodEnd1Type == 3)
                        {
                            volume1 = ((getPi * Math.Pow((diameter1 / 2), 2)) / 2) * thickness1;
                            volume2 = length1 * diameter1 * thickness1;

                            if (rodDiameter > thickness1)
                            {
                                volume3 = rodDiameter * (length + overLength1) * thickness1;
                            }
                            else
                            {
                                volume3 = getPi * Math.Pow((rodDiameter / 2), 2) * (length + overLength1);
                            }
                            weight1 = (volume1 + volume2 - volume3) * materialDensity;
                        }
                        else if (rodEnd1Type == 4 || rodEnd1Type == 5)
                        {
                            volume1 = getPi * Math.Pow((diameter1 / 2), 2) * thickness1;
                            volume2 = getPi * Math.Pow((rodDiameter / 2), 2) * (length + overLength1);
                            weight1 = (volume1 - volume2) * materialDensity;
                        }
                    }
                    if (rodEnd2Type > 1)
                    {
                        if (rodEnd2Type == 2)
                        {
                            volume1 = getPi * (Math.Pow((diameter1 / 2), 2)) * rodDiameter;
                            volume2 = getPi * Math.Pow((rodDiameter / 2), 2) * (diameter1 / 2);
                            weight2 = (volume1 - volume2) * materialDensity;
                        }
                        else if (rodEnd2Type == 3)
                        {
                            volume1 = (getPi * (Math.Pow((diameter1 / 2), 2) / 2)) * thickness1;
                            volume2 = length1 * diameter1 * thickness1;
                            if (rodDiameter > thickness1)
                            {
                                volume3 = rodDiameter * (length + overLength1) * thickness1;
                            }
                            else
                            {
                                volume3 = getPi * (Math.Pow((rodDiameter / 2), 2)) * (length + overLength2);
                            }
                            weight2 = (volume1 + volume2 - volume3) * materialDensity;
                        }
                        else if (rodEnd2Type == 4 || rodEnd2Type == 5)
                        {
                            volume1 = getPi * (Math.Pow((diameter1 / 2), 2)) * thickness1;
                            volume2 = getPi * (Math.Pow((rodDiameter / 2), 2)) * (length + overLength1);
                            weight2 = (volume1 - volume2) * materialDensity;
                        }
                    }
                    if (rodCenterType > 1)
                    {
                        if (rodCenterType == 2)
                        {
                            volume1 = getPi * (Math.Pow((diameter1 / 2), 2)) * length1;
                            volume2 = getPi * (Math.Pow((rodDiameter / 2), 2)) * length1;
                            weight3 = (volume1 - volume2) * materialDensity;
                        }
                        else if (rodCenterType == 3)
                        {
                            volume1 = diameter1 * diameter1 * length1;
                            volume2 = getPi * (Math.Pow((rodDiameter / 2), 2)) * length1;
                            weight3 = (volume1 - volume2) * materialDensity;
                        }
                        else if (rodCenterType == 4)
                        {
                            volume1 = getPi * (Math.Pow((diameter1 / 2), 2)) * rodDiameter;
                            volume2 = getPi * (Math.Pow((rodDiameter / 2), 2)) * diameter1;
                            weight3 = ((2 * volume1) - volume2) * materialDensity;
                        }
                    }

                    weight = weight + weight1 + weight2 + weight3;
                }

                if (HgrCompareDoubleService.cmpdbl(dryCogX, 0)==false )
                {
                    cogX = dryCogX;
                }
                else
                {
                    cogX = 0;
                }
                if (HgrCompareDoubleService.cmpdbl(dryCogY, 0)==false )
                {
                    cogY = dryCogY;
                }
                else
                {
                    cogY = 0;
                }
                if (HgrCompareDoubleService.cmpdbl(dryCogZ, 0)==false )
                {
                    cogZ = dryCogZ;
                }
                else
                {
                    cogZ = length / 2;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }

            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Rod"));
                }
            }
        }
        #endregion
    }

}
