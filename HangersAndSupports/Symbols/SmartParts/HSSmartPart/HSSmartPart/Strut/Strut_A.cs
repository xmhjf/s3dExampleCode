//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.

//   Struct_A.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strut_A
//   Author       :  Hema
//   Creation Date:  28.May.2013 
//   Description:    Converted StructA Smartpart VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   28.May.2013    Hema   CR-CP-222466 Converted StrutA Smartpart VB Project to C# .Net 
//   22.04.2014     PVK    DM-CP-250839  [TR] Change the existing Strut parts as per customer requirement to 2015
//   30-11=2015     VDP     Integrate the newly developed SmartParts into Product(DI-CP-282644)
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.Generic;

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

    public class Strut_A : SmartPartComponentDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Strut,Ingr.SP3D.Content.Support.Symbols.Strut_A"
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
                AddStrutAInputs(2, out endIndex, additionalInputs);
                additionalInputs.Add((Input)new InputDouble(++endIndex, "MinLen", "MinLen", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "MaxLen", "MaxLen", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "RepOverLen1", "RepOverLen1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "WtPerLen", "WtPerLen", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Weight", "Weight", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Stroke", "Stroke", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Extension", "Extension", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Retraction", "Retraction", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "Allowance", "Allowance", 0, false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "DesignLoad1", "DesignLoad1", 0, false));
                additionalInputs.Add((Input)new InputString(++endIndex, "BOMLenUnits", "BOMLenUnits", "No Value", false));
                additionalInputs.Add((Input)new InputDouble(++endIndex, "SetPosition", "SetPosition", 0, true));
                return additionalInputs;
            }
        }
        #endregion

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddStrutAOutputs(additionalOutputs);
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

                int endIndex, startIndex;

                StrutAInputs strutA = LoadStrutAData(2, out endIndex);

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                startIndex = endIndex;

                Double minLengthValue = GetDoubleInputValue(++startIndex);
                Double maxLengthValue = GetDoubleInputValue(++startIndex);
                Double repOverLength1 = GetDoubleInputValue(++startIndex);
                Double wtPerLength = GetDoubleInputValue(++startIndex);
                Double weight = GetDoubleInputValue(++startIndex);
                Double stroke = GetDoubleInputValue(++startIndex);
                Double extension = GetDoubleInputValue(++startIndex);
                Double retraction = GetDoubleInputValue(++startIndex);
                Double allowance = GetDoubleInputValue(++startIndex);
                Double designLoad1 = GetDoubleInputValue(++startIndex);
                Double setposition = GetDoubleInputValue(++startIndex+1);
                double eyelength = 0;

                if ((extension + retraction) * (1 + allowance) > stroke)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidTravelAllowance, "Travel excess allowance"));

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                strutA.angle1 = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Angle, strutA.angle1, UnitName.ANGLE_DEGREE);

                 if (strutA.rodEndType == 2)
                {
                    eyelength = strutA.diameter1 / 2;
                }



                //ports
                if (strutA.rodEndType == 1 || strutA.rodEndType == 2 || strutA.rodEndType == 3 || strutA.rodEndType == 4)
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Port1", new Position(0, 0, 0), new Vector(Math.Cos(strutA.angle1 * Math.PI / 180), Math.Sin(strutA.angle1 * Math.PI / 180), 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port1"] = port1;
                }
                else if (strutA.rodEndType == 5)   // Nut (Female) end
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Port1", new Position(0, 0, 0), new Vector(-Math.Cos(strutA.angle1) * Math.PI / 180, Math.Sin(strutA.angle1) * Math.PI / 180, 0), new Vector(0, 0, -1));
                    m_PhysicalAspect.Outputs["Port1"] = port1;
                }

                if (setposition>stroke)
                {
                    setposition = stroke;
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidSetposition , "Setposition can not be greater than stroke of " + stroke * 1000 + "mm" + ". Resetting it to Maximum Setpostion"));
                    SupportComponent supportComponnet = (SupportComponent)Occurrence;
                    supportComponnet.SetPropertyValue(setposition, "IJOAHgrSetPosition", "SetPosition");
                }

                //add the graphic
                AddStrutA(ref strutA, new Matrix4X4(), m_PhysicalAspect.Outputs, "Strut", setposition);
                if (part.SupportsInterface("IJOAHgrSetPosition"))
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Port2", new Position(0, 0, strutA.length1 + strutA.shape1.ShapeLength + strutA.shape2.ShapeLength + strutA.shape3.ShapeLength + strutA.shape4.ShapeLength  + strutA.shape5.ShapeLength + eyelength), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_PhysicalAspect.Outputs["Port2"] = port2;
                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Port2", new Position(0, 0, strutA.length), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_PhysicalAspect.Outputs["Port2"] = port2;
                }

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Strut_A"));
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double lengthValue, repOverLength1,minlen,maxlen;

                if (catalogPart.SupportsInterface("IJUAhsLength"))
                    lengthValue = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                else
                lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;

                repOverLength1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsRepOverLen1", "RepOverLen1")).PropValue;
                minlen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMinLen", "MinLen")).PropValue;
                maxlen = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsMaxLen", "MaxLen")).PropValue;
                string length = string.Empty, bomValue = string.Empty;
                try
                {
                    PropertyValueCodelist lBomLengthUnits = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                    bomValue = lBomLengthUnits.PropertyInfo.CodeListInfo.GetCodelistItem(lBomLengthUnits.PropValue).DisplayName;
                }
                catch
                {
                    bomValue = "in";
                }

                if (bomValue.ToUpper() == "IN")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_INCH);
                else if (bomValue.ToUpper() == "FT")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_FOOT);
                else if (bomValue.ToUpper() == "MM")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);
                else if (bomValue.ToUpper() == "M")
                    length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_METER);

                if (lengthValue < minlen || lengthValue > maxlen) 
                {
                    string maxLength = string.Empty, minLength = string.Empty ;

                    try
                    {
                        PropertyValueCodelist bomList = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsBOMLenUnits", "BOMLenUnits");
                        bomValue = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).DisplayName;
                    }
                    catch
                    {
                        bomValue = "in";
                    }
                    if (bomValue.ToUpper() == "IN")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_INCH);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_INCH);
                    }
                    else if (bomValue.ToUpper() == "FT")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_FOOT);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_FOOT);
                    }
                    else if (bomValue.ToUpper() == "MM")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_MILLIMETER);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_MILLIMETER);
                    }
                    else if (bomValue.ToUpper() == "M")
                    {
                        minLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, minlen, UnitName.DISTANCE_METER);
                        maxLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, maxlen, UnitName.DISTANCE_METER);
                    }

                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidMinMaxLength, "Length of the strut must be between" + minLength + "and" + maxLength));

                    if (lengthValue < minlen)
                        lengthValue = minlen;
                    if (lengthValue > maxlen)
                        lengthValue = maxlen;

                    try
                    {
                        oSupportOrComponent.SetPropertyValue(lengthValue, "IJOAHgrOccLength", "Length");
                    }
                    catch { }
                }


                lengthValue = lengthValue + repOverLength1;
                bomDescription = catalogPart.PartDescription;

                if (lengthValue > 0)
                    bomDescription = bomDescription + ",Length=" + length;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrStrutABOMDescription, "Error in BOMDescription of Struct_A.cs."));
            }
            return bomDescription;
        }
        #endregion

        #region "ICustomHgrWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                //System WCG Attributes
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ, dLength = 0, length1, varLength = 0, shape1Length, shape2Length, shape3Length, shape4Length, shape5Length, weightPerUnitLength, fixedWeight;

                int stretchShape = (int)((PropertyValueInt)catalogPart.GetPropertyValue("IJUAhsStretchShape", "StretchShape")).PropValue;
                try
                {
                    length1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                }
                catch
                {
                    length1 = 0;
                }
                try
                {
                    shape1Length = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;
                }
                catch
                {
                    shape1Length = 0;
                }
                shape2Length = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsShape2", "Shape2Length")).PropValue;
                try
                {
                    shape3Length = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsShape3", "Shape3Length")).PropValue;
                }
                catch
                {
                    shape3Length = 0;
                }
                try
                {
                    shape4Length = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsShape4", "Shape4Length")).PropValue;
                }
                catch
                {
                    shape4Length = 0;
                }
                try
                {
                    shape5Length =(double)((PropertyValueDouble) catalogPart.GetPropertyValue("IJUAhsShape5", "Shape5Length")).PropValue;
                }
                catch
                {
                    shape5Length = 0;
                }
                //Custom Part Attributes
                try
                {
                    weightPerUnitLength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWtPerLen", "WtPerLen")).PropValue;
                }
                catch
                {
                    weightPerUnitLength = 0;
                }
                try
                {
                    fixedWeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsWeight", "Weight")).PropValue;
                }
                catch
                {
                    fixedWeight = 0;
                }
                if (stretchShape == 0)  //None
                    varLength = 0;
                else if (stretchShape == 1)  //Shape 1
                    varLength = dLength - length1 - shape2Length - shape3Length - shape4Length - shape5Length;
                else if (stretchShape == 2)  //Shape 2
                    varLength = dLength - length1 - shape1Length - shape3Length - shape4Length - shape5Length;
                else if (stretchShape == 3)  //Shape 3
                    varLength = dLength - length1 - shape1Length - shape2Length - shape4Length - shape5Length;
                else if (stretchShape == 4)   //Shape 4
                    varLength = dLength - length1 - shape1Length - shape2Length - shape3Length - shape5Length;
                else if (stretchShape == 5)  //Shape 5
                    varLength = dLength - length1 - shape1Length - shape2Length - shape3Length - shape4Length;

                //if (catalogPart.SupportsInterface("IJUAhsLength"))
                //    dLength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsLength", "Length")).PropValue;
                //else
                //    dLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;

                //Weight
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = weightPerUnitLength * varLength + fixedWeight;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in weightCG of Struct_A."));
                }
            }
        }
        #endregion
    }
}
