//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HgrBeam.cs
//    SmartSteel,Ingr.SP3D.Content.Support.Symbols.HgrBeam
//   Author       :  Vijay
//   Creation Date:  29.11.2013
//   Description: 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29.11.2013     Vijay    CR-CP-244141 Convert HS_SmartSteel VB Project to C# .Net
//   10-03-2015     Chethan  TR-CP-269406  Make HgrBeam a Non-Cached Symbol      
//   06-06-2016     Vinay    TR-CP-296065	Fix new coverity issues found in H&S Content  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle.Services;
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

    [SymbolVersion("1.1.0.0")]
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    public class HgrBeam : ConnectionComponentDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "SmartSteel,Ingr.SP3D.Content.Support.Symbols.HgrBeam"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "BeginOverLength", "BeginOverLength", 0.999999)]
        public InputDouble BeginOverLength;
        [InputDouble(3, "EndOverLength", "EndOverLength", 0.999999)]
        public InputDouble EndOverLength;
        [InputDouble(4, "Length", "Length", 0.999999)]
        public InputDouble Length;
        [InputDouble(5, "CP1", "CP1", 1)]
        public InputDouble CP1;
        [InputDouble(6, "CP2", "CP2", 1)]
        public InputDouble CP2;
        [InputDouble(7, "CP3", "CP3", 1)]
        public InputDouble CP3;
        [InputDouble(8, "CP4", "CP4", 1)]
        public InputDouble CP4;
        [InputDouble(9, "CP5", "CP5", 1)]
        public InputDouble CP5;
        [InputDouble(10, "CP6", "CP6", 1)]
        public InputDouble CP6;
        [InputDouble(11, "BeginCapXOffset", "BeginCapXOffset", 0.999999)]
        public InputDouble BeginCapXOffset;
        [InputDouble(12, "BeginCapYOffset", "BeginCapYOffset", 0.999999)]
        public InputDouble BeginCapYOffset;
        [InputDouble(13, "BeginCapRotZ", "BeginCapRotZ", 0.999999)]
        public InputDouble BeginCapRotZ;
        [InputDouble(14, "EndCapXOffset", "EndCapXOffset", 0.999999)]
        public InputDouble EndCapXOffset;
        [InputDouble(15, "EndCapYOffset", "EndCapYOffset", 0.999999)]
        public InputDouble EndCapYOffset;
        [InputDouble(16, "EndCapRotZ", "EndCapRotZ", 0.999999)]
        public InputDouble EndCapRotZ;
        [InputDouble(17, "FlexPortXOffset", "FlexPortXOffset", 0.999999)]
        public InputDouble FlexPortXOffset;
        [InputDouble(18, "FlexPortYOffset", "FlexPortYOffset", 0.999999)]
        public InputDouble FlexPortYOffset;
        [InputDouble(19, "FlexPortZOffset", "FlexPortZOffset", 0.999999)]
        public InputDouble FlexPortZOffset;
        [InputDouble(20, "FlexPortRotX", "FlexPortRotX", 0.999999)]
        public InputDouble FlexPortRotX;
        [InputDouble(21, "FlexPortRotY", "FlexPortRotY", 0.999999)]
        public InputDouble FlexPortRotY;
        [InputDouble(22, "FlexPortRotZ", "FlexPortRotZ", 0.999999)]
        public InputDouble FlexPortRotZ;
        [InputDouble(23, "MaterialGrade", "MaterialGrade", 135)]
        public InputDouble MaterialGrade;
        [InputDouble(24, "MaterialCategory", "MaterialCategory", 110)]
        public InputDouble MaterialCategory;
        [InputDouble(25, "CoatingType", "CoatingType", 3)]
        public InputDouble CoatingType;
        [InputDouble(26, "CoatingRequirement", "CoatingRequirement", 5)]
        public InputDouble CoatingRequirement;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BeginCap", "BeginCap")]
        [SymbolOutput("EndCap", "EndCap")]
        [SymbolOutput("BeginFace", "BeginFace")]
        [SymbolOutput("EndFace", "EndFace")]
        [SymbolOutput("Neutral", "Neutral")]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("BeginCapSurface", "BeginCapSurface")]
        [SymbolOutput("EndCapSurface", "EndCapSurface")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("Port6", "Port6")]
        [SymbolOutput("Port7", "Port7")]
        [SymbolOutput("Port8", "Port8")]
        [SymbolOutput("Port9", "Port9")]
        [SymbolOutput("Port10", "Port10")]
        [SymbolOutput("Port11", "Port11")]
        [SymbolOutput("Port12", "Port12")]
        [SymbolOutput("Port13", "Port13")]
        [SymbolOutput("Port14", "Port14")]
        [SymbolOutput("Port15", "Port15")]
        [SymbolOutput("Port16", "Port16")]
        [SymbolOutput("Port17", "Port17")]
        [SymbolOutput("Port18", "Port18")]
        [SymbolOutput("Port19", "Port19")]
        [SymbolOutput("Port20", "Port20")]
        [SymbolOutput("Port21", "Port21")]
        [SymbolOutput("Port22", "Port22")]
        [SymbolOutput("Port23", "Port23")]
        [SymbolOutput("Port24", "Port24")]
        [SymbolOutput("Port25", "Port25")]
        [SymbolOutput("Port26", "Port26")]
        [SymbolOutput("Port27", "Port27")]
        [SymbolOutput("Port28", "Port28")]
        [SymbolOutput("Port29", "Port29")]
        [SymbolOutput("Port30", "Port30")]
        public AspectDefinition m_Symbolic;

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
                Part part = (Part)PartInput.Value;

                Double beginOverLength = BeginOverLength.Value;
                Double endOverLength = EndOverLength.Value;
                Double length = Length.Value;
                Double beginCapXOffset = BeginCapXOffset.Value;
                Double beginCapYOffset = BeginCapYOffset.Value;
                Double beginCapRotZ = BeginCapRotZ.Value;
                Double endCapXOffset = EndCapXOffset.Value;
                Double endCapYOffset = EndCapYOffset.Value;
                Double endCapRotZ = EndCapRotZ.Value;
                Double flexPortXOffset = FlexPortXOffset.Value;
                Double flexPortYOffset = FlexPortYOffset.Value;
                Double flexPortZOffset = FlexPortZOffset.Value;
                Double flexPortRotX = FlexPortRotX.Value;
                Double flexPortRotY = FlexPortRotY.Value;
                Double flexPortRotZ = FlexPortRotZ.Value;
                long cardinalPt1 = (long)CP1.Value;
                long cardinalPt2 = (long)CP2.Value;
                long cardinalPt3 = (long)CP3.Value;
                long cardinalPt4 = (long)CP4.Value;
                long cardinalPt5 = (long)CP5.Value;
                long cardinalPt6 = (long)CP6.Value;
                Object beginCapSurface, endCapSurface;

                //Route Port Orientation
                Double xDirX6 = 0;
                Double xDirY6 = 0;
                Double xDirZ6 = 0;
                Double zDirX6 = 0;
                Double zDirY6 = 0;
                Double zDirZ6 = 0;

                HangerBeamInputs hgrBeamInput = new HangerBeamInputs();
                hgrBeamInput.BeginOverLength = beginOverLength;
                hgrBeamInput.EndOverLength = endOverLength;
                hgrBeamInput.Length = length;
                hgrBeamInput.CardinalPoint = (int)cardinalPt1;
                hgrBeamInput.Part = part;
                hgrBeamInput.Density = 0.25; //default value
                ReadOnlyCollection<BusinessObject> ports = CreateConnectionComponentPorts(hgrBeamInput);

                int portCount = ports.Count - 2;
                for (int i = 1; i <= portCount; i++)
                {
                    m_Symbolic.Outputs["Port" + i] = ports[i - 1];
                }

                beginCapSurface = ports[portCount];
                endCapSurface = ports[portCount + 1];
                m_Symbolic.Outputs["BeginCapSurface"] = beginCapSurface;
                m_Symbolic.Outputs["EndCapSurface"] = endCapSurface;

                CrossSectionServices crossSectionServices = new CrossSectionServices();
                CrossSection crossSection;
                try
                {
                    crossSection = (CrossSection)part.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartSteelLocalizer.GetString(SmartStellResourceIDs.ErrCrossSectionNotFound, "Could not get Cross-section object."));
                    return;
                }
                double cpOffsetX, cpOffsetY;

                //BeginCap Port
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt1, out cpOffsetX, out cpOffsetY);
                Double X = cpOffsetX + beginCapXOffset, Y = cpOffsetY + beginCapYOffset, Z = 0.0;

                Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(X, Y, Z), new Vector(1, 0, 0), new Vector(0, 0, 1));

                if (beginCapRotZ != 0)
                {
                    port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(X, Y, Z), new Vector(Math.Cos(beginCapRotZ), Math.Sin(beginCapRotZ), 0), new Vector(0, 0, 1));
                }
                m_Symbolic.Outputs["BeginCap"] = port1;

                //EndCap Port
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt2, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX + endCapXOffset;
                Y = cpOffsetY + endCapXOffset;
                Z = 0.0;

                Port port2 = new Port(OccurrenceConnection, part, "EndCap", new Position(X, Y, Z + length), new Vector(1, 0, 0), new Vector(0, 0, 1));

                if (endCapRotZ != 0)
                {
                    port2 = new Port(OccurrenceConnection, part, "EndCap", new Position(X, Y, Z + length), new Vector(Math.Cos(endCapRotZ), Math.Sin(endCapRotZ), 0), new Vector(0, 0, 1));
                }
                m_Symbolic.Outputs["EndCap"] = port2;

                //BeginFace Port
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt3, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX;
                Y = cpOffsetY;
                Z = 0.0;

                Port port3 = new Port(OccurrenceConnection, part, "BeginFace", new Position(X, Y, Z - beginOverLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["BeginFace"] = port3;

                //EndFace Port
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt4, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX;
                Y = cpOffsetY;
                Z = 0.0;

                Port port4 = new Port(OccurrenceConnection, part, "EndFace", new Position(X, Y, Z + endOverLength + length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["EndFace"] = port4;

                //Neutral Port
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt5, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX;
                Y = cpOffsetY;
                Z = length / 2;

                Port port5 = new Port(OccurrenceConnection, part, "Neutral", new Position(X, Y, Z), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Neutral"] = port5;

                //Route Port
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPt6, out cpOffsetX, out cpOffsetY);

                X = cpOffsetX + flexPortXOffset;
                Y = cpOffsetY + flexPortYOffset;
                Z = length + flexPortZOffset;

                if (flexPortRotX == 0 && flexPortRotY == 0 && flexPortRotZ == 0)
                {
                    xDirX6 = 1;
                    xDirY6 = 0;
                    xDirZ6 = 0;
                    zDirX6 = 0;
                    zDirY6 = 0;
                    zDirZ6 = 1;
                }
                else
                {
                    Matrix4X4 tempMatrix = new Matrix4X4();
                    Vector tempVector = new Vector();
                    tempMatrix.SetIdentity();

                    if (flexPortRotX != 0)
                    {
                        tempVector = new Vector(1, 0, 0);
                        tempMatrix.Rotate(flexPortRotX, tempVector);
                    }

                    if (flexPortRotY != 0)
                    {
                        tempVector = new Vector(0, 1, 0);
                        tempMatrix.Rotate(flexPortRotY, tempVector);
                    }

                    if (flexPortRotZ != 0)
                    {
                        tempVector = new Vector(0, 0, 1);
                        tempMatrix.Rotate(flexPortRotZ, tempVector);
                    }

                    xDirX6 = tempMatrix.GetIndexValue(0);
                    xDirY6 = tempMatrix.GetIndexValue(1);
                    xDirZ6 = tempMatrix.GetIndexValue(2);
                    zDirX6 = tempMatrix.GetIndexValue(8);
                    zDirY6 = tempMatrix.GetIndexValue(9);
                    zDirZ6 = tempMatrix.GetIndexValue(10);
                }

                Port port6 = new Port(OccurrenceConnection, part, "Route", new Position(X, Y, Z), new Vector(xDirX6, xDirY6, xDirZ6), new Vector(zDirX6, zDirY6, zDirZ6));
                m_Symbolic.Outputs["Route"] = port6;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartSteelLocalizer.GetString(SmartStellResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HgrBeam.cs."));
                    return;
                }
            }
        }

        #endregion


        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double LValue = 0, beginOverLength = 0, endOverLength = 0;
                String lengthValue; object length = null;
                if (oSupportOrComponent.SupportsInterface("IJUAHgrOccLength"))
                {
                    try
                    {
                        LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                    }
                    catch (InvalidOperationException)
                    {
                        LValue = 0;
                    }
                }

                if (oSupportOrComponent.SupportsInterface("IJUAHgrOccOverLength"))
                {
                    try
                    {
                        beginOverLength = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                    }
                    catch (InvalidOperationException)
                    {
                        beginOverLength = 0;
                    }
                }
                if (oSupportOrComponent.SupportsInterface("IJUAHgrOccOverLength"))
                {
                    try
                    {
                        endOverLength = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                    }
                    catch (InvalidOperationException)
                    {
                        endOverLength = 0;
                    }
                }
                double totalLength = beginOverLength + LValue + endOverLength;
                UnitName unit = UnitName.DISTANCE_METER;
                object iDecimals = 0;

                RelationCollection hgrRelation = oSupportOrComponent.GetRelationship("SupportHasComponents", "Support");
                BusinessObject businessObject = hgrRelation.TargetObjects[0];
                GenericHelper genericHelper = new GenericHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                Collection<object> collection = new Collection<object>();

                UOMFormat uomFormat = MiddleServiceProvider.UOMMgr.GetDefaultUnitFormat(UnitType.Distance);

                genericHelper.GetDataByRule("HgrStructuralBOMUnits", businessObject, out collection);
                if (collection != null)
                    unit = (UnitName)collection[0];
                if (unit == UnitName.DISTANCE_METER || unit == UnitName.DISTANCE_MILLIMETER)
                {
                    try
                    {
                        genericHelper.GetDataByRule("HgrStructuralBOMDecimals", businessObject, out collection);
                        if (collection != null)
                            iDecimals = collection[0];
                    }
                    catch
                    {
                        iDecimals = 2;
                    }
                    uomFormat.PrecisionType = PrecisionType.PRECISIONTYPE_DECIMAL;
                    uomFormat.DecimalPrecision = (short)iDecimals;
                    uomFormat.UnitsDisplayed = true;
                    if (unit == UnitName.DISTANCE_METER)
                        lengthValue = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, totalLength, uomFormat, UnitName.DISTANCE_METER);
                    else
                        lengthValue = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, totalLength, uomFormat, UnitName.DISTANCE_MILLIMETER);
                    length = lengthValue;
                }
                else
                {
                    uomFormat.PrecisionType = PrecisionType.PRECISIONTYPE_FRACTIONAL;
                    uomFormat.FractionalPrecision = 16;
                    uomFormat.UnitsDisplayed = true;
                    uomFormat.ReduceFraction = true;
                    if (unit == UnitName.DISTANCE_INCH || unit == UnitName.DISTANCE_FOOT)
                    {
                        if (totalLength > 0.305)
                            lengthValue = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, totalLength, uomFormat, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET);
                        else
                            lengthValue = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, totalLength, uomFormat, UnitName.DISTANCE_INCH);
                        length = lengthValue.Split('.').GetValue(0);
                    }
                }

                bomDescription = part.PartDescription + ", Length: " + length;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartSteelLocalizer.GetString(SmartStellResourceIDs.ErrBOMDescription, "Error in BOMDescription of HgrBeam.cs."));
                return "";
            }
        }

        #endregion

        #region ICustomWeightCG Members

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                double length = 0, beginOverLength = 0, endOverLength = 0, weight, cogX, cogY, cogZ, X, Y;
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                if (supportComponentBO.SupportsInterface("IJUAHgrOccLength"))
                {
                    try
                    {
                        length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                    }
                    catch (InvalidOperationException)
                    {
                        length = 0;
                    }
                }

                if (supportComponentBO.SupportsInterface("IJUAHgrOccOverLength"))
                {
                    try
                    {
                        beginOverLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                    }
                    catch (InvalidOperationException)
                    {
                        beginOverLength = 0;
                    }
                }
                if (supportComponentBO.SupportsInterface("IJUAHgrOccOverLength"))
                {
                    try
                    {
                        endOverLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                    }
                    catch (InvalidOperationException)
                    {
                        endOverLength = 0;
                    }
                }
                double totalLength = beginOverLength + length + endOverLength;
                CrossSection crossSection;
                try
                {
                    crossSection = (CrossSection)catalogPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects.First();
                }
                catch
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartSteelLocalizer.GetString(SmartStellResourceIDs.ErrCrossSectionNotFound, "Could not get Cross-section object."));
                    return;
                }
                CrossSectionServices crossSectionServices = new CrossSectionServices();
                crossSectionServices.GetCardinalPointOffset(crossSection, 10, out X, out Y);
                double varWeight = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionUnitWeight", "UnitWeight")).PropValue;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = varWeight * totalLength;
                }

                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = X;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = Y;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = totalLength / 2.0;
                }
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartSteelLocalizer.GetString(SmartStellResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of HgrBeam.cs."));
            }
        }

        #endregion
    }
}
