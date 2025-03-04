//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.

//   HgrChannel.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.HgrChannel
//   Author       :  Shilpi	
//   Creation Date:  30.August.2012
//   Description:
//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.August.2012  Shilpi     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   10.Jan.2010     Hema       CR-CP-219113 Changed the implementation of WeightCG
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Support.Middle;


namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Support.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class HgrChannel : ConnectionComponentDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {

        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Support.Content.Symbols.HgrChannel"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "CardinalPoint", "Cardinality Point of the Beam", 1)]
        public InputDouble m_dCardinalPoint;
        [InputDouble(3, "Length", "Length of the Beam", 0.5)]
        public InputDouble m_dLength;
        [InputString(4, "MaterialGrade", "MaterialGrade of the Beam", "A36")]
        public InputString m_dMaterialGrade;
        [InputDouble(5, "Orientation", "Orientation", 10)]
        public InputDouble m_dOrientation;
        [InputDouble(6, "BeginOverLength", "Begin OverLength of the Beam", 0)]
        public InputDouble m_dBeginOverLength;
        [InputDouble(7, "EndOverLength", "End OverLength of the Beam", 0)]
        public InputDouble m_dEndOverLength;
        [InputDouble(8, "BeginMiter", "BeginMiter", 1)]
        public InputDouble m_dBeginMiter;
        [InputDouble(9, "EndMiter", "EndMiter", 1)]
        public InputDouble m_dEndMiter;
        [InputDouble(10, "MaterialType", "MaterialType", 1)]
        public InputDouble m_dMaterialType;
        [InputDouble(11, "HolePattern", "Hole Pattern", 1)]
        public InputDouble m_dHolePattern;
        [InputDouble(12, "Thickness", "Thickness", 1)]
        public InputDouble m_dThickness;

        #endregion
        #region "Definitions of Aspects and their outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BeginCap", "Begin Cap")]
        [SymbolOutput("EndCap", "End Cap")]
        [SymbolOutput("Neutral", "Neutral")]
        [SymbolOutput("BeginCapSurface", "Begin Cap Surface")]
        [SymbolOutput("EndCapSurface", "End Cap Surface")]
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
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                double cardinalX, cardinalY;
                long cardinalPoint = (long)m_dCardinalPoint.Value;
                double length = m_dLength.Value;
                double beginOverLength = m_dBeginOverLength.Value;
                double endOverLength = m_dEndOverLength.Value;

                Part part = (Part)m_PartInput.Value;

               
                //=================================================
                // Construction of Physical Aspect
                //=================================================   

                Part sectionPart = (Part)m_PartInput.Value;

                ReadOnlyCollection<BusinessObject> ports;

                HangerBeamInputs hgrBeamInput = new HangerBeamInputs();
                hgrBeamInput.BeginOverLength = beginOverLength;
                hgrBeamInput.CardinalPoint = 1;
                hgrBeamInput.EndOverLength = endOverLength;
                hgrBeamInput.Length = length;
                hgrBeamInput.Part = sectionPart;
                hgrBeamInput.Density = 0.25; //default value

                ports = CreateConnectionComponentPorts(hgrBeamInput);

                int portCount = ports.Count - 2;

                if (portCount > 0)
                {
                    for (int iIndex = 0; iIndex < portCount; iIndex++)
                        m_PhysicalAspect.Outputs["Port" + iIndex] = ports[iIndex];
                }

                //Now add the Cap Surfaces
                m_PhysicalAspect.Outputs["BeginCapSurface"] = ports[(int)portCount];
                m_PhysicalAspect.Outputs["EndCapSurface"] = ports[(int)portCount + 1];

                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                RelationCollection hgrRelation;
                CrossSection crossSection;
                CrossSectionServices crossSectionServices = new CrossSectionServices();

                hgrRelation = sectionPart.GetRelationship("HgrCrossSection", "CrossSection");
                crossSection = (CrossSection)hgrRelation.TargetObjects.First();

                double width, depth;
                width = crossSection.Width;
                depth = crossSection.Depth;

                //ports
                crossSectionServices.GetCardinalPointOffset(crossSection, (int)cardinalPoint, out cardinalX, out cardinalY);

                double oX = cardinalX;
                double oY = cardinalY;
                double oZ = 0;

                Port beginCap = new Port(OccurrenceConnection, part, "BeginCap", new Position(oX, oY, oZ), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["BeginCap"] = beginCap;
                Port endCap = new Port(OccurrenceConnection, part, "EndCap", new Position(oX, oY, oZ + length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["EndCap"] = endCap;
                Port neutral = new Port(OccurrenceConnection, part, "Neutral", new Position(oX + width / 2, oY + depth / 2, oZ + length / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Neutral"] = neutral;

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HgrChannel."));
                    return;
                }
            }
        }
        #endregion
        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string Description = "";
            try
            {

                Part sectionPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                double length, beginLength, endLength;

                try
                {
                    length = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                }
                catch
                {
                    length = 0;
                }

                try
                {
                    beginLength = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                }
                catch
                {
                    beginLength = 0;
                }

                try
                {
                    endLength = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                }
                catch
                {
                    endLength = 0;
                }

                double cutLength = length + beginLength + endLength;
                string sLength = "";

                Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)oSupportOrComponent.GetRelationship("SupportHasComponents", "Support").TargetObjects[0];
                GenericHelper genericHelper = new GenericHelper(support);
                double unitValue, precision = 0;
                genericHelper.GetDataByRule("HgrStructuralBOMUnits", support, out unitValue);

                if ((UnitName.DISTANCE_METER == (UnitName)unitValue) || ((UnitName)unitValue == UnitName.DISTANCE_MILLIMETER))
                    genericHelper.GetDataByRule("HgrStructuralBOMDecimals", support, out precision);
                if (UnitName.DISTANCE_INCH == (UnitName)unitValue)
                {
                    if (precision > 0)
                        sLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, cutLength, UnitName.DISTANCE_INCH);
                    else
                        sLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, cutLength, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET);
                }
                else
                    sLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, cutLength, (UnitName)unitValue);

                Description = sectionPart.PartDescription + ", Length: " + sLength;
                return Description;

            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of HgrChannel."));
                return "";
            }
        }

        #endregion

        #region ICustomHgrWeightCG Members

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {

                //Get the Cross Section from the Relationship between the HgrPart and the CrossSection
                double weight, cogX, cogY, cogZ;
                CrossSectionServices crossSectionServices = new CrossSectionServices();
                Part sectionPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                RelationCollection hgrRelation = sectionPart.GetRelationship("HgrCrossSection", "CrossSection");
                CrossSection crossSection = (CrossSection)hgrRelation.TargetObjects.First();

                double length, beginLength, endLength;

                try
                {
                    length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                }
                catch
                {
                    length = 0;
                }

                try
                {
                    beginLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                }
                catch
                {
                    beginLength = 0;
                }

                try
                {
                    endLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                }
                catch
                {
                    endLength = 0;
                }

                string channelType = (string)((PropertyValueString)crossSection.GetPropertyValue("IStructCrossSection", "SectionName")).PropValue;

                string channelInterface = "IJOAHgrBLineChannel" + channelType;

                PropertyValueCodelist holePatternCodeList = ((PropertyValueCodelist)supportComponentBO.GetPropertyValue(channelInterface, "HolePattern"));
                long holePatternValue = holePatternCodeList.PropValue;
                if (holePatternValue == -1)
                    holePatternValue = 1;
                if (holePatternValue <= 0 || holePatternValue > 10)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidChHolePattern, "Channel HolePattern Code list value should be between 1 and 10"));

                PropertyValueCodelist materialCodeList = ((PropertyValueCodelist)supportComponentBO.GetPropertyValue(channelInterface, "MaterialType"));
                long material = (long)materialCodeList.PropValue;
                if (material <= 0 || material > 4)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrChMaterial, "Channel Material Code list value should be between 1 and 4"));

                string materialType = materialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)material).DisplayName;
                string holePattern = holePatternCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)holePatternValue).DisplayName;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass auxilaryTable = (PartClass)catalogBaseHelper.GetPartClass("BlineChWeightsAUX");
                ReadOnlyCollection<BusinessObject> classItems = auxilaryTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                double weightPerUnitlength = 0;
                foreach (BusinessObject classItem in classItems)
                {
                    if (classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "ChHolePattern").ToString() == holePattern && classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Material").ToString() == materialType && classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Type").ToString() == channelType)
                        weightPerUnitlength = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Weight_Per_Length")).PropValue;
                }
                double totalLength = (beginLength + length + endLength);
                weight = weightPerUnitlength * totalLength;
                crossSectionServices.GetCardinalPointOffset(crossSection, 10, out cogX, out cogY);
                cogZ = totalLength / 2;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of HgrChannel."));
            }
        }
        #endregion
    }
}