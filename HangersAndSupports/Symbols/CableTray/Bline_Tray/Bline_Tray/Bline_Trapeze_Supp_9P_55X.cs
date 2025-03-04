//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_Trapeze_Supp_9P_55X.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Trapeze_Supp_9P_55X
//   Author       :  Hema
//   Creation Date:  10.August.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10.August.2012    Hema   CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   26.Nov.2012     Hema     CR-CP-219113 Modified the code with SymbolGeomHelper
//   10.Jan.2010     Hema     CR-CP-219113 Changed the implementation of WeightCG
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System.Collections.ObjectModel;
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
    //<CompanyName>.SP3D.Content.<Specialization>.iiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Bline_Trapeze_Supp_9P_55X : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Trapeze_Supp_9P_55X"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Gap", "Gap between Clamp and Web Rail", 0)]
        public InputDouble m_dGA;
        [InputDouble(3, "ClampLength", "Clamp Length", 0)]
        public InputDouble m_dCL;
        [InputDouble(4, "DistEdge", "Distance from edge to hole", 0)]
        public InputDouble m_dLS;
        [InputDouble(5, "ClampWidth", "Clamp Width", 0)]
        public InputDouble m_dWC;
        [InputDouble(6, "ClampHeight", "Clamp Height", 0)]
        public InputDouble m_dHC;
        [InputDouble(7, "HoleDiameter", "Hole Diameter", 0)]
        public InputDouble m_dDO;
        [InputDouble(8, "TrayWT", "Tray Web Thickness", 0)]
        public InputDouble m_dTW;
        [InputDouble(9, "LC", "Length of Channel", 0)]
        public InputDouble m_dLC;
        [InputDouble(10, "TCCL", "Trapeze Channel Cut Length", 0)]
        public InputDouble m_dCCL;
        [InputDouble(11, "WT", "Tray Width", 0)]
        public InputDouble m_dWT;
        [InputDouble(12, "Inside_Outside", "Inside or Outside", 2)]
        public InputDouble m_dInOrOut;
        [InputDouble(13, "CT", "Thickness of Channel", 0)]
        public InputDouble m_dCT;
        [InputDouble(14, "ChWidth", "Width of Channel", 0)]
        public InputDouble m_dCW;
        [InputDouble(15, "CH", "Height of Channel", 0)]
        public InputDouble m_dCH;
        [InputDouble(16, "WH", "Distance from Hanger Rod to Tray center", 0)]
        public InputDouble m_dWH;
        [InputDouble(17, "Material", "Material", 1)]
        public InputDouble m_dMaterial;
        [InputDouble(18, "HolePattern", "Hole Pattern", 1)]
        public InputDouble m_dHolePattern;
        [InputDouble(19, "Thickness", "Thickness", 1)]
        public InputDouble m_dThickness;
        [InputDouble(20, "Clamp_Guide", "Clamp or Guide", 1)]
        public InputDouble m_dClamp_Guide;

        #endregion

        #region "Definitions of Aspects and their outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("L_CLAMP", "L_CLAMP")]
        [SymbolOutput("R_CLAMP", "R_CLAMP")]
        [SymbolOutput("L_INTHRDRH", "L_INTHRDRH")]
        [SymbolOutput("R_INTHRDRH", "R_INTHRDRH")]
        [SymbolOutput("CHANNEL", "CHANNEL")]

        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

              
                double GA = m_dGA.Value;
                double CL = m_dCL.Value;
                double WC = m_dWC.Value;
                double HC = m_dHC.Value;
                double DO = m_dDO.Value;
                double TW = m_dTW.Value;
                double LC = m_dLC.Value;
                double CCL = m_dCCL.Value;
                double WT = m_dWT.Value;
                double inOrOut = m_dInOrOut.Value;
                double CT = m_dCT.Value;
                double CW = m_dCW.Value;
                double CH = m_dCH.Value;
                double WH = m_dWH.Value;

                if (CCL > LC)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidCCL, "The Channel Cut Length should be smaller than the Length of the Channel. Please check the value"));                   
                if (WH > (CCL / 2))
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidWH, "The distance from hanger rod to cabletray center should be smaller than the half the Channel Cut Length"));
                if (CCL == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidCCLZero, "Channel Cut Length value should equal to 0."));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                if (DO <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDO, "Hole Diameter value should  be greater than 0."));
                if (CL <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidCL, "Clamp Length value should  be greater than 0."));
                if (WC <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidWCgtZero, "Width of Clamp value should  be greater than 0."));
                if (HC <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidHC, "Clamp Height value should  be greater than 0."));

                //Add Ports
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port trayPort = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = trayPort;

                Port inThrdRH1Port = new Port(OccurrenceConnection, part, "InThrdRH1", new Position(0, (CCL - 0.0254 - WH), -CH - 0.008), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = inThrdRH1Port;

                Port inThrdRH2Port = new Port(OccurrenceConnection, part, "InThrdRH2", new Position(0, -WH, -CH - 0.008), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = inThrdRH2Port;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-CW / 2, -(0.5 * 0.0254) - WH, 0));
                pointCollection.Add(new Position(-CW / 2 + CT, -(0.5 * 0.0254) - WH, 0));
                pointCollection.Add(new Position(-CW / 2 + CT, -(0.5 * 0.0254) - WH, -CH + CT));
                pointCollection.Add(new Position(CW / 2 - CT, -(0.5 * 0.0254) - WH, -CH + CT));
                pointCollection.Add(new Position(CW / 2 - CT, -(0.5 * 0.0254) - WH, 0));
                pointCollection.Add(new Position(CW / 2, -(0.5 * 0.0254) - WH, 0));
                pointCollection.Add(new Position(CW / 2, -(0.5 * 0.0254) - WH, -CH));
                pointCollection.Add(new Position(-CW / 2, -(0.5 * 0.0254) - WH, -CH));
                pointCollection.Add(new Position(-CW / 2, -(0.5 * 0.0254) - WH, 0));

                Vector projectionVector = new Vector(0, CCL, 0);
                Projection3d oChannel = new Projection3d( new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["CHANNEL"] = oChannel;

                PropertyValueCodelist insideOutsideCodelistValue = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                if (insideOutsideCodelistValue.PropValue <= 0 || insideOutsideCodelistValue.PropValue > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInside_Outside, "Bline Inside_Outside Code list value should be between 1 and 2"));
                }       

                if (insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem((int)inOrOut).DisplayName.ToLower() == "outside")
                {
                    Vector normal = new Position(0, -WT / 2 - TW - GA - CL / 2, 0).Subtract(new Position(0, -WT / 2 - TW - GA - CL / 2, HC * 1.25));
                    symbolGeometryHelper.ActivePosition = new Position(0, -WT / 2 - TW - GA - CL / 2, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d leftInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal.Length);
                    m_Symbolic.Outputs["L_INTHRDRH"] = leftInthrdrh;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    Vector normal1 = new Position(0, WT / 2 + TW + GA + CL / 2, 0).Subtract(new Position(0, WT / 2 + TW + GA + CL / 2, HC * 1.25));
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW + GA + CL / 2, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d rightInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                    m_Symbolic.Outputs["R_INTHRDRH"] = rightInthrdrh;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -WT / 2 - TW - GA, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                    Projection3d rightClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL, WC, HC);
                    m_Symbolic.Outputs["R_CLAMP"] = rightClamp;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW + GA, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                    Projection3d leftClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL, WC, HC);
                    m_Symbolic.Outputs["L_CLAMP"] = leftClamp;
                }
                else
                {
                    Vector normal = new Position(0, -WT / 2 + TW + GA + CL / 2, 0).Subtract(new Position(0, -WT / 2 + TW + GA + CL / 2, HC * 1.25));
                    symbolGeometryHelper.ActivePosition = new Position(0, -WT / 2 + TW + GA + CL / 2, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d leftInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal.Length);
                    m_Symbolic.Outputs["L_INTHRDRH"] = leftInthrdrh;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    Vector normal1 = new Position(0, WT / 2 - TW - GA - CL / 2, 0).Subtract(new Position(0, WT / 2 - TW - GA - CL / 2, HC * 1.25));
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 - TW - GA - CL / 2, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d rightInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                    m_Symbolic.Outputs["R_INTHRDRH"] = rightInthrdrh;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 - TW - GA, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1,0), new Vector(1, 0, 0));
                    Projection3d leftClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL,WC,HC);
                    m_Symbolic.Outputs["L_CLAMP"] = leftClamp;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -WT / 2 + TW + GA, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                    Projection3d rightClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL, WC, HC);
                    m_Symbolic.Outputs["R_CLAMP"] = rightClamp;
                }
            }
            catch //General Unhandled exception 
            {               
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_Trapeze_Supp_9P_55X."));
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
                string cutLength = "";

                PropertyValueCodelist inside_OutsideCodelist, clamp_GuideCodelist, holePatternCodelist, rod_HardwareCodelist;

                inside_OutsideCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                if (inside_OutsideCodelist.PropValue <= 0 || inside_OutsideCodelist.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInside_Outside, "Bline Inside_Outside Code list value should be between 1 and 2"));
                string inside_Outside = inside_OutsideCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(inside_OutsideCodelist.PropValue).DisplayName;

                clamp_GuideCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineClampGuide", "Clamp_Guide");
                if (clamp_GuideCodelist.PropValue <= 0 || clamp_GuideCodelist.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidClampGuide, "Clamp Guide Code list value should be between 1 and 2"));
                string clamp_Guide = clamp_GuideCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(clamp_GuideCodelist.PropValue).DisplayName;

                holePatternCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineChannelProp", "HolePattern");
                if (holePatternCodelist.PropValue == -1)
                    holePatternCodelist.PropValue = 1;
                if (holePatternCodelist.PropValue <= 0 || holePatternCodelist.PropValue > 10)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidChHolePattern, "Channel HolePattern Code list value should be between 1 and 10"));
                string holePattern = holePatternCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(holePatternCodelist.PropValue).DisplayName;

                rod_HardwareCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineRodHardware", "Rod_Hardware");
                if (rod_HardwareCodelist.PropValue <= 0 || rod_HardwareCodelist.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidRodHardware, "Rod Hardware Code list value should be between 1 and 2"));
                string rod_Hardware = rod_HardwareCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(rod_HardwareCodelist.PropValue).DisplayName;

                double TCCL = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineTrapChCutL", "TCCL")).PropValue;
                cutLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TCCL, UnitName.DISTANCE_INCH);

                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                bomDescription = catalogPart.PartDescription + ", " + clamp_Guide + ": " + inside_Outside + "," + " Cut Length: " + cutLength + ", Hole Pattern: " + holePattern + ", Rod Hardware: " + rod_Hardware;
                return bomDescription;
            }
            catch 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_Trapeze_Supp_9P_55X."));
                return "";
            }
        }
        #endregion

        #region ICustomHgrWeightCG Members

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                double weight, cogX, cogY, cogZ;
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrBlineTrapChCutL", "TCCL")).PropValue;
                double clampWeight = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrBlineTrHWWeight", "Trapeze_HW_Weight")).PropValue;
                string channelType = (string)((PropertyValueString)supportComponentBO.GetPropertyValue("IJUAHgrBlineChannelType", "ChannelType")).PropValue;

                PropertyValueCodelist holePatternCodeList = ((PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrBlineChannelProp", "HolePattern"));
                long holePattern = holePatternCodeList.PropValue;
                if (holePattern == -1)
                    holePattern = 1;
                if (holePattern <= 0 || holePattern > 10)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidChHolePattern, "Channel HolePattern Code list value should be between 1 and 10"));

                PropertyValueCodelist materialCodeList = ((PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrBlineChannelProp", "Material"));
                long material = (long)materialCodeList.PropValue;
                if (material <= 0 || material > 4)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrChMaterial, "Channel Material Code list value should be between 1 and 4"));

                
                string materialType = materialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)material).DisplayName;
                string holePatternValue = holePatternCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)holePattern).DisplayName;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass auxilaryTable = (PartClass)catalogBaseHelper.GetPartClass("BlineChWeightsAUX");
                ReadOnlyCollection<BusinessObject> classItems = auxilaryTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                double weightPerUnitlength = 0;
                foreach (BusinessObject classItem in classItems)
                {
                    if (classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "ChHolePattern").ToString() == holePatternValue && classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Material").ToString() == materialType && classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Type").ToString() == channelType)
                        weightPerUnitlength = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Weight_Per_Length")).PropValue;

                }
                weight = weightPerUnitlength * length + clampWeight;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of Bline_Trapeze_Supp_9P_55X."));
            }
        }
        #endregion
    }
}



