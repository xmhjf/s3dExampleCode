//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_CentHung_9ZN_52X.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_CentHung_9ZN_52X
//   Author       :  Shilpi	
//   Creation Date:  10.August.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10.August.2012  Shilpi     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   26.Nov.2012     Rajeswari  CR-CP-219113 Modified the code with SymbolGeomHelper
//   10.Jan.2010     Hema       CR-CP-219113 Changed the implementation of WeightCG
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
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Bline_CentHung_9ZN_52X : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_CentHung_9ZN_52X"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Gap", "Gap between Clamp and Web Rail", 0)]
        public InputDouble m_dGap;
        [InputDouble(3, "ClampLength", "Clamp Length", 0)]
        public InputDouble m_dClampLength;
        [InputDouble(4, "DistEdge", "Distance from edge to hole", 0)]
        public InputDouble m_dDistEdge;
        [InputDouble(5, "ClampWidth", "Clamp Width", 0)]
        public InputDouble m_dClampWidth;
        [InputDouble(6, "ClampHeight", "Clamp Height", 0)]
        public InputDouble m_dClampHeight;
        [InputDouble(7, "HoleDiameter", "Hole Diameter", 0)]
        public InputDouble m_dHoleDiameter;
        [InputDouble(8, "TrayWT", "Tray Web Thickness", 0)]
        public InputDouble m_dTrayWT;
        [InputDouble(9, "Inside_Outside", "Inside or Outside", 1)]
        public InputDouble m_dInside_Outside;
        [InputDouble(10, "Clamp_Guide", "Clamp or Guide", 1)]
        public InputDouble m_dClamp_Guide;
        [InputDouble(11, "WT", "Tray Width", 0)]
        public InputDouble m_dWT;
        [InputDouble(12, "HH", "Height of Hanger", 0)]
        public InputDouble m_dHH;
        [InputDouble(13, "LC", "Length of Coupling", 0)]
        public InputDouble m_dLC;
        [InputDouble(14, "DH", "Diameter of Hanger", 0)]
        public InputDouble m_dDH;
        [InputDouble(15, "CCL", "Channel Cut Length", 0)]
        public InputDouble m_dCCL;
        [InputDouble(16, "CT", "Thickness of Channel", 0)]
        public InputDouble m_dCT;
        [InputDouble(17, "ChWidth", "Width of Channel", 0)]
        public InputDouble m_dChWidth;
        [InputDouble(18, "CH", "Channel Height", 0)]
        public InputDouble m_dCH;
        [InputDouble(19, "Material", "Material", 0)]
        public InputDouble m_dMaterial;
        [InputDouble(20, "HolePattern", "Hole Pattern", 0)]
        public InputDouble m_dHolePattern;
        [InputDouble(21, "Thickness", "Thickness", 0)]
        public InputDouble m_dThickness;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("L_CLAMP", "L_CLAMP")]
        [SymbolOutput("R_CLAMP", "R_CLAMP")]
        [SymbolOutput("L_INTHRDRH", "L_INTHRDRH")]
        [SymbolOutput("R_INTHRDRH", "R_INTHRDRH")]
        [SymbolOutput("ROD_HANGER", "ROD_HANGER")]
        [SymbolOutput("CHANNEL", "CHANNEL")]

        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                double GA = m_dGap.Value;
                double CL = m_dClampLength.Value;
                double LS = m_dDistEdge.Value;
                double WC = m_dClampWidth.Value;
                double HC = m_dClampHeight.Value;
                double DO = m_dHoleDiameter.Value;
                double TW = m_dTrayWT.Value;
                double inOrOut = m_dInside_Outside.Value;
                double WT = m_dWT.Value;
                double HH = m_dHH.Value;
                double LC = m_dLC.Value;
                double DH = m_dDH.Value;
                double CCL = m_dCCL.Value;
                double CT = m_dCT.Value;
                double CW = m_dChWidth.Value;
                double CH = m_dCH.Value;

                PropertyValueCodelist insideOutsideCodelistValue = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");

                //Exceptions
                if (DH <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDH, "Diameter of Hanger should be greater than 0"));
                    return;
                }
                if (CCL == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidCCLZero, "Channel Cut Length value should equal to 0."));
                    return;
                }
                if (CCL > LC)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidCCL, "The Channel Cut Length should be smaller than the Length of the Channel. Please check the value."));
                if (DO <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDO, "Hole Diameter value should  be greater than 0."));
                if (CL <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidCL, "Clamp Length value should  be greater than 0."));
                if (WC <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidWCgtZero, "Width of Clamp value should  be greater than 0."));
                if (HC <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidHC, "Clamp Height value should  be greater than 0."));
                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //ports

                Port trayPort = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = trayPort;

                Port inThrdRH = new Port(OccurrenceConnection, part, "InThrdRH1", new Position(0, 0, HH - 0.008), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = inThrdRH;

                Vector normal = new Position(0, 0, HH).Subtract(new Position(0, 0, 0));
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                Projection3d rodHanger = symbolGeometryHelper.CreateCylinder(null, (DH) / 2, normal.Length);
                m_PhysicalAspect.Outputs["ROD_HANGER"] = rodHanger;

                if (insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem((int)inOrOut).DisplayName.ToLower() == "outside")
                {
                    Vector normal1 = new Position(0, 0.5 * (WT + TW) + 0.5 * TW + GA + LS, HC * 1.25).Subtract(new Position(0, 0.5 * (WT + TW) + 0.5 * TW + GA + LS, 0));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0.5 * (WT + TW) + 0.5 * TW + GA + LS, 0);
                    symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                    Projection3d lInthrfrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                    m_PhysicalAspect.Outputs["L_INTHRDRH"] = lInthrfrh;

                    Vector normal2 = new Position(0, -(0.5 * (WT + TW) + 0.5 * TW + GA + LS), HC * 1.25).Subtract(new Position(0, -(0.5 * (WT + TW) + 0.5 * TW + GA + LS), 0));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -(0.5 * (WT + TW) + 0.5 * TW + GA + LS), 0);
                    symbolGeometryHelper.SetOrientation(normal2, new Vector(1, 0, 0));
                    Projection3d rInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal2.Length);
                    m_PhysicalAspect.Outputs["R_INTHRDRH"] = rInthrdrh;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW + GA, HC / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                    Projection3d lClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL, WC, HC);
                    m_PhysicalAspect.Outputs["L_CLAMP"] = lClamp;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 + TW + GA), HC / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                    Projection3d rClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL, WC, HC);
                    m_PhysicalAspect.Outputs["R_CLAMP"] = rClamp;
                }
                else
                {
                    Vector normal1 = new Position(0, 0.5 * WT - GA - LS, HC * 1.25).Subtract(new Position(0, 0.5 * WT - GA - LS, 0));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0.5 * WT - GA - LS, 0);
                    symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                    Projection3d lInthrfrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                    m_PhysicalAspect.Outputs["L_INTHRDRH"] = rodHanger;

                    Vector normal2 = new Position(0, -0.5 * WT + GA + LS, HC * 1.25).Subtract(new Position(0, -0.5 * WT + GA + LS, 0));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -0.5 * WT + GA + LS, 0);
                    symbolGeometryHelper.SetOrientation(normal2, new Vector(1, 0, 0));
                    Projection3d rInthrdrh = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal2.Length);
                    m_PhysicalAspect.Outputs["R_INTHRDRH"] = rInthrdrh;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 - GA, HC / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                    Projection3d lClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL, WC, HC);
                    m_PhysicalAspect.Outputs["L_CLAMP"] = lClamp;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 - GA), HC / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                    Projection3d rClamp = (Projection3d)symbolGeometryHelper.CreateBox(null, CL, WC, HC);
                    m_PhysicalAspect.Outputs["R_CLAMP"] = rClamp;
                }

                //Add Extruded U-Shape for CHANNEL
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-CW / 2, -CCL / 2, 0));
                pointCollection.Add(new Position(-CW / 2 + CT, -CCL / 2, 0));
                pointCollection.Add(new Position(-CW / 2 + CT, -CCL / 2, -CH + CT));
                pointCollection.Add(new Position(CW / 2 - CT, -CCL / 2, -CH + CT));
                pointCollection.Add(new Position(CW / 2 - CT, -CCL / 2, 0));
                pointCollection.Add(new Position(CW / 2, -CCL / 2, 0));
                pointCollection.Add(new Position(CW / 2, -CCL / 2, -CH));
                pointCollection.Add(new Position(-CW / 2, -CCL / 2, -CH));
                pointCollection.Add(new Position(-CW / 2, -CCL / 2, 0));

                Vector projectionVector = new Vector(0, CCL, 0);
                Projection3d clampBody = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_PhysicalAspect.Outputs["CHANNEL"] = clampBody;
            }
            catch
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_CentHung_9ZN_52X."));
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
                PropertyValueCodelist inOrOutPropertyValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                PropertyValueCodelist clampOrGuidePropertyValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineClampGuide", "Clamp_Guide");
                PropertyValueDouble cutLengthPropertyValue = (PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineChannelCutL", "CCL");
                double cutLengthValue = cutLengthPropertyValue.PropValue.Value;
                if (inOrOutPropertyValue.PropValue <= 0 || inOrOutPropertyValue.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInside_Outside, "Bline Inside_Outside Code list value should be between 1 and 2"));
                if (clampOrGuidePropertyValue.PropValue <= 0 || clampOrGuidePropertyValue.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidClampGuide, "Clamp Guide Code list value should be between 1 and 2"));

                string inOrOut = inOrOutPropertyValue.PropertyInfo.CodeListInfo.GetCodelistItem(inOrOutPropertyValue.PropValue).DisplayName;
                string clampOrGuide = clampOrGuidePropertyValue.PropertyInfo.CodeListInfo.GetCodelistItem(clampOrGuidePropertyValue.PropValue).DisplayName;
                string cutLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, cutLengthValue, UnitName.DISTANCE_INCH);

                Part centHung = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                bomDescription = centHung.PartDescription + ", " + clampOrGuide + ", Installed " + inOrOut + "," + " Channel Cut Length: " + cutLength;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_CentHung_9ZN_52X."));
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
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrBlineChannelCutL", "CCL")).PropValue;
                double clampWeight = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrBlineCHHWWeight", "CentHung_HW_Weight")).PropValue;
                string channelType = (string)((PropertyValueString)supportComponentBO.GetPropertyValue("IJUAHgrBlineChannelType", "ChannelType")).PropValue;

                PropertyValueCodelist holePatternCodeList = ((PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrBlineChannelProp", "HolePattern"));
                long holePattern = holePatternCodeList.PropValue;
                if (holePattern == -1)
                    holePattern = 1;
                if (holePattern <= 0 || holePattern > 10)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidChHolePattern, "Channel HolePattern Code list value should be between 1 and 10"));
                }

                PropertyValueCodelist materialCodeList = ((PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrBlineChannelProp", "Material"));
                long material = (long)materialCodeList.PropValue;
                if (material <= 0 || material > 4)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrChMaterial, "Channel Material Code list value should be between 1 and 4"));
                }


                string materialType = materialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)material).DisplayName;
                string holePatternValue = holePatternCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)holePattern).DisplayName;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass auxilaryTable = (PartClass)catalogBaseHelper.GetPartClass("BlineChWeightsAUX");
                ReadOnlyCollection<BusinessObject> classItems = auxilaryTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                double weightPerUnitlength = 0;
                foreach (BusinessObject classItem in classItems)
                {
                    if (classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "ChHolePattern").ToString() == holePatternValue && classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Material").ToString() == materialType && classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Type").ToString() == channelType)
                    {
                        weightPerUnitlength = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrBlineChWeightAUX", "Weight_Per_Length")).PropValue;

                    }
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of Bline_CentHung_9ZN_52X."));
            }
        }

        #endregion
    }
}




